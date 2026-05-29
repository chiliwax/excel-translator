from __future__ import annotations

import json
import time
from dataclasses import dataclass
from typing import Any, Protocol, Sequence


class TranslationError(RuntimeError):
    """Raised when a translation provider fails or returns invalid data."""


class TranslationClient(Protocol):
    def translate_batch(
        self,
        texts: Sequence[str],
        *,
        source_language: str,
        target_language: str,
    ) -> list[str]:
        """Translate texts in order and return one translated string per input."""


@dataclass(frozen=True)
class LiteLLMTranslationClient:
    model: str = "openai/gpt-4o-mini"
    temperature: float = 0.0
    timeout_seconds: float = 120.0
    retries: int = 2
    retry_sleep_seconds: float = 1.0

    def translate_batch(
        self,
        texts: Sequence[str],
        *,
        source_language: str,
        target_language: str,
    ) -> list[str]:
        if not texts:
            return []

        try:
            from litellm import completion
        except ImportError as exc:  # pragma: no cover - depends on runtime environment
            raise TranslationError(
                "LiteLLM is not installed. Install the project dependencies first."
            ) from exc

        messages = _build_translation_messages(
            texts=texts,
            source_language=source_language,
            target_language=target_language,
        )

        last_error: Exception | None = None
        for attempt in range(self.retries + 1):
            try:
                response = completion(
                    model=self.model,
                    messages=messages,
                    temperature=self.temperature,
                    timeout=self.timeout_seconds,
                )
                content = _response_content(response)
                translations = _parse_translations(content)
                if len(translations) != len(texts):
                    raise TranslationError(
                        f"Provider returned {len(translations)} translations for "
                        f"{len(texts)} input strings."
                    )
                return translations
            except Exception as exc:  # pragma: no cover - provider behavior varies
                last_error = exc
                if attempt >= self.retries:
                    break
                time.sleep(self.retry_sleep_seconds * (2**attempt))

        raise TranslationError(f"Translation failed: {last_error}") from last_error


def _build_translation_messages(
    *,
    texts: Sequence[str],
    source_language: str,
    target_language: str,
) -> list[dict[str, str]]:
    if source_language.lower() == "auto":
        source_instruction = "Auto-detect the source language for each item."
    else:
        source_instruction = f"The source language is {source_language}."

    payload = {"texts": list(texts)}
    return [
        {
            "role": "system",
            "content": (
                "You are a precise translation engine for spreadsheet cells. "
                "Translate each input string to the requested target language. "
                "Preserve leading/trailing whitespace, line breaks, placeholders, "
                "HTML/XML tags, URLs, numbers, and product names unless they must be "
                "translated naturally. Return only valid JSON."
            ),
        },
        {
            "role": "user",
            "content": (
                f"{source_instruction}\n"
                f"Target language: {target_language}\n"
                "Return JSON with exactly this shape: "
                '{"translations": ["translated text", "..."]}.\n'
                "Keep the output array in the exact same order and length as the input.\n"
                f"Input JSON:\n{json.dumps(payload, ensure_ascii=False)}"
            ),
        },
    ]


def _response_content(response: Any) -> str:
    choices = getattr(response, "choices", None)
    if choices is None and isinstance(response, dict):
        choices = response.get("choices")
    if not choices:
        raise TranslationError("Provider response did not include choices.")

    first_choice = choices[0]
    message = getattr(first_choice, "message", None)
    if message is None and isinstance(first_choice, dict):
        message = first_choice.get("message")

    content = getattr(message, "content", None)
    if content is None and isinstance(message, dict):
        content = message.get("content")
    if not isinstance(content, str) or not content.strip():
        raise TranslationError("Provider response did not include text content.")
    return content


def _parse_translations(content: str) -> list[str]:
    try:
        data = json.loads(content)
    except json.JSONDecodeError:
        data = json.loads(_extract_json(content))

    if isinstance(data, dict):
        translations = data.get("translations")
    else:
        translations = data

    if not isinstance(translations, list):
        raise TranslationError("Provider response JSON must contain a translations list.")
    if not all(isinstance(item, str) for item in translations):
        raise TranslationError("Every translation must be a string.")
    return translations


def _extract_json(content: str) -> str:
    object_start = content.find("{")
    object_end = content.rfind("}")
    if object_start >= 0 and object_end > object_start:
        return content[object_start : object_end + 1]

    array_start = content.find("[")
    array_end = content.rfind("]")
    if array_start >= 0 and array_end > array_start:
        return content[array_start : array_end + 1]

    raise TranslationError("Provider response did not contain JSON.")
