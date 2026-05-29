from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Sequence

from .ai import LiteLLMTranslationClient
from .workbook import TranslationStats, stderr_progress, translate_workbook


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="translate-xlsx",
        description="Translate string cells in a multi-sheet XLSX file with an AI provider.",
    )
    parser.add_argument("input", type=Path, help="Input .xlsx file")
    parser.add_argument("output", type=Path, help="Output .xlsx file")
    parser.add_argument(
        "--env-file",
        type=Path,
        default=Path(".env"),
        help="Environment file containing API keys. Defaults to .env.",
    )
    parser.add_argument(
        "--target",
        required=True,
        help="Target language, for example 'en', 'fr', or 'English'.",
    )
    parser.add_argument(
        "--source",
        default="auto",
        help="Source language. Defaults to auto-detect per cell.",
    )
    parser.add_argument(
        "--model",
        default=os.environ.get("XLSX_TRANSLATOR_MODEL", "openai/gpt-4o-mini"),
        help="LiteLLM model name. Defaults to XLSX_TRANSLATOR_MODEL or openai/gpt-4o-mini.",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=50,
        help="Number of unique strings to translate per AI request.",
    )
    parser.add_argument(
        "--max-batch-chars",
        type=int,
        default=12_000,
        help="Maximum source characters per AI request before starting a new batch.",
    )
    parser.add_argument(
        "--concurrency",
        type=int,
        default=8,
        help="Maximum number of translation requests to run at once.",
    )
    parser.add_argument(
        "--rpm",
        type=int,
        default=500,
        help="Requests-per-minute budget for local throttling.",
    )
    parser.add_argument(
        "--tpm",
        type=int,
        default=200_000,
        help="Tokens-per-minute budget for local throttling.",
    )
    parser.add_argument(
        "--exclude-sheet",
        "--exlude-sheet",
        action="append",
        default=[],
        help="Sheet name to skip. Can be passed multiple times.",
    )
    parser.add_argument(
        "--temperature",
        type=float,
        default=0.0,
        help="AI sampling temperature. Keep 0 for stable translations.",
    )
    parser.add_argument(
        "--request-timeout",
        type=float,
        default=120.0,
        help="Per-request AI timeout in seconds.",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Do not print per-batch progress.",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    argv = list(argv) if argv is not None else sys.argv[1:]
    env_file = _parse_env_file_arg(argv)
    try:
        _load_env_file(env_file)
    except Exception as exc:
        print(f"translate-xlsx: error: {exc}", file=sys.stderr)
        return 1

    parser = build_parser()
    args = parser.parse_args(argv)

    client = LiteLLMTranslationClient(
        model=args.model,
        temperature=args.temperature,
        timeout_seconds=args.request_timeout,
    )

    try:
        stats = translate_workbook(
            args.input,
            args.output,
            client=client,
            source_language=args.source,
            target_language=args.target,
            exclude_sheets=args.exclude_sheet,
            batch_size=args.batch_size,
            max_batch_chars=args.max_batch_chars,
            concurrency=args.concurrency,
            requests_per_minute=args.rpm,
            tokens_per_minute=args.tpm,
            progress=None if args.quiet else stderr_progress,
        )
    except Exception as exc:
        print(f"translate-xlsx: error: {exc}", file=sys.stderr)
        return 1

    print(_format_stats(stats, output=args.output))
    return 0


def _parse_env_file_arg(argv: Sequence[str]) -> Path:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--env-file", type=Path, default=Path(".env"))
    args, _ = parser.parse_known_args(argv)
    return args.env_file


def _load_env_file(path: Path) -> None:
    if not path.exists():
        return
    if not path.is_file():
        raise ValueError(f"env file is not a file: {path}")

    for line_number, raw_line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export ") :].strip()
        if "=" not in line:
            raise ValueError(f"invalid env assignment in {path}:{line_number}")

        key, value = line.split("=", 1)
        key = key.strip()
        if not key:
            raise ValueError(f"empty env key in {path}:{line_number}")

        os.environ.setdefault(key, _strip_env_value(value))


def _strip_env_value(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
        return value[1:-1]
    return value


def _format_stats(stats: TranslationStats, *, output: Path) -> str:
    lines = [
        f"Saved translated workbook: {output}",
        f"Sheets translated: {stats.sheets_translated}/{stats.sheets_seen}",
        f"String cells translated: {stats.string_cells_found}",
        f"Unique strings sent to AI: {stats.unique_strings}",
        f"Duplicate cells reused from cache: {stats.duplicate_strings_reused}",
        f"Formula cells preserved: {stats.formula_cells_skipped}",
    ]
    if stats.excluded_sheets:
        lines.append(f"Excluded sheets: {', '.join(stats.excluded_sheets)}")
    if stats.missing_excluded_sheets:
        lines.append(
            f"Exclude-sheet names not found: {', '.join(stats.missing_excluded_sheets)}"
        )
    return "\n".join(lines)


if __name__ == "__main__":
    raise SystemExit(main())
