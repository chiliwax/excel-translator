from __future__ import annotations

from xlsx_ai_translate.ai import _parse_translations


def test_parse_translations_from_json_object():
    assert _parse_translations('{"translations": ["Bonjour", "Au revoir"]}') == [
        "Bonjour",
        "Au revoir",
    ]


def test_parse_translations_from_wrapped_json():
    content = 'Here is the JSON: {"translations": ["Hola"]}'

    assert _parse_translations(content) == ["Hola"]
