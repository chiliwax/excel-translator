from __future__ import annotations

import os

from xlsx_ai_translate.cli import _load_env_file, _parse_env_file_arg, build_parser


def test_load_env_file_sets_values_without_overriding_existing_env(tmp_path, monkeypatch):
    env_file = tmp_path / ".env"
    env_file.write_text(
        "\n".join(
            [
                "# comment",
                "OPENAI_API_KEY=from-file",
                "export XLSX_TRANSLATOR_MODEL=openai/gpt-4o-mini",
                "QUOTED='quoted value'",
            ]
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv("OPENAI_API_KEY", "from-shell")
    monkeypatch.delenv("XLSX_TRANSLATOR_MODEL", raising=False)
    monkeypatch.delenv("QUOTED", raising=False)

    _load_env_file(env_file)

    assert os.environ["OPENAI_API_KEY"] == "from-shell"
    assert os.environ["XLSX_TRANSLATOR_MODEL"] == "openai/gpt-4o-mini"
    assert os.environ["QUOTED"] == "quoted value"


def test_load_env_file_ignores_missing_file(tmp_path):
    _load_env_file(tmp_path / ".env")


def test_parse_env_file_arg_defaults_to_dotenv():
    assert _parse_env_file_arg([]).name == ".env"


def test_parse_env_file_arg_accepts_custom_file():
    assert _parse_env_file_arg(["--env-file", "secrets.env"]).name == "secrets.env"


def test_parser_uses_aggressive_openai_rate_defaults():
    args = build_parser().parse_args(["input.xlsx", "output.xlsx", "--target", "en"])

    assert args.concurrency == 8
    assert args.rpm == 500
    assert args.tpm == 200_000
