# XLSX AI Translate

Translate string cells in `.xlsx` workbooks with an AI provider while preserving workbook structure, formulas, and repeated text consistency.

## Features

- Reads and writes `.xlsx` files.
- Translates all sheets by default.
- Supports `--exclude-sheet` for sheets that must remain unchanged.
- Translates worksheet tab names by default.
- Preserves formulas, numbers, dates, booleans, blank cells, and whitespace-only cells.
- Translates exact duplicate strings only once per file, then reuses the translated value.
- Uses LiteLLM so OpenAI works by default and other providers can be plugged in with a model name.

## Install

```bash
pip install -e .
```

For development:

```bash
pip install -e '.[dev]'
```

## Usage

Create a local `.env` file for your API key:

```bash
cp .env.example .env
```

Then edit `.env`:

```env
OPENAI_API_KEY=your-openai-api-key
XLSX_TRANSLATOR_MODEL=openai/gpt-4o-mini
```

Run the translator:

```bash
translate-xlsx input.xlsx output.xlsx --target en
```

The CLI loads `.env` by default. Use a different file with `--env-file`:

```bash
translate-xlsx input.xlsx output.xlsx --target en --env-file ./secrets.env
```

Set the source language manually when auto-detection is not desired:

```bash
translate-xlsx input.xlsx output.xlsx --source fr --target en
```

Exclude one or more sheets:

```bash
translate-xlsx input.xlsx output.xlsx --target en --exclude-sheet "Metadata" --exclude-sheet "Do Not Translate"
```

Worksheet tab names are translated by default. Disable that when needed:

```bash
translate-xlsx input.xlsx output.xlsx --target en --no-translate-sheet-names
```

Large workbooks are split into batches by both item count and source character count:

```bash
translate-xlsx input.xlsx output.xlsx --target en --batch-size 50 --max-batch-chars 12000
```

Batches run in parallel by default using your `gpt-4o-mini` rate limits:

```bash
translate-xlsx input.xlsx output.xlsx --target en --concurrency 8 --rpm 500 --tpm 200000
```

Lower `--concurrency` if your provider returns rate-limit errors. Set `--quiet` to hide per-batch progress.

Choose a different LiteLLM model/provider:

```bash
translate-xlsx input.xlsx output.xlsx --target en --model openai/gpt-4o-mini
translate-xlsx input.xlsx output.xlsx --target en --model anthropic/claude-3-5-sonnet-latest
translate-xlsx input.xlsx output.xlsx --target en --model azure/my-deployment
```

You can also set the default model with:

```bash
export XLSX_TRANSLATOR_MODEL="openai/gpt-4o-mini"
```

Environment variables already set in your shell take priority over values in `.env`.

## Behavior

- `--source` defaults to `auto`, which asks the model to detect language per cell.
- `--target` is required.
- Duplicate detection is exact: `Hello`, `hello`, and `Hello ` are different strings.
- The duplicate cache is per workbook run only.
- Formula cells are loaded with `data_only=False` and left unchanged.
- Formula references to renamed sheets are updated so sheet links keep working.
- Translated sheet names are sanitized for Excel's title rules and made unique if necessary.
- Default throttling is `--concurrency 8 --rpm 500 --tpm 200000`.

## Tests

```bash
pytest
```

## Web UI

Run the FastAPI web interface locally:

```bash
pip install -e '.[dev]'
uvicorn web.app:app --reload --host 127.0.0.1 --port 8000
```

Open `http://127.0.0.1:8000`, upload an `.xlsx`, and download the translated result when the job finishes.

The web app stores job files under `.data/jobs` by default. Override that with:

```bash
export XLSX_TRANSLATOR_DATA_DIR=/path/to/data
```

Deployment examples are available in `deploy/excel-translator.service` and `deploy/nginx.conf`.
