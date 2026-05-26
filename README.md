# Document Intelligence Pipeline

Three-stage pipeline that turns PDFs into a structured Excel workbook using an Azure Document Intelligence custom model.

```
input/*.pdf  ──►  intel_output/*_raw.json  ──►  clean_labeled_data/*_clean.json  ──►  excel_output/licenses.xlsx
                  Stage 1                       Stage 2                              Stage 3
```

Every run also writes a manifest to `runs/<timestamp>.json` capturing per-file status, durations, and errors.

## Setup

```bash
pip install -r requirements.txt
cp env.example .env
# edit .env with your Azure endpoint, key, and custom model id
mkdir input
# drop your PDFs into input/
```

## Usage

```bash
python docai_pipeline.py                  # run all three stages
python docai_pipeline.py --stage 1        # just upload + analyze
python docai_pipeline.py --stage 2        # re-clean existing raw JSON
python docai_pipeline.py --stage 3        # regenerate the Excel from clean JSON
python docai_pipeline.py --force          # re-analyze PDFs even if raw JSON exists
```

Stage 1 skips any PDF whose `*_raw.json` already exists, so re-running is cheap. Use `--force` to override.

## Stages

**Stage 1 — Analyze.** Reads each PDF in `input/`, sends it to your Azure custom model in parallel (default 4 workers), and writes the full response to `intel_output/<name>_raw.json`. Retries on 429/5xx with exponential backoff.

**Stage 2 — Clean & label.** Walks the raw Azure response, extracts each labeled field with its value, type, content, and confidence, and writes a flattened `clean_labeled_data/<name>_clean.json`. Handles strings, numbers, dates, currency, addresses, arrays, and nested objects.

**Stage 3 — Aggregate to Excel.** Builds `excel_output/licenses.xlsx`:
- `main` sheet — one row per detected document, one column per simple field plus a `<field>__conf` column with confidence.
- One additional sheet per array field (e.g. line items), joined back to the main sheet via the `id_value` column (configured by `DOCAI_ID_FIELD`).

## Configuration

All settings live in `.env`. Required:

| Variable | Description |
|---|---|
| `AZURE_DI_ENDPOINT` | Your Azure Document Intelligence resource endpoint |
| `AZURE_DI_KEY` | API key |
| `AZURE_DI_MODEL_ID` | Custom model id |

Optional (defaults shown):

| Variable | Default | Purpose |
|---|---|---|
| `DOCAI_INPUT_DIR` | `./input` | Where PDFs are read from |
| `DOCAI_INTEL_OUTPUT_DIR` | `./intel_output` | Raw Azure responses |
| `DOCAI_CLEAN_LABELED_DIR` | `./clean_labeled_data` | Cleaned per-document JSON |
| `DOCAI_EXCEL_OUTPUT_DIR` | `./excel_output` | Final workbook |
| `DOCAI_EXCEL_FILENAME` | `licenses.xlsx` | Workbook filename |
| `DOCAI_RUNS_DIR` | `./runs` | Per-run manifests |
| `DOCAI_MAX_WORKERS` | `4` | Concurrent Stage 1 requests |
| `DOCAI_MAX_RETRIES` | `3` | Retry attempts on transient Azure errors |
| `DOCAI_MAX_FILE_SIZE_MB` | `500` | Skip PDFs larger than this |
| `DOCAI_ID_FIELD` | `License Number` | Field used to join array sheets to main rows |

`DOCAI_ID_FIELD` must match the exact label in your custom model.

## Run manifest

Each invocation produces `runs/<UTC-timestamp>.json` with:

```json
{
  "run_id": "20260526T143012Z",
  "started_at": "...", "finished_at": "...", "duration_s": 42.1,
  "args": { "stage": "all", "force": false },
  "config": { "model_id": "...", "max_workers": 4, ... },
  "stage_1": {
    "workers": 4,
    "counts": { "ok": 12, "skipped": 3, "failed": 1, "too_large": 0 },
    "files": [
      { "source": "doc.pdf", "size_bytes": 1234567, "status": "ok",
        "duration_s": 4.2, "raw_path": "...", "error": null }
    ]
  },
  "stage_2": { "files": [ ... ] },
  "stage_3": { "workbook": "...", "main_rows": 12, "array_sheets": { ... } }
}
```

Useful for spotting failures, comparing run timings, or auditing what changed between runs.

## Requirements

- Python 3.10+ (uses `list[Path]` syntax)
- An Azure Document Intelligence resource with a trained custom model

## Tuning parallelism

`DOCAI_MAX_WORKERS=4` is conservative. You can raise it (8–16) if you have a higher-tier Azure resource, but watch for 429s in the manifest — Azure tiers have per-second and per-minute caps. The retry logic handles occasional throttling, but sustained overrun will slow runs down rather than speed them up.
