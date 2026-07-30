# Automotive OEM Validation Automation

Browser automation pipeline that validates automotive part reference codes (OEM/parallel) against a live parts catalog, replacing a manual, spreadsheet-driven lookup process with a resumable batch job.

## Problem it solves

Dealership and parts-recambista inventories are tracked in Excel, with each row holding an "original" (OEM) reference and a "parallel" (aftermarket) reference for a given part. Confirming that each reference actually returns hits on the vendor's catalog was previously done by hand, one lookup at a time. This project automates that OEM reference validation end-to-end: it drives a real browser against the catalog site, counts matching listings per reference, and writes the result back into the spreadsheet — with zero-result cases left visible in the output so they can be routed to manual review (Human-in-the-Loop) instead of silently accepted.

Runs are batch-oriented and interruption-safe: progress is written as structured, line-by-line JSONL, so a run stopped mid-batch can resume from the exact last row instead of restarting.

## Install

```powershell
python -m venv .venv
& ".\.venv\Scripts\Activate.ps1"
pip install -r verificacion_oem/requirements.txt
playwright install chromium
```

## Usage

Run the CLI against a folder containing a single input `.xlsx` (columns: `Pieza`, `Ref. Original (Concesionarios)`, `Ref. Paralelo (Recambistas)`):

```powershell
python -m verificacion_oem.cli --folder "C:\Users\<usuario>\OneDrive\Desktop\Verificacion OEM" --batch 500 --start 0 --autosave 50 --delay 0.5 --max-pages 20 --per-page 30
```

If `--folder` is omitted, the CLI looks for a `Verificacion OEM` folder on the Desktop (including OneDrive) before falling back to the current directory.

Key flags:

- `--batch` / `--start` — how many rows to process and where to start (0-indexed).
- `--autosave` — writes a partial `.xlsx` every N rows (`0` disables it).
- `--resume` — replays `run_artifacts/run_log.jsonl` to skip already-processed rows and continue where the last run left off (JSONL-based resume).
- `--headful` — shows the browser window (headless by default).
- `--max-pages` / `--per-page` — pagination limits per search; the scraper also stops early once a row's match count crosses the configured early-stop threshold (`max_count` in `CounterConfig`), since only the count relative to that threshold matters, not the exact total.
- `--proxy`, `--timeout`, `--require-country` — network and connection controls.

Run the test suite:

```powershell
python -m pytest -q
```

### Output

- Partial saves: `<input>_validated_partial_<start>_<end>.xlsx`.
- Final result: `<input>_validated.xlsx`, with `Validacion Original` / `Validacion Paralelo` columns filled in.
- Per-row structured logging: `run_artifacts/run_log.jsonl` (one JSON object per row — status, counts, timing, cache hits, errors).
- Run summary: `run_artifacts/summary.json` (totals, throughput, error/timeout counts).

## Project structure

```
verificacion_oem/       # Maintained package
  cli.py                 # Argument parsing and CLI entry point
  config.py               # AppConfig / CounterConfig (columns, batch size, early-stop threshold, etc.)
  excel_handler.py         # Excel discovery, reading, validation, saving
  query_builder.py         # Search-query construction and filtering rules
  scraper.py                # Playwright browser automation against the catalog site
  processor.py               # Orchestrates rows -> queries -> counts -> logging -> autosave
  logger.py                   # JSONL structured logging, resume, and run summaries
tests/                   # pytest test suite
archive/                # Superseded scraper implementations, kept for reference only
oem_verifier/, oem_verify_OK.py  # Legacy pre-refactor scripts, no longer maintained
```

## Stack

Python · Pandas · Playwright · OpenPyXL · Excel Automation · Browser Automation · JSONL Logging · CLI · Batch Processing

## Note

This is a portfolio version of the project. It contains no real inventory data, no production `.xlsx` files, and no `run_artifacts/` output — those are excluded via `.gitignore` and generated locally when the tool is run.
