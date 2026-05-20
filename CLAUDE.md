# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Single-file PyQt6 desktop app (`main.py`) that converts large RFC 4180 CSV files (500k+ rows, 670 MB+) to XLSX without the type-coercion corruption Excel/Sheets inflict on UUIDs, phone numbers, leading-zero zips, and money-range strings. Target machine is Apple Silicon macOS.

## Commands

```bash
./run.sh                  # creates .venv on first run, installs deps, launches app
source .venv/bin/activate # then `python main.py` for normal dev iteration

# Package as a standalone .app (see README for full incantation):
pyinstaller --windowed --name "CSV to XLSX Converter" \
    --osx-bundle-identifier com.legenex.csvconverter main.py
# After editing main.py, delete build/ and dist/ before rebuilding.
```

There is no test suite, linter, or formatter configured. Don't fabricate one.

## Architecture

Three threads of execution, all wired through Qt signals:

1. **Main (UI) thread** — `MainWindow` owns all widgets; `DropZone` accepts drag/drop and click-to-browse. Never block here.
2. **`RowCountWorker` (QThread)** — fires on file select to compute row count + 5-row preview. Polars lazy frame (`pl.scan_csv` + `pl.len()`) avoids materializing data; pandas fallback uses a cheap byte-level newline count (approximate when CSVs contain quoted newlines — this is acceptable because the count only drives the progress bar).
3. **`ConvertWorker` (QThread)** — does the actual conversion. Tries Polars first, falls back to pandas chunked (50k rows) on any non-`MemoryError` exception. Always re-raises `MemoryError` so the UI can show a clear OOM message.

### Two invariants that everything depends on

- **All cells stay text.** Polars: `infer_schema_length=0`, `null_values=[]`, `try_parse_dates=False`. Pandas: `dtype=str`, `na_filter=False`, `keep_default_na=False`. xlsxwriter: `strings_to_numbers=False`, `strings_to_formulas=False`, `strings_to_urls=False`, and writes via `write_string` only. Breaking any of these reintroduces the silent corruption the app exists to prevent.
- **Constant-memory write path.** `xlsxwriter.Workbook(..., {"constant_memory": True})` means only the current row is in RAM, prior rows are flushed to disk. Don't add post-write formatting passes, autofilters, or anything that requires re-reading earlier rows — constant_memory forbids that and a 500k-row workbook won't fit otherwise.

### Other things worth knowing before editing

- Excel's hard per-sheet limit is 1,048,576 rows. `_write_xlsx` splits across `Data_1`, `Data_2`, … sheets when exceeded; the header is rewritten on each new sheet. `EXCEL_MAX_ROWS` constant captures this.
- Empty/None cells are skipped (left blank) rather than written — otherwise the strings `"NaN"` / `"None"` would appear.
- Progress is throttled via `emit_every = max(1, min(5000, total_rows // 100 or 1))` to keep the UI responsive without flooding the event loop.
- Polars and pandas are both probed at import time (`POLARS_AVAILABLE` flag) so the worker thread can branch without re-paying import cost per conversion.
