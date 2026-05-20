# CSV to XLSX Converter

Native macOS desktop app for converting large, RFC 4180–compliant CSV files
(500k+ rows, 670 MB+) to XLSX without the column corruption you get when
Excel or Google Sheets tries to import them.

All columns are treated as **text** — UUIDs, phone numbers, leading-zero
zip codes, and income ranges like `$20,000 to $44,999` are preserved
exactly as written.

## What it does

- Drag & drop or browse to a `.csv` file
- Shows file size, row count, and a 5-row preview
- Reads with Polars (`infer_schema_length=0`), falls back to pandas if needed
- Writes XLSX with `xlsxwriter` in constant-memory mode
- Splits across `Data_1`, `Data_2`, … sheets if the row count exceeds
  Excel's hard limit of 1,048,576 rows per sheet
- Runs the conversion on a background thread so the UI stays responsive
- Logs parsing warnings to a collapsible panel without failing the conversion

## Install (Mac)

You need Python 3.10 or newer. Check with `python3 --version`. If you
don't have it, install via [python.org](https://www.python.org/downloads/macos/)
or `brew install python`.

```bash
cd ~/csv_converter
python3 -m venv .venv
source .venv/bin/activate
pip3 install --upgrade pip
pip3 install -r requirements.txt
```

## Run

Easy mode — double-click `run.sh` in Finder, or:

```bash
cd ~/csv_converter
./run.sh
```

`run.sh` creates the `.venv` on first run, installs dependencies, then
launches the app. Later runs just activate and launch.

If double-clicking `.sh` in Finder opens a text editor instead of running
it, right-click → **Open With → Terminal**, or run
`chmod +x run.sh` once and then set Finder to open `.sh` with Terminal.

## Bonus: package as a standalone .app

If you'd rather not touch Terminal at all, build a self-contained
`.app` bundle with PyInstaller:

```bash
cd ~/csv_converter
source .venv/bin/activate
pip3 install pyinstaller
pyinstaller --windowed --name "CSV to XLSX Converter" \
    --osx-bundle-identifier com.legenex.csvconverter \
    main.py
```

The bundle lands in `dist/CSV to XLSX Converter.app`. Drag it to
`/Applications`. First launch may need a right-click → **Open** to
get past Gatekeeper, since the app isn't code-signed.

To rebuild after editing `main.py`, delete `build/` and `dist/` then
re-run the `pyinstaller` command.

## Performance notes

- Target: ~500k rows / 670 MB in under 2 minutes on Apple Silicon
- Polars reads the whole CSV into memory once (typically 1–3 GB for a
  670 MB file expanded to strings); xlsxwriter writes incrementally
  to disk in constant-memory mode
- If memory is tight, close other apps before converting. The app catches
  `MemoryError` and surfaces a clear suggestion if it happens

## Troubleshooting

- **"Polars path failed, falling back to pandas"** — benign; the pandas
  fallback handles the same data, just slower
- **Conversion fails with a parsing error** — the error dialog includes
  the line number Polars or pandas choked on. Open the CSV in Modern CSV,
  jump to that line, and check for an unterminated quote
- **App won't launch** — run `./run.sh` from Terminal and read the error.
  The most common cause is a stale `.venv` after a Python upgrade — just
  `rm -rf .venv` and re-run
