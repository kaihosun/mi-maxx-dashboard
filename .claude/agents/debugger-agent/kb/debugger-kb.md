# Debugger KB — MI Technologies MTY MAXX

## Common Error Catalog

### 1. UnicodeDecodeError
- Trigger: CSV read on files with Spanish accents (á, é, ñ, etc.)
- Fix: `open(file, encoding='utf-8-sig')` or `pd.read_csv(file, encoding='utf-8-sig')`
- Files most affected: Almacén, Logística, Clasificación, Auditoría CSVs

### 2. KeyError in TRANSLATIONS
- Trigger: Spanish cell text not in TRANSLATIONS dict
- Fix: Add the missing key-value pair to TRANSLATIONS in `build_activity_log.py`
- Never remove the lookup — find the missing key

### 3. openpyxl MergedCell write error
- Trigger: Attempting to write to a cell that is part of a merged region
- Fix: Call `ws.unmerge_cells(range)` before writing, then re-merge after
- Check: `ws.merged_cells` to see all merged ranges

### 4. AttributeError: NoneType / None cell value
- Trigger: CSV row has empty cell, written directly to Excel
- Fix: Use `value or ""` or `str(value) if value else ""`

### 5. FileNotFoundError on CSV
- Trigger: CSV filename has accented characters not matching filesystem
- Fix: Use `os.listdir()` to find exact filename, then open with that exact name

### 6. openpyxl column width TypeError
- Trigger: Setting `column_dimensions[letter].width = None`
- Fix: Always compute a numeric width: `max(len(str(cell.value or "")) for cell in column)`

### 7. Sheet count mismatch
- Trigger: SHEET_NAMES has 15 entries but output has fewer sheets
- Diagnosis: Check which CSV file is missing or failed to read
- Fix: Wrap each department's CSV read in try/except and log which one fails

## Debug Flow
```
1. Run: python3 build_activity_log.py 2>&1
2. Read full traceback — identify file:line:error
3. Categorize error using catalog above
4. Apply minimum fix
5. Re-run: python3 build_activity_log.py
6. Confirm: output files exist and sheet count == 15
```

## Runtime Info
- OS: macOS (darwin)
- Shell: zsh
- Python: python3
- Key library: openpyxl (no pandas required for basic builds)

## Skills Available
- `debug-python` — systematic Python error diagnosis and fix
- `trace-data-error` — traces a data anomaly in a CSV back to its source
