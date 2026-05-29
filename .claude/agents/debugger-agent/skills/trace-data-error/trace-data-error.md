# Skill: Trace Data Error

## Purpose
Trace a data anomaly in an Excel output or CSV back to its origin — identify which row, column, and source file caused the problem.

## When to Use
- A cell in the output Excel has wrong text, wrong encoding, or is blank when it shouldn't be
- A department is missing from the output workbook
- A translation appears wrong or untranslated in the output
- User reports "the FFT sheet shows garbage characters" or "Almacén is missing a row"

## Steps
1. Identify the symptom: which sheet, which cell range, what's wrong
2. Map sheet → department → CSV using SHEET_NAMES in `build_activity_log.py`
3. Open the source CSV: `cat "<csv_file>" | head -20` to inspect raw data
4. Check encoding: `file "<csv_file>"` — note if UTF-8 or ISO-8859-1
5. Search for the problematic value in the CSV: `grep -n "<value>" "<csv_file>"`
6. Check if the value has a TRANSLATIONS entry: search `build_activity_log.py` for the key
7. If encoding artifact: re-read with `utf-8-sig` or `latin-1` and compare outputs
8. If missing row: compare CSV row count vs Excel row count
9. If wrong translation: update TRANSLATIONS dict with correct mapping
10. Re-run build and verify fix in output file

## Diagnostic Commands
```bash
# Check CSV encoding
file "Registro Actividades (Personal Admin) - Almacén.csv"

# Count rows
wc -l "Registro Actividades (Personal Admin) - Almacén.csv"

# Search for value
grep -n "texto_buscado" "Registro Actividades (Personal Admin) - Almacén.csv"

# Check for encoding artifacts
python3 -c "
with open('file.csv', encoding='utf-8-sig') as f:
    for i, line in enumerate(f):
        if '\\ufffd' in line:
            print(f'Artifact at line {i}: {line[:80]}')
"
```

## Report Format
```
DATA TRACE REPORT
Symptom: <description>
Sheet: <sheet name> → Department: <department> → CSV: <filename>
Root cause: <encoding issue / missing key / wrong CSV / etc.>
Source location: <file>:<line>
Fix applied: <what was changed>
Verification: <before/after cell value or row count>
```
