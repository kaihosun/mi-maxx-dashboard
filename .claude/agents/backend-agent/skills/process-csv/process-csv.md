# Skill: Process CSV

## Purpose
Read, validate, and prepare a department CSV for ingestion into the Activity Log build pipeline.

## When to Use
- Before running `build_activity_log.py` on a new/updated CSV
- When a CSV has unknown structure and needs schema discovery
- When validating data quality before a release build

## Steps
1. Identify the target CSV from the department name
2. Read with explicit encoding: `open(file, encoding='utf-8-sig')`
3. Parse header row — print exact column names (for TRANSLATIONS updates)
4. Count rows and flag: empty rows, None-heavy columns, unexpected characters
5. Validate required columns exist: `NOMBRE`, `Semana`, and at least one `actividad N`
6. Check for encoding artifacts: scan for `�` (replacement char) — if found, re-read with `latin-1`
7. Report schema:
   - Column names (exact, case-sensitive)
   - Row count
   - Any quality issues found
8. Output: a clean Python dict or dataframe ready for build_activity_log.py

## Validation Checklist
- [ ] File exists and is readable
- [ ] Encoding is UTF-8 or UTF-8-BOM (no replacement chars)
- [ ] Header row is row 1 (not row 0 if there's a title)
- [ ] NOMBRE column present
- [ ] At least one `actividad` column present
- [ ] No fully empty rows in the middle of data
- [ ] All Spanish text in TRANSLATIONS dict (flag any missing keys)

## Output Format
```
CSV: <filename>
Rows: <n>
Columns: <list>
Issues:
  - <issue description> (row/col reference)
Status: READY / NEEDS_FIX
```
