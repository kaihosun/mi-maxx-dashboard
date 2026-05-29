# Backend KB — MI Technologies MTY MAXX

## Core Script
`build_activity_log.py` — Python 3, openpyxl. Reads CSVs, applies translations, writes Excel workbooks.

## Key Constants (build_activity_log.py)
- `SHEET_NAMES` — dict mapping Spanish department names to English sheet names
- `TRANSLATIONS` — dict mapping Spanish cell text to English equivalents
- `WHITE`, `BLACK`, `LIGHT_GRAY` — hex color strings (visual only — do not change logic)

## Department CSV Files
| Department | CSV File | Enriched |
|-----------|----------|---------|
| Almacén | `Registro Actividades...Almacén.csv` | `_enriquecido.csv` |
| Auditoría | `...Auditoría.csv` | `_enriquecido.csv` |
| Calidad | `...Calidad.csv` | `_enriquecido.csv` |
| Clasificación | `...Clasificación.csv` | `_enriquecido.csv` |
| FFT | `...FFT.csv` | `_enriquecido.csv` |
| Incoming | `...Incoming.csv` | `_enriquecido.csv` |
| Logística | `...Logistica.csv` | `_enriquecido.csv` |
| Open Cell | `...Opencel.csv` | `_enriquecido.csv` |
| Shipping B2B | `...Shipping B2B.csv` | `_enriquecido.csv` |
| Shipping B2C | `...Shipping B2C.csv` | `_enriquecido.csv` |
| Paletizado | `...Paletizado.csv` | (none yet) |
| EHS | `...EHS.csv` | (none yet) |
| Mantenimiento | `...Mantenimiento.csv` | (none yet) |
| IT | `...IT.csv` | (none yet) |
| RH | `...RH.csv` | (none yet) |

## Encoding Rules
- All CSVs: UTF-8 or UTF-8-BOM (`utf-8-sig`)
- Spanish accents (á é í ó ú ñ) must be handled — use `encoding='utf-8-sig'` on CSV reads
- Excel output: openpyxl handles UTF-8 natively, no special config needed

## Activity Label Pattern
- Spanish: `actividad 1` through `actividad 21`
- English: `task 1` through `task 21`
- All in TRANSLATIONS dict

## Build Command
```bash
python3 build_activity_log.py
```
Expected: no errors, two `.xlsx` files created/updated in project root.

## Skills Available
- `process-csv` — reads, validates, and transforms a department CSV
- `build-activity-log` — runs the full build pipeline end-to-end
