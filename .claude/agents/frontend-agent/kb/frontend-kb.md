# Frontend KB — MI Technologies MTY MAXX

## Project Outputs
| File | Purpose |
|------|---------|
| `Activity Log (Admin Staff) - COMPLETE.xlsx` | Full bilingual activity log (15 sheets) |
| `Activity Log (Admin Staff) - Traffic Control.xlsx` | Traffic control summary view |

## Established Color Palette
```python
WHITE      = "FFFFFF"
BLACK      = "000000"
LIGHT_GRAY = "F2F2F2"
```
Do not introduce new colors without approval.

## Department → Sheet Name Mapping
| Spanish | English Sheet |
|---------|--------------|
| Almacén | Warehouse |
| Auditoría | Audit |
| Calidad | Quality |
| Clasificación | Sorting |
| FFT | FFT Line |
| Incoming | Incoming |
| Logística | Logistics |
| Open Cell | Open Cell |
| Shipping B2B | Shipping B2B |
| Shipping B2C | Shipping B2C |
| Paletizado | Palletizing |
| EHS | EHS |
| Mantenimiento | Maintenance |
| IT | IT |
| RH | HR |

## Excel Style Rules
- Headers: Font bold, PatternFill LIGHT_GRAY, Alignment center
- Data rows: Font normal, no fill, Alignment left
- Borders: thin Side on all 4 edges
- freeze_panes: always freeze row 1 (header row)
- Column widths: auto-fit based on content (use `column_dimensions[letter].width`)
- Merged cells: only for title rows — check existing merges before writing

## Skills Available
- `generate-dashboard` — creates a new department sheet from scratch
- `update-chart` — updates an existing chart or table in place

## Source References
- CSV column headers vary by department — always read the file header row before building
- Enriched CSVs (`_enriquecido.csv`) have additional metadata columns — prefer these for dashboards
- `build_activity_log.py` is the single source of truth for constants
