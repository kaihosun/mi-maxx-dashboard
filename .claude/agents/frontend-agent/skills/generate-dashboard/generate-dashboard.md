# Skill: Generate Dashboard

## Purpose
Create a new department sheet in the Activity Log Excel workbook from scratch using openpyxl.

## When to Use
- Adding a new department that doesn't have a sheet yet
- Rebuilding a sheet that got corrupted or deleted
- Creating a new visual report layout for an existing department

## Steps
1. Read the department's CSV to get exact column headers
2. Determine target sheet name from `SHEET_NAMES` in `build_activity_log.py`
3. Open (or create) the target workbook with `openpyxl.load_workbook()` or `Workbook()`
4. Create/overwrite the sheet: `wb.create_sheet(sheet_name)` or `wb[sheet_name]`
5. Write header row:
   - Row 1: merged title cell (company name + department), bold, BLACK bg, WHITE font
   - Row 2: column headers, LIGHT_GRAY fill, BLACK font, bold, centered
6. Write data rows from CSV — translate cell values using TRANSLATIONS dict
7. Apply borders (thin, all sides) to all data cells
8. Set `freeze_panes = "A3"` (freeze title + header rows)
9. Set column widths: compute `max(len(str(v)) for v in col_values) + 2`
10. Add auto_filter on header row
11. Save workbook

## Template (openpyxl)
```python
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
thin = Side(style='thin', color='000000')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
header_fill = PatternFill("solid", fgColor="F2F2F2")
title_fill  = PatternFill("solid", fgColor="000000")
```

## Output Validation
- Sheet exists in workbook
- Row count = CSV row count + 2 (title + header)
- All column headers present and translated
- No None values in cells (use `""` as fallback)
