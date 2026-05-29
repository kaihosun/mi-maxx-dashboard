# Skill: Update Chart

## Purpose
Update an existing chart, table, or visual element in the Activity Log Excel workbook without rebuilding the entire sheet.

## When to Use
- A chart's data range needs to extend to new rows/columns
- Color scheme needs updating on an existing chart
- Adding a KPI summary row to an existing sheet
- Modifying a specific column's formatting

## Steps
1. Identify the target sheet and cell range to modify
2. Load workbook: `wb = openpyxl.load_workbook(target_file)`
3. Select sheet: `ws = wb[sheet_name]`
4. Check for merged cells before writing: inspect `ws.merged_cells`
5. For chart data updates: modify the Reference objects pointing to the new data range
6. For formatting updates: re-apply Font/PatternFill/Border to the target range only
7. Do NOT touch rows/columns outside the specified update range
8. Save workbook with `wb.save(target_file)`
9. Run `python3 build_activity_log.py` to verify no conflicts

## Safety Rules
- Always `load_workbook()` — never overwrite the file from scratch
- Backup check: confirm the file exists before opening
- Merged cell check: never write to a cell inside a merge region
- Range discipline: only modify the explicitly requested range

## Common Update Patterns
```python
# Extend data range for existing chart
ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}{max_row}"

# Update a single cell's style
ws["A1"].font = Font(bold=True, color="FFFFFF")
ws["A1"].fill = PatternFill("solid", fgColor="000000")
```
