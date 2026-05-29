# Skill: Build Activity Log

## Purpose
Run the full end-to-end build pipeline: read all 15 department CSVs, apply translations, and generate the two Activity Log Excel workbooks.

## When to Use
- After any CSV update to regenerate the Excel outputs
- Before a release or QA review
- After modifying `build_activity_log.py` to verify correctness

## Steps
1. Verify all 15 CSV files exist (list from backend-kb.md)
2. Run: `python3 build_activity_log.py 2>&1`
3. Check exit code: 0 = success, non-zero = failure
4. On success:
   - Confirm `Activity Log (Admin Staff) - COMPLETE.xlsx` exists and is non-zero size
   - Confirm `Activity Log (Admin Staff) - Traffic Control.xlsx` exists and is non-zero size
   - Open COMPLETE.xlsx and verify sheet count == 15
5. On failure:
   - Capture full traceback
   - Identify error category (see Debugger KB)
   - Escalate to Debugger Agent with: error type, file, line number, and relevant CSV

## Pre-Build Checklist
- [ ] All 15 CSVs present in project root
- [ ] `build_activity_log.py` has no syntax errors: `python3 -m py_compile build_activity_log.py`
- [ ] TRANSLATIONS dict covers all Spanish text in current CSVs
- [ ] SHEET_NAMES has exactly 15 entries

## Post-Build Validation
```bash
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('Activity Log (Admin Staff) - COMPLETE.xlsx')
print(f'Sheets: {len(wb.sheetnames)}')
print(wb.sheetnames)
"
```
Expected: 15 sheets matching SHEET_NAMES values.

## Escalation
If build fails after 2 attempts → hand off to Debugger Agent with full traceback.
