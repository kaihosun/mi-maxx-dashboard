# Skill: Debug Python

## Purpose
Systematic diagnosis and fix of Python errors in `build_activity_log.py` or any project script, without consuming the main orchestrator's context.

## When to Use
- Any `python3 build_activity_log.py` run produces an error
- A script produces unexpected output (wrong row count, missing sheets, corrupted cells)
- The Frontend or Backend agent hits an unresolved error after 1 attempt

## Steps
1. Capture full error: `python3 build_activity_log.py 2>&1 | head -50`
2. Read the traceback bottom-up: last line = actual error, work upward to find source
3. Identify error category from Debugger KB error catalog
4. Isolate: create a minimal reproduction in a temp script if the error is in a complex function
5. Apply the fix from the catalog (or derive one if novel)
6. Test: `python3 build_activity_log.py 2>&1`
7. Verify fix: confirm no new errors, output files exist, sheet count == 15
8. Clean up: remove any debug print statements added during diagnosis
9. Report: root cause + fix applied + test result

## Error Categories (quick ref — full detail in debugger-kb.md)
- `UnicodeDecodeError` → encoding issue on CSV read
- `KeyError` in TRANSLATIONS → missing translation key
- `MergedCell` write error → write to merged region
- `AttributeError: NoneType` → None value in cell
- `FileNotFoundError` on CSV → filename mismatch (accents in filename)

## Escalation
If error persists after 3 fix attempts:
- Document: error, attempts made, what each attempt changed
- Hand back to orchestrator with full diagnostic report
- Do NOT keep trying the same fix in a loop
