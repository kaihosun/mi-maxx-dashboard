# /build — Run Full Activity Log Build Pipeline

Trigger the Backend Agent to run the complete build pipeline and validate outputs.

## Workflow

1. Invoke the **Backend Agent** with the `build-activity-log` skill
2. Pre-build: validate all 15 CSVs exist
3. Run: `python3 build_activity_log.py`
4. On SUCCESS:
   - Confirm both Excel files created/updated
   - Report sheet count and any warnings
5. On FAILURE:
   - Escalate to **Debugger Agent** with `debug-python` skill
   - Return diagnosis and fix, then re-run build

## Usage
```
/build
```
