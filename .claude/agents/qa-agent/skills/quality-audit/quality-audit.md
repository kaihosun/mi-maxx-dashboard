# Skill: Quality Audit

## Purpose
Full pre-release quality check — validates code, data, and documentation before any build is approved for delivery.

## When to Use
- Before running `/approved` (git push)
- After a major change to `build_activity_log.py`
- When the user requests a QA review of the current project state

## Steps

### Phase 1: Code Quality
1. Check `build_activity_log.py` syntax: `python3 -m py_compile build_activity_log.py`
2. Verify `SHEET_NAMES` has exactly 15 entries
3. Verify `TRANSLATIONS` covers all activity labels (actividad 1–21) in both directions
4. Scan for hardcoded department strings outside of constants — flag any found
5. Check no raw `None` values are written to Excel cells

### Phase 2: Data Quality
6. For each CSV in the project: check file exists and is readable
7. Spot-check 3 departments for encoding: `Almacén`, `Logística`, `Clasificación` (accent-heavy)
8. Verify enriched CSVs (`_enriquecido.csv`) have a superset of base CSV columns

### Phase 3: Build Validation
9. Run `python3 build_activity_log.py 2>&1`
10. Verify exit code == 0
11. Verify output: `Activity Log (Admin Staff) - COMPLETE.xlsx` exists, size > 0
12. Verify sheet count == 15

### Phase 4: Documentation Check
13. Verify `CLAUDE.md` references all 5 agents
14. Verify each agent has a KB and at least one skill
15. Verify `.gitignore` (if repo exists) excludes `.DS_Store` and `__pycache__`

## Verdict Output
```
QA AUDIT — MI Technologies MTY MAXX
Date: <today>
Phase 1 - Code:    PASS / FAIL
Phase 2 - Data:    PASS / FAIL
Phase 3 - Build:   PASS / FAIL
Phase 4 - Docs:    PASS / FAIL
Overall: PASS / FAIL

Findings:
  [P/F] <item> — <detail>

Recommendation: <APPROVE FOR RELEASE / FIX REQUIRED>
```
