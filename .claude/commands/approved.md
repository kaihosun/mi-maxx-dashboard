# /approved — Release & Push to GitHub

Run a full QA audit, then stage, commit, and push the current changes to GitHub.

## Workflow

### Step 1: Pre-release QA
Invoke the **QA Agent** with the `quality-audit` skill.
- Must receive PASS verdict before proceeding.
- If FAIL: report findings to user and STOP. Do not commit.

### Step 2: Git Commit & Push
Invoke the **Git Control Agent** with the `commit-push` skill.
- Stage only relevant files (scripts, agents, docs) — skip `.DS_Store`, `__pycache__`, raw CSVs unless user confirms.
- Use commit type `release` with version if applicable.
- Push to `origin main` (confirm branch with user if uncertain).

### Step 3: Report
Return:
- QA verdict summary
- Commit hash and message
- Files committed
- Remote push status

## Usage
```
/approved
```
Or with a custom message:
```
/approved feat(dashboard): add Palletizing weekly chart
```
