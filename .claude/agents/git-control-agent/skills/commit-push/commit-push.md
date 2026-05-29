# Skill: Commit and Push

## Purpose
Stage specific files, create a properly formatted commit, and push to the remote repository in one validated flow.

## When to Use
- User says "save progress", "commit changes", "push to GitHub", or "/approved"
- After a successful QA review
- After completing a feature or fix

## Steps
1. Run `git status` — read the full output before proceeding
2. Identify changed files — group by type (scripts, docs, agents, CSVs)
3. Confirm with user if any sensitive files appear (raw CSVs, xlsx with employee data)
4. Stage specific files by name: `git add <file1> <file2> ...`
   - NEVER `git add .` without explicit user approval
   - SKIP: `.DS_Store`, `__pycache__/`, `*.pyc`, `graphify-out/`, `handoff.md`, `thinking-plan.md`
5. Run `git diff --cached` to review exactly what will be committed
6. Write commit message following convention: `type(scope): short description`
7. Commit: `git commit -m "$(cat <<'EOF'\n<message>\n\nCo-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>\nEOF\n)"`
8. Confirm push: ask user for remote and branch if not obvious
9. Push: `git push origin <branch>`
10. Report: commit hash, files committed, remote status

## Commit Message Examples
- `feat(dashboard): add Palletizing sheet to COMPLETE workbook`
- `fix(backend): resolve UTF-8-BOM encoding on Almacén CSV`
- `chore(agents): add QA and Debugger agents with KBs and skills`
- `docs(sop): update incoming SOP section 12 KPIs`

## Never Do
- Force-push to main without explicit instruction
- Commit .DS_Store or __pycache__
- Skip the `git diff --cached` review step
- Commit without reading git status first
