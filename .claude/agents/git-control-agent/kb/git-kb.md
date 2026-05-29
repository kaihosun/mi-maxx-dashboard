# Git Control KB — MI Technologies MTY MAXX

## Repository Status
- Location: `/Users/eduardoflores/AI_Projects/Dashboards/Especiales/`
- Git initialized: CHECK with `git status` — if error "not a git repository", run init workflow below

## Init Workflow (if no git repo)
```bash
git init
git add build_activity_log.py CLAUDE.md .claude/
git commit -m "chore: initialize Especiales project with agents and scripts"
```

## .gitignore Essentials
```
.DS_Store
__pycache__/
*.pyc
*.pyo
graphify-out/
handoff.md
thinking-plan.md
```
Note: Decide with user whether raw CSVs and xlsx files go in git (they may contain sensitive employee data).

## Commit Message Convention
Format: `type(scope): short description`
- Types: `feat`, `fix`, `chore`, `docs`, `refactor`, `style`, `test`
- Scopes: `dashboard`, `backend`, `agents`, `qa`, `sop`, `csv`
- Examples:
  - `feat(dashboard): add Palletizing weekly chart`
  - `fix(backend): handle UTF-8-BOM on Almacén CSV`
  - `docs(sop): update incoming SOP to v2`
  - `chore(agents): add QA agent skills`

## Branch Strategy
- `main` — stable, production-ready
- `feature/<name>` — new features
- `fix/<name>` — bug fixes
- `release/vX.Y.Z` — release candidates

## Release Workflow (use /approved command)
1. Bump version in any relevant file
2. `git add` specific files
3. `git commit -m "release: vX.Y.Z"`
4. `git tag vX.Y.Z`
5. `git push origin main --tags`

## Skills Available
- `commit-push` — stages, commits, and pushes in one validated flow
