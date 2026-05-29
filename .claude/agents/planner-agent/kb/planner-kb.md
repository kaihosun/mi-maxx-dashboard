# Planner KB — MI Technologies MTY MAXX

## Project Landscape

| Layer | What exists | Key file |
|-------|------------|----------|
| Data | 15 dept CSVs + `_enriquecido.csv` variants | project root |
| Build | openpyxl pipeline | `build_activity_log.py` |
| Output | 16-sheet Excel (EN) + 15-sheet Excel (ES) | project root |
| Knowledge graph | 601 nodes, 898 edges, 75 communities | `graphify-out/graph.json` |
| Web viz plan | FastAPI + D3-force + Canvas + SSE | `thinking-plan.md` |
| Graph HTML | vis-network.js (reference only) | `graphify-out/graph.html` |

## Agent Roster & Capabilities

| Agent | Executes | Tools |
|-------|---------|-------|
| frontend-agent | Excel layouts, charts, HTML reports | Read, Write, Edit, Bash |
| backend-agent | Python, CSV, build pipeline | Read, Write, Edit, Bash |
| debugger-agent | Python errors, encoding bugs | Read, Bash, Edit |
| qa-agent | Code review, SOP validation, pre-release | Read, Bash |
| git-control-agent | git commit, branch, push | Bash, Read |
| **planner-agent** | Implementation plans (read-only) | Read, Bash |

## Slash Commands Available
- `/build` — runs build_activity_log.py via Backend Agent
- `/approved` — QA audit + commit + push
- `/handoff` — saves session context

## Known Constraints
- Python: macOS darwin, UTF-8 required for Spanish accents (á, é, í, ó, ú, ñ)
- Excel: openpyxl — always set Font/Fill/Alignment/Border explicitly, never rely on inheritance
- No React/Vue in frontend tasks — Vanilla JS only unless explicitly approved
- CSV schema: `_enriquecido.csv` columns are additive — never remove existing columns
- Graph: `graphify-out/` is generated — do NOT edit graph.json manually

## Plan Template

```markdown
# Plan: [Feature Name]

## Objective
One sentence: what will exist when this is done that doesn't exist now.

## Constraints
- [technical constraint 1]
- [dependency on existing file/agent]

## Phases

### Phase 1 — [Name]
**Agent:** [agent-name]
**Input:** [file or data]
**Output:** [file or observable result]
**Test:** [how to verify phase 1 is complete]
Steps:
1. ...

### Phase 2 — [Name]
...

## File Map
```
path/to/new/file.py    ← what it does
path/to/edited/file.py ← what changes
```

## Agent Assignments
| Phase | Agent | Handoff to |
|-------|-------|-----------|
| 1 | backend-agent | frontend-agent |
| 2 | frontend-agent | qa-agent |
| 3 | qa-agent | git-control-agent |

## Open Questions
- [ ] Question that needs user input before execution
```

## Current Active Plans
- `thinking-plan.md` — Web grafo con sinapsis en vivo (FastAPI + D3-force + Canvas + SSE)
  - Status: Diseño completo, Fase 1–3 definidas, **pendiente de ejecución**
