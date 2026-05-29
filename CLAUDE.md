## Project
MI Technologies MTY MAXX — TV refurbishment maquiladora, Monterrey.
Main deliverable: `Activity Log (Admin Staff) - COMPLETE.xlsx` (15-department activity log, built from CSVs via `build_activity_log.py`).

## Agents
This project has 6 local agents in `.claude/agents/`. Each has a KB (`/kb/`) and skills (`/skills/`). Always delegate to the appropriate agent:

| Agent | When to use |
|-------|------------|
| **planner-agent** | Design plans before complex tasks, phase breakdowns, agent delegation maps, `thinking-plan.md` |
| **frontend-agent** | Excel layouts, charts, HTML reports, visual changes |
| **backend-agent** | Python scripting, CSV processing, build pipeline, data logic |
| **git-control-agent** | git commits, branches, pushes, releases |
| **qa-agent** | Quality reviews, SOP validation, pre-release audits |
| **debugger-agent** | Python errors, encoding issues, data anomalies — isolated debugging |

**Rule:** For any task spanning >2 files or >1 agent — invoke planner-agent first.

## Agent Naming Convention
- All sub-agent names must be lowercase, no spaces, no capitals
- Canonical IDs: `planner-agent`, `frontend-agent`, `backend-agent`, `git-control-agent`, `qa-agent`, `debugger-agent`
- Never use display names like "Planner Agent" — use the short id only

## Orchestrator Rules
- When delegating to a sub-agent, use only the Task tool
- Pass clean, focused prompts — no implementation context unless strictly necessary
- If the agent needs to review a file, pass only the file path + one clear instruction

## Slash Commands
- `/approved` — QA audit + commit + push to GitHub
- `/build` — run full build pipeline via Backend Agent
- `/handoff` — save session state for context handoff
- `/ReporteKIPs` — generate a self-contained HTML KPI report from any Excel/CSV file (light theme, Nordic/Cool palette, English insights, delegates to backend → frontend → qa agents)

## graphify

This project has a knowledge graph at graphify-out/ with god nodes, community structure, and cross-file relationships.

Rules:
- ALWAYS read graphify-out/GRAPH_REPORT.md before reading any source files, running grep/glob searches, or answering codebase questions. The graph is your primary map of the codebase.
- IF graphify-out/wiki/index.md EXISTS, navigate it instead of reading raw files
- For cross-module "how does X relate to Y" questions, prefer `graphify query "<question>"`, `graphify path "<A>" "<B>"`, or `graphify explain "<concept>"` over grep — these traverse the graph's EXTRACTED + INFERRED edges instead of scanning files
- After modifying code, run `graphify update .` to keep the graph current (AST-only, no API cost).
