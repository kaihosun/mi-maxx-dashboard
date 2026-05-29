# Skill: Task Breakdown

## Purpose
Decompose a vague or large request into discrete, delegatable tasks — one task per agent, with clear inputs, outputs, and sequence.

## When to Use
- User gives a high-level goal ("build the web graph", "fix the pipeline", "add a new department")
- The orquestador needs to know which agents to invoke and in what order
- A plan already exists and needs to be converted into executable steps

## Steps

1. **Read the existing plan** if one exists (`thinking-plan*.md`)

2. **Identify all agents needed** — map each step to exactly one agent:
   - Data transformation → backend-agent
   - Visual output → frontend-agent
   - Python error → debugger-agent
   - Validation → qa-agent
   - Git operations → git-control-agent
   - New plan needed → planner-agent (recursive)

3. **Define task boundaries** — a task is bounded when:
   - It can be described in one sentence
   - It has a single observable output (file, test result, printed value)
   - It can fail independently without breaking other tasks

4. **Sequence tasks** — identify:
   - Parallel tasks (no dependency between them → run concurrently)
   - Sequential tasks (output of A is input of B → run in order)

5. **Output format** — return a table:

```markdown
| # | Task | Agent | Input | Output | Depends on |
|---|------|-------|-------|--------|-----------|
| 1 | ... | backend-agent | graph.json | server.py | — |
| 2 | ... | frontend-agent | graph.json | index.html | — |
| 3 | ... | qa-agent | server.py, index.html | PASS/FAIL | 1, 2 |
| 4 | ... | git-control-agent | all outputs | commit | 3 |
```

6. **Flag blockers** — list any task that can't start without user input or external dependency

## Quality Check
- [ ] Every task maps to exactly one agent
- [ ] No task has "and" in its description (if it does, split it)
- [ ] Dependencies are explicit — no implicit ordering
- [ ] Parallel tasks are identified to save time
