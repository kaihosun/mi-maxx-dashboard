# Skill: Create Plan

## Purpose
Produce a structured implementation plan for a new feature or complex task and save it as `thinking-plan.md` (or `thinking-plan-<feature>.md`).

## When to Use
- User asks to plan a new feature before building
- A task spans more than one agent or more than 2 files
- The orquestador needs a phase breakdown before delegating

## Steps

1. **Understand the goal**
   - Read the user's request and any existing plan files (`thinking-plan*.md`) to avoid duplication.

2. **Survey what exists**
   - Run `ls` on relevant directories to see what files already exist.
   - Read key files that the plan will touch (e.g., `build_activity_log.py`, `graph.json` structure, relevant CSVs).
   - Check `graphify-out/GRAPH_REPORT.md` for cross-file relationships if the feature touches multiple modules.

3. **Identify constraints**
   - Encoding requirements (UTF-8, Spanish accents)
   - Existing schemas that cannot break (`_enriquecido.csv`, `SHEET_NAMES`, `TRANSLATIONS`)
   - Tools available per agent (see KB agent roster)
   - macOS-only paths (no Windows separators)

4. **Design phases**
   - Each phase: one agent, one input, one testable output
   - Max 3 phases before requesting user confirmation
   - Each phase names its handoff target

5. **Write the plan** using the template in `planner-kb.md`

6. **Save the plan**
   - New feature: `thinking-plan-<feature>.md` in project root
   - Replacing existing plan: overwrite `thinking-plan.md`

7. **Report to orquestador**
   - Summary: objective, phases, agent assignments, and open questions that need user input

## Quality Check
- [ ] Every file path is concrete (not "some file")
- [ ] Every phase has a testable output
- [ ] Every phase names the executing agent
- [ ] Open questions are listed, not buried in prose
- [ ] Plan references existing files — not proposing duplicates
