# /handoff — Save Context & Prepare for New Session

Save the current session state so the next Claude session can pick up exactly where this one left off.

## Workflow

1. **Summarize current state**: what was the goal, what was completed, what is still in progress
2. **Save to `handoff.md`** in project root with this structure:

```markdown
# Handoff — <date>

## Goal
<what the user was trying to accomplish>

## Completed
- <item 1>
- <item 2>

## In Progress
- <item> — <current status>

## Next Steps
1. <first thing to do next session>
2. <second thing>

## Context to Remember
- <key decision made>
- <blocker or issue found>
- <agent states: which agents were active>

## Files Modified This Session
- <file> — <what changed>
```

3. **Update CLAUDE.md** if any new agents, skills, or commands were created this session
4. Confirm handoff.md was saved
5. Tell user: "Context saved. Start next session with: 'Read handoff.md and continue from where we left off.'"

## Usage
```
/handoff
```
