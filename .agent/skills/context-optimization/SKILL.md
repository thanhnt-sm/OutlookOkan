---
name: context-optimization
description: Strategies for managing AI agent context window effectively. Use when conversations are long, working with large codebases, or context is degrading.
---

# Context Optimization Skill

## When This Skill Activates
- Context approaching 60-70% of model's limit
- Agent responses becoming less accurate or repetitive
- User reports agent "forgetting" earlier work
- Working with >10 files simultaneously
- After 10+ substantive exchanges in one conversation

## Strategy 1: Context Partitioning (Isolate)
Split work across focused sessions:
- One conversation per component/module
- Use `_handoff.md` to carry state between sessions
- Each session focuses on ONE clear objective
- Complex feature = Planning session → Implementation session → Verification session

## Strategy 2: Observation Masking (Compress)
Replace verbose tool outputs with concise references:
- Instead of keeping full file contents: "Read `utils.py` (250 lines, helper functions for date/string manipulation)"
- Instead of full command output: "Build succeeded, 0 errors, 2 warnings (unused imports in auth.py:L12, L45)"
- Instead of full search results: "Found 3 callers of `processPayment()`: checkout.py:L89, refund.py:L34, subscription.py:L112"

## Strategy 3: Progressive Disclosure (Select)
Read code structure BEFORE content:
1. `view_file_outline` → understand what exists (cheapest)
2. `view_code_item` → read only the relevant function (targeted)
3. `view_file` with line range → only when editing requires surrounding context (precise)
4. Full `view_file` → only as last resort for understanding complex interactions (expensive)

## Strategy 4: Dependency-Aware Reading (Select)
When modifying function A:
1. Read A's implementation
2. Find callers of A (`grep_search`)
3. Read caller signatures ONLY (not full files)
4. Make change to A
5. Update callers if interface changed
6. Run tests if available

## Strategy 5: Context Health Monitor (Assess)
After every 5 substantive exchanges:
- [ ] Are all files in context still relevant to CURRENT task?
- [ ] Has the task scope changed since we started?
- [ ] Can completed sub-tasks be summarized to free context?
- [ ] Is the agent repeating itself or losing accuracy?
- [ ] Should we suggest `/context-handoff`?

## Strategy 6: Just-In-Time Context (Select)
Do NOT pre-load all potentially relevant files at the start:
- Load files only when the task specifically requires them
- If a file MIGHT be needed, note its existence but don't read it yet
- Read it only when you actually need to make a decision about it

## Strategy 7: Externalized Memory (Write)
For long multi-step tasks, write intermediate state to files:
- Save analysis results to `_analysis.md` instead of keeping in context
- Save implementation plans to artifacts instead of conversation memory
- Reference files by path instead of quoting their contents

## Strategy 8: Sub-Agent Partitioning (Isolate)
For truly complex tasks that require >150K tokens:
- Break into independent sub-tasks
- Each sub-task runs in its own conversation/agent
- Coordinate through shared files (`_handoff.md`, `task.md`)
- Main conversation only orchestrates, doesn't do deep work

## Context Budget Reference

| Model | Hard Limit | Safe Operating Range | Danger Zone |
|-------|:---:|:---:|:---:|
| Claude Sonnet/Opus | 200K | <120K | >150K |
| Gemini 3 Pro High | 1M | <600K | >800K |
| Gemini 3 Pro Low | 1M | <400K | >600K |
| Gemini 3 Flash | 1M | <300K | >500K |
