# ⚠️ RETRY INSTRUCTION — Overflow Auto-Recovery
_Auto-generated: 2026-03-04T10:51:58.354Z_
_Recovery attempt: #3 (source: pb-size-estimate)_

## Error
Previous request failed: **prompt is too long**
- Actual tokens: **202,450**
- Maximum allowed: **200,000**
- Over by: **2,450 tokens**

## What Happened
- Context Guardian detected the overflow error automatically
- Emergency compact was executed to free context space
- Session summary saved for continuity

## Next Steps (for AI)
1. Read `SESSION_SUMMARY.md` in this directory for context
2. **RETRY the user's last action** — context has been compacted
3. Do NOT re-read files already discussed — use the session summary
4. If still over limit, suggest `/new-session` to the user

> **CRITICAL**: Start a new session if this is the 2nd+ recovery attempt.