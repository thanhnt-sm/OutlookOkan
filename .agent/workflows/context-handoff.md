---
description: Auto-handoff context to a new conversation when context is getting large
---
// turbo-all

# Context Handoff Workflow

When context is getting large or you're about to hit token limits, use this workflow to create a clean handoff.

## Steps

1. Create a summary of ALL work done in this conversation using the following structure:

```markdown
## Context Handoff — [Current Date]

### ✅ Completed
- [List all completed tasks with file paths]

### 🔄 In Progress
- [Current task and its state]

### 🔑 Key Decisions Made
- [Important architectural or design decisions]

### 📁 Files Changed
- [List each file with a one-line description of changes]

### ⚠️ Known Issues
- [Any bugs, warnings, or concerns discovered]

### 📋 Next Steps (Priority Order)
1. [Most important next task]
2. [Second priority]
3. [Third priority]
```

2. Save this summary to a file called `_handoff.md` in the project root.

3. Tell the user: "Context is getting large. I've created `_handoff.md` with a full summary. Please start a new conversation and paste the contents of `_handoff.md` as your first message."
