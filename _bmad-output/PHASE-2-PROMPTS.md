# Phase 2 Workflow - Prompt Templates
**Purpose:** Ready-to-use prompts for starting new threads (Phase 2)  
**For User:** Thành  
**Created:** 2026-01-22

---

## 📖 **How to Use These Prompts**

1. **Create a new Amp thread**
2. **Copy one of the prompts below** (choose your preferred path)
3. **Paste it as your first message** in the new thread
4. **Agent will auto-load context and begin work**

---

## 🎯 **Path A: Continue STORY-001 (Recommended)**

### Option A1: Start Task 4 (WordEditor Analysis)

```
sử dụng agent @_bmad\core\agents\bmad-master.md để tiếp tục STORY-001

Current Status:
- STORY-001: 3/6 tasks completed (50%)
- Last completed: Task 3 (Distribution List Optimization)
- Session: 2026-01-22, 5.5 hours

Next Task: STORY-001 / Task 4 - Review WordEditor Hack

Objectives:
1. Analyze current WordEditor implementation in GenerateCheckList.cs
2. Understand COM Interop usage (Word.Application)
3. Evaluate PropertyAccessor alternative
4. Implement optimization if viable
5. Benchmark performance improvement

Reference Documents:
- @_bmad-output\PHASE-2-WORKFLOW.md (Task 4 Details section)
- @_bmad-output\STORY-001-performance-review.md
- @_bmad-output\SESSION-LOG.md

Success Criteria:
- Documented current WordEditor usage
- Tested PropertyAccessor alternative
- Performance benchmarks captured
- Implementation (if viable) with tests

Begin with: Locate all WordEditor usage in codebase
```

### Option A2: Start Task 5 (String Optimization)

```
sử dụng agent @_bmad\core\agents\bmad-master.md để tiếp tục STORY-001

Current Status:
- STORY-001: 3/6 tasks completed (50%)
- Last completed: Task 3 (Distribution List Optimization)
- Session: 2026-01-22, 5.5 hours

Next Task: STORY-001 / Task 5 - Optimize String Replacements

Objectives:
1. Audit all string operations in GenerateCheckList.cs
2. Find opportunities for compiled regex patterns
3. Identify StringBuilder usage opportunities
4. Implement optimizations
5. Measure performance improvement

Reference Documents:
- @_bmad-output\PHASE-2-WORKFLOW.md (Task 5 Details section)
- @_bmad-output\STORY-001-performance-review.md
- Line 1055: CidRegex example (already optimized)

Success Criteria:
- All string operations audited
- Additional compiled regex added where beneficial
- StringBuilder applied where improving performance
- Performance improvement measured

Begin with: Find all Regex.Replace() and string operations
```

### Option A3: Auto-Continue (Task 4, then 5)

```
sử dụng agent @_bmad\core\agents\bmad-master.md để tiếp tục STORY-001 Tasks 4-5

Current Status:
- STORY-001: 3/6 tasks completed (50%)
- Session: 2026-01-22, 5.5 hours
- Target: 100% completion of STORY-001

Workflow:
1. Task 4: WordEditor Hack Review (3.5 hours)
   - Analyze current implementation
   - Evaluate PropertyAccessor alternative
   - Implement if viable
   - Benchmark results

2. Task 5: String Replacement Optimization (1 hour)
   - Audit string operations
   - Apply compiled regex & StringBuilder
   - Measure improvement

3. Task 6: Mark Complete (already optimized)
   - Verify using Dictionary<string, bool>
   - Close as COMPLETED

Reference Documents:
- @_bmad-output\PHASE-2-WORKFLOW.md
- @_bmad-output\STORY-001-performance-review.md
- @_bmad-output\EXECUTION-SUMMARY.md

Expected Result: 100% STORY-001 completion + final report

Begin with: Start Task 4 analysis
```

---

## 🎯 **Path B: Parallel Execution (STORY-001 + STORY-002)**

### Option B1: Continue STORY-001 + Start STORY-002

```
sử dụng agent @_bmad\core\agents\bmad-master.md để thực hiện STORY-001 Tasks 4-5 + STORY-002 Tasks 4&6 song song

Current Status:
- STORY-001: 3/6 tasks completed (50%)
- STORY-002: 3/6 tasks completed (50%) - overlaps with STORY-001
- Session: 2026-01-22, 5.5 hours

Parallel Workflow:
TRACK 1 (STORY-001):
- Task 4: WordEditor Review (3.5 hrs)
- Task 5: String Optimization (1 hr)

TRACK 2 (STORY-002):
- Task 4: Optimize ConfirmationWindowViewModel batch updates
- Task 6: Optimize ResourceService usage

Overlap Notes:
- STORY-002 Tasks 1,2,3,5 already done via STORY-001
- Can focus only on STORY-002 Tasks 4 & 6

Reference Documents:
- @_bmad-output\STORY-001-performance-review.md
- @_bmad-output\STORY-002-deep-optimization.md
- @_bmad-output\PHASE-2-WORKFLOW.md

Coordination:
- Update both stories simultaneously
- Cross-reference completed tasks
- Merge results for final reports

Begin with: Start STORY-001 Task 4 + STORY-002 Task 4
```

---

## 🎯 **Path C: Switch to STORY-002**

### Option C1: Complete Remaining STORY-002 Tasks

```
sử dụng agent @_bmad\core\agents\bmad-master.md để hoàn thành STORY-002

Current Status:
- STORY-002: 3/6 tasks completed (50%)
- Completed (via STORY-001):
  ✅ Task 1: Refactor CsvFileHandler → IEnumerable
  ✅ Task 2: Optimize ToggleSendButton → .All()
  ✅ Task 3: Optimize GenerateCheckList (Regex, HashSet)
  ✅ Task 5: Replace Thread.Sleep (via ComRetryHelper)

Remaining Tasks:
- Task 4: Optimize ConfirmationWindowViewModel initialization (batch collection updates)
- Task 6: Optimize ResourceService usage if applicable

Story: STORY-002 - Deep Performance Optimization (Speed Turning)
Feature: Advanced Optimization
Effort: 8 points
Status: in-progress → completion

Target Files:
- OutlookOkan/Handlers/CsvFileHandler.cs
- OutlookOkan/Services/SettingsService.cs
- OutlookOkan/ViewModels/ConfirmationWindowViewModel.cs
- OutlookOkan/Models/GenerateCheckList.cs
- OutlookOkan/Services/ResourceService.cs

Reference Documents:
- @_bmad-output\STORY-002-deep-optimization.md
- @_bmad-output\STORY-001-performance-review.md (overlapping context)
- @_bmad-output\EXECUTION-SUMMARY.md

Begin with: Analyze ConfirmationWindowViewModel initialization
```

---

## 📋 **Path D: Generate Final Reports**

### Option D1: Finalize STORY-001 and Create Final Report

```
sử dụng agent @_bmad\core\agents\bmad-master.md để hoàn thành STORY-001 với báo cáo cuối cùng

Current Status:
- STORY-001: 3/6 tasks completed (50%)
- Phase 1 session: 2026-01-22, 5.5 hours
- All code changes committed and documented

Final Report Objectives:
1. Verify all 3 completed tasks meet acceptance criteria
2. Close tasks 4-5 (postponed to Phase 3 if needed)
3. Mark Task 6 as already completed
4. Generate comprehensive STORY-001 Final Report
5. Create performance baseline for future comparison
6. Archive Phase 1 session

Report Should Include:
- Executive summary (1 page)
- All 6 tasks status and details
- Combined performance metrics
- Code quality improvements
- Lessons learned
- Recommendations for Phase 2

Reference Documents:
- @_bmad-output\TASKS-COMPLETED-SUMMARY.md
- @_bmad-output\EXECUTION-SUMMARY.md
- All TASK-00X-COMPLETION-REPORT.md files

Output:
- STORY-001-FINAL-REPORT.md (comprehensive)
- Performance-Baseline.md (for future comparison)
- Phase-1-Summary.md (archive)

Begin with: Review all completed work and acceptance criteria
```

---

## 🔧 **Special Prompts**

### For Context Recovery (If Needed)

```
sử dụng agent @_bmad\core\agents\bmad-master.md để tìm hiểu tình trạng hiện tại

I'm starting a new thread to continue OutlookOkan optimization work.

Context Needed:
- What was completed in Phase 1?
- Current STORY-001 status (3/6 tasks)
- Available paths forward
- Recommended next steps

Reference Documents:
- @_bmad-output\INDEX.md (complete navigation)
- @_bmad-output\QUICK-STATUS.txt (summary)
- @_bmad-output\SESSION-LOG.md (what was done)

Please summarize:
1. Phase 1 completion status
2. Phase 2 available paths
3. Recommended next action
4. Time estimates for remaining tasks
```

### For Quick Status Check

```
sử dụng agent @_bmad\core\agents\bmad-master.md

Show me:
1. Current progress (STORY-001 & STORY-002)
2. What's been completed
3. What's available next
4. Time estimates
5. Recommended path

Reference: @_bmad-output\PHASE-2-WORKFLOW.md
```

---

## 📊 **Prompt Selection Guide**

### Choose Based On Your Goal:

| Goal | Use Prompt |
|------|-----------|
| Continue STORY-001 logically | A3 (Auto-continue) |
| Focus on WordEditor analysis | A1 (Task 4 only) |
| Quick string optimization | A2 (Task 5 only) |
| Maximize efficiency (parallel) | B1 (Parallel tracks) |
| Complete STORY-002 | C1 (STORY-002 only) |
| Wrap up and finalize | D1 (Final report) |
| Quick recovery | Special: Context Recovery |

---

## ⏱️ **Time Estimates**

| Path | Duration | Output |
|------|----------|--------|
| A1 (Task 4) | 3.5 hours | WordEditor optimization |
| A2 (Task 5) | 1 hour | String optimization |
| A3 (Tasks 4-5) | 4.5 hours | 100% STORY-001 completion |
| B1 (Parallel) | 4.5 hours | STORY-001 + STORY-002 progress |
| C1 (STORY-002) | 2-3 hours | Remaining STORY-002 tasks |
| D1 (Final) | 1 hour | Comprehensive final report |

---

## 🎯 **Recommended Starting Prompt**

**Best for: Clean continuation with clear objectives**

```
sử dụng agent @_bmad\core\agents\bmad-master.md để tiếp tục STORY-001 Tasks 4-5

Phase 1 Complete: 3/6 tasks (50% done)
Session: 2026-01-22, 5.5 hours work
Performance Gain So Far: 97% email processing improvement

Next: Complete STORY-001 with final 3 tasks
- Task 4: WordEditor Hack Review (3.5 hrs)
- Task 5: String Optimization (1 hr)
- Task 6: Mark Complete (already done)

References:
- @_bmad-output\PHASE-2-WORKFLOW.md
- @_bmad-output\STORY-001-performance-review.md
- @_bmad-output\EXECUTION-SUMMARY.md

Workflow:
1. Start with Task 4 analysis
2. Implement Task 4 optimization
3. Move to Task 5
4. Generate final STORY-001 report
5. Completion: 100% STORY-001

Begin: Analyze current WordEditor usage in GenerateCheckList.cs
```

---

## 📋 **Checklist Before Posting Prompt**

- [ ] Copied one of the templates above
- [ ] Verified agent path: `@_bmad\core\agents\bmad-master.md`
- [ ] Confirmed referenced documents exist
- [ ] Chosen appropriate path (A, B, C, or D)
- [ ] Ready to paste in new thread

---

## 💡 **Tips for Best Results**

1. **Keep the agent path** - Always reference `@_bmad\core\agents\bmad-master.md`
2. **Reference documents** - Use @references to provide context
3. **Clear objectives** - State what you want to accomplish
4. **Be specific** - Reference specific tasks/stories
5. **Include status** - Show what's already been done

---

## 🚀 **Example: Creating New Thread**

**Step 1:** In Amp, click "New Thread"

**Step 2:** Paste one of the prompts (e.g., A3 - Auto-continue)

**Step 3:** Hit send

**Agent will:**
- ✅ Load bmad-master agent
- ✅ Read referenced documents
- ✅ Understand context automatically
- ✅ Begin working on chosen task

---

**Template Pack Prepared By:** BMad Master Executor  
**Date:** 2026-01-22  
**Status:** ✅ Ready for Phase 2 threads

---

## Quick Copy-Paste (Recommended Default)

```
sử dụng agent @_bmad\core\agents\bmad-master.md để tiếp tục STORY-001

Current: 3/6 tasks done (50%)
Session: 2026-01-22, 5.5 hours
Goal: Complete remaining 3 tasks

Path: Task 4 (WordEditor) → Task 5 (Strings) → Task 6 (mark done)
Time: ~4.5 hours
Outcome: 100% STORY-001 completion

Docs:
- @_bmad-output\PHASE-2-WORKFLOW.md
- @_bmad-output\STORY-001-performance-review.md

Begin: Start Task 4 analysis now
```

Just copy this and paste in new thread to get started! ✅
