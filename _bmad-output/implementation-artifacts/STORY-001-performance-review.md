# Story: Performance Optimization & Speed Turning

**ID:** STORY-001
**Title:** Performance Optimization Code Review
**Status:** in-progress
**Effort:** 5
**Feature:** Core Logic Optimization

## Description
Conduct a deep-dive analysis of the OutlookOkan codebase to identify performance bottlenecks, specifically in the email sending flowchart and checklist generation logic. The goal is to maximize execution speed ("turning") without altering the fundamental solution design.

## Acceptance Criteria
- [ ] **AC1:** `GenerateCheckList` algorithm is analyzed for Big-O complexity and optimized.
- [ ] **AC2:** `ThisAddIn.Application_ItemSend` logic is verified for minimal blocking of the Outlook UI thread.
- [ ] **AC3:** Loop iterations and Dictionary lookups are optimized for speed.
- [ ] **AC4:** VSTO COM Interop calls are minimized or batched where possible.

## Tasks
- [ ] **Task 1:** [Critical] Remove `Thread.Sleep` calls in `GenerateCheckList.cs` and implement non-blocking retry logic.
- [ ] **Task 2:** [Critical] Implement `SettingsCache` in `ThisAddIn.cs` to prevent disk I/O on every `ItemSend`.
- [ ] **Task 3:** [Critical] Refactor `GetExchangeDistributionListMembers` to limit recursion depth and batch COM calls.
- [ ] **Task 4:** [Medium] Review `WordEditor` hack and attempt to replace with lightweight PropertyAccessor.
- [ ] **Task 5:** [Medium] Optimize `CheckList` string replacements using compiled Regex or StringBuilder.
- [ ] **Task 6:** [Medium] Convert `Whitelist` searches to use `HashSet` instead of `List` for O(1) lookup.

## Dev Agent Record
- **File List:**
    - OutlookOkan/ThisAddIn.cs
    - OutlookOkan/Models/GenerateCheckList.cs
    - OutlookOkan/CsvTools.cs
    - OutlookOkan/ViewModels/ConfirmationWindowViewModel.cs
