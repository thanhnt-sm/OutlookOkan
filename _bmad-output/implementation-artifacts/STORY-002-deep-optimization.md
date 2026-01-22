# Story: Deep Performance Optimization (Speed Turning)

**ID:** STORY-002
**Title:** System-Wide Speed Turning & Optimization
**Status:** planning
**Effort:** 8
**Feature:** Advanced Optimization

## Description
Based on the Phase 2 Adversarial Review, we will implement system-wide optimizations to reduce memory pressure and CPU cycles (Speed Turning).

## Acceptance Criteria
- [ ] **AC1:** Whitelists and Domain lists use `HashSet<string>` for O(1) lookups.
- [ ] **AC2:** UI logic in `ConfirmationWindowLogger` stops iterating lists repeatedly.
- [ ] **AC3:** Regex replacements are compiled and reused.
- [ ] **AC4:** List allocations in hot loops are minimized.

## Tasks
- [x] **Task 1:** [Critical] Refactor `CsvFileHandler` to return `IEnumerable<T>` and update `SettingsService` to use `Dictionary<string, bool>` for all Whitelists.
- [x] **Task 2:** [Critical] Optimize `ConfirmationWindowViewModel.ToggleSendButton` to use `.All()` (short-circuiting) instead of `.Count()`.
- [x] **Task 3:** [Medium] Optimize `GenerateCheckList` internal logic:
    - [x] Use Compiled Regex for line break replacements (`CidRegex`).
    - [x] Use `HashSet` (Dictionary) for local domain tracking in `CountRecipientExternalDomains`.
- [ ] **Task 4:** [Low] Optimize `ConfirmationWindowViewModel` initialization (batch collection updates).
- [ ] **Task 5:** [Low] Replace hardcoded `Thread.Sleep` in any remaining helper classes if found.
- [ ] **Task 6:** [Medium] Optimize `ResourceService` usage if applicable.

## Dev Agent Record
- **Target Files:**
    - `Handlers/CsvFileHandler.cs`
    - `Services/SettingsService.cs`
    - `ViewModels/ConfirmationWindowViewModel.cs`
    - `Models/GenerateCheckList.cs`
