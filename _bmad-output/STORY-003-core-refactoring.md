# Story: Core Logic Refactoring (GenerateCheckList)

**ID:** STORY-003
**Title:** Modularization of Core Business Logic
**Status:** planning
**Effort:** 13
**Feature:** Refactoring & Tech Debt

## Description
The `GenerateCheckList` class has grown to over 2000 lines, violating the Single Responsibility Principle. It handles email analysis, recipient resolution, whitelist checking, and UI model generation. This story focuses on breaking it down into smaller, testable services.

## Acceptance Criteria
- [ ] **AC1:** `GenerateCheckList.cs` is reduced to < 500 lines (Orchestrator pattern).
- [ ] **AC2:** `WhitelistService` created to handle all domain/address checking.
- [ ] **AC3:** `RecipientResolver` created to handle Exchange/Contact parsing.
- [ ] **AC4:** `EmailAnalyzer` created to handle Body/Subject/Attachment inspections.
- [ ] **AC5:** No regression in checking logic (verified by tests).

## Tasks
- [ ] **Task 1:** [Architectural] Design the new Service Interfaces (`IWhitelistService`, `IRecipientResolver`, `IEmailAnalyzer`).
- [ ] **Task 2:** [Refactor] Extract Whitelist/Domain logic to `WhitelistService`.
- [ ] **Task 3:** [Refactor] Extract Recipient/Exchange resolution to `RecipientResolver`.
- [ ] **Task 4:** [Refactor] Extract Body/Subject/Attachment analysis to `EmailAnalyzer`.
- [ ] **Task 5:** [Cleanup] Update `GenerateCheckList` to use dependency injection (or factory) for these services.

## Dev Agent Record
- **Target Files:**
    - `Models/GenerateCheckList.cs`
    - `Services/WhitelistService.cs` (New)
    - `Services/RecipientResolver.cs` (New)
    - `Services/EmailAnalyzer.cs` (New)
