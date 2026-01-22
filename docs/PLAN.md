# PLAN: OutlookOkan Performance Code Review

## 1. Objective
Conduct a comprehensive ADVERSARIAL code review focusing on performance, algorithm optimization, and execution speed (speed turning) for the `OutlookOkan` project.

## 2. Orchestration Strategy (2-Phase)
### Phase 1: Planning (Current)
- Establish this `PLAN.md`.
- Create a specific "Performance Optimization" Story file to drive the `code-review` workflow.
- **Agent:** `project-planner` (Acting).

### Phase 2: Execution (Code Review)
- **Workflow:** `@[/bmad-bmm-workflows-code-review]`
- **Agents:**
  1. `backend-specialist`: To analyze C# VSTO logic, LINQ usage, and Interop calls.
  2. `performance-optimizer` (Simulated): To focus purely on algo complexity (Big O) and latency.
  3. `test-engineer`: To verify findings and run existing tests.

## 3. Scope of Review
- **Targets:**
  - `OutlookOkan\ThisAddIn.cs` (Entry points)
  - `OutlookOkan\Models\GenerateCheckList.cs` (Core Logic)
  - `OutlookOkan\CsvTools.cs` (Data I/O)
- **Goals:**
  - Identify blocking calls on UI thread.
  - Optimize loop logic (foreach vs for, LINQ overhead).
  - Reduce VSTO COM Interop calls (marshaling cost).

## 4. Phase 2: Comprehensive Speed Turning (Current)
- **Objective:** System-wide turning for maximum speed.
- **Workflow:** `@[/bmad-bmm-workflows-code-review]`
- **Lead Agent:** `backend-specialist`
- **Scope:**
    - `Handlers/CsvFileHandler.cs`: Optimize memory usage (HashSet).
    - `ViewModels/ConfirmationWindowViewModel.cs`: Optimize UI binding and initialization.
    - `Helpers/*.cs`: Resource handling and static lookups.
    - `Types/*.cs`: Data structure efficiency.

## 5. Verification
- Run `code-review` workflow to generate Findings Report.
- Manual verification of identified hotspots.
