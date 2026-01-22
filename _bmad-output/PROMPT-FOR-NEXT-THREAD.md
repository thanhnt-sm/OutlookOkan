# Prompt for Next Thread - STORY-001 Task 5 Complete Execution
**Prepared:** 2026-01-22  
**Type:** BMad Master Orchestration with Full Workflow  
**Purpose:** Execute Task 5 with COMPLETE validation & verification

---

## 🎯 **COPY & PASTE THIS EXACT PROMPT IN NEW THREAD**

```
sử dụng agent @_bmad\core\agents\bmad-master.md để điều phối hoàn thành STORY-001 Task 5

════════════════════════════════════════════════════════════════════

TASK: String Replacement Optimization (Task 5 of STORY-001)

CURRENT STORY STATUS:
  • Progress: 67% complete (4 of 6 tasks done)
  • Task 4: WordEditor Hack Optimization (just completed)
  • Task 5: String Replacement (THIS TASK)
  • Task 6: Whitelist (already done)

SCOPE: Optimize all string replacement operations in GenerateCheckList.cs
  ├─ Audit all .Replace() calls
  ├─ Apply compiled Regex for repetitive patterns
  ├─ Use StringBuilder for string concatenation in loops
  └─ Measure and document performance improvement

ESTIMATED EFFORT: 1 hour

════════════════════════════════════════════════════════════════════

REQUIRED DOCUMENTS:
  1. @_bmad-output\TASK-005-PREPARATION.md (complete prep guide)
  2. docs\PERFORMANCE_REVIEW_FINDINGS.md (section 3: string allocations)
  3. OutlookOkan\Models\GenerateCheckList.cs (target file)
  4. OutlookOkan\Models\GenerateCheckList.cs:1055 (example: CidRegex pattern)

════════════════════════════════════════════════════════════════════

WORKFLOW PHASES (MANDATORY - EXECUTE IN ORDER):

Phase 1: AUDIT (20 minutes)
  ├─ Search: Find all .Replace() calls in GenerateCheckList.cs
  ├─ Document: Location, pattern, frequency, impact
  ├─ Identify: Repetitive patterns suitable for compiled Regex
  ├─ Identify: String concatenation in loops (StringBuilder candidate)
  └─ Output: Audit report with findings

Phase 2: IMPLEMENTATION (30 minutes)
  ├─ Add: Compiled Regex constants for repetitive patterns
  │  Example: private static readonly Regex pattern = new Regex(..., RegexOptions.Compiled);
  ├─ Replace: All matching .Replace() calls with compiled Regex
  ├─ Implement: StringBuilder for string building in loops
  ├─ Verify: Code syntax valid after changes
  └─ Output: Code changes with [OPTIMIZATION-TASK5] markers

Phase 3: TESTING & VALIDATION (5 minutes)
  ├─ Check: Existing unit tests still pass
  ├─ Verify: No breaking changes to method signatures
  ├─ Confirm: String output remains identical
  ├─ Validate: Exception handling still in place
  └─ Output: Test validation report

Phase 4: PERFORMANCE MEASUREMENT (5 minutes)
  ├─ Benchmark: Before/after execution time (if possible)
  ├─ Estimate: Memory allocation reduction
  ├─ Calculate: Annual impact per user & per 1000 users
  ├─ Compare: Against TASK-004 improvements
  └─ Output: Performance measurement report

════════════════════════════════════════════════════════════════════

ACCEPTANCE CRITERIA (ALL MUST BE MET):

Code Quality:
  ☐ All .Replace() calls in GenerateCheckList.cs identified
  ☐ Compiled Regex created for repetitive patterns
  ☐ StringBuilder used for string concatenation in loops
  ☐ Code syntax valid (no compile errors)
  ☐ [OPTIMIZATION-TASK5] markers added to all changes
  ☐ Comments explain WHY optimization applied

Performance:
  ☐ 5-10% improvement measured (string operation time)
  ☐ Memory allocations reduced (quantified)
  ☐ No performance regression
  ☐ Benchmarks documented (before/after)

Correctness:
  ☐ Zero breaking changes
  ☐ String output identical to original code
  ☐ No API signature changes
  ☐ Exception handling preserved
  ☐ All edge cases handled

Documentation:
  ☐ Audit report created (Phase 1 output)
  ☐ Implementation report created (Phase 2 output)
  ☐ Test validation report created (Phase 3 output)
  ☐ Performance measurement report created (Phase 4 output)
  ☐ Code changes fully documented

════════════════════════════════════════════════════════════════════

MANDATORY VERIFICATION (BEFORE COMPLETION):

1. CODE AUDIT:
   ├─ List all locations modified
   ├─ Show before/after code for each location
   ├─ Verify syntax correctness
   └─ Confirm no missing implementations

2. BUILD VALIDATION:
   ├─ Check OutlookOkan.sln builds without new errors
   ├─ Verify no unresolved references
   ├─ Confirm no compiler warnings from changes
   └─ Validate assembly integrity

3. FUNCTIONAL TESTING:
   ├─ Existing unit tests pass
   ├─ String operations produce same results
   ├─ Edge cases handled correctly
   └─ Exception handling works

4. PERFORMANCE VALIDATION:
   ├─ Memory usage reduced (measurements)
   ├─ String operation faster (benchmarks)
   ├─ Annual impact calculated
   ├─ Compared with Task 4 impact
   └─ Overall STORY-001 impact updated

════════════════════════════════════════════════════════════════════

OUTPUT DELIVERABLES (REQUIRED):

1. TASK-005-AUDIT-REPORT.md
   ├─ All .Replace() calls found and documented
   ├─ Repetitive patterns identified
   ├─ StringBuilder opportunities found
   └─ Summary with location references

2. TASK-005-IMPLEMENTATION-REPORT.md
   ├─ Code changes made (before/after for each location)
   ├─ Compiled Regex patterns created
   ├─ StringBuilder implementations added
   ├─ All changes marked with [OPTIMIZATION-TASK5]
   └─ Syntax validation results

3. TASK-005-TEST-VALIDATION-REPORT.md
   ├─ Unit test results
   ├─ String output verification
   ├─ Breaking changes check
   └─ Edge cases verification

4. TASK-005-PERFORMANCE-REPORT.md
   ├─ Benchmark results (before/after)
   ├─ Memory reduction quantified
   ├─ Annual impact per user
   ├─ Annual impact per 1000 users
   ├─ Comparison with Task 4
   └─ Overall STORY-001 impact updated

5. TASK-005-COMPLETION-REPORT.md
   ├─ Official completion status
   ├─ All 4 phases verified
   ├─ All acceptance criteria met
   ├─ Code ready for production
   └─ Lessons learned

════════════════════════════════════════════════════════════════════

VALIDATION REQUIREMENTS:

❌ DO NOT accept incomplete code
❌ DO NOT accept shell code (functions without implementation)
❌ DO NOT accept theoretical improvements without measurement
❌ DO NOT accept missing documentation
❌ DO NOT accept unverified claims
❌ DO NOT report completion if verification failed

✅ REQUIRE complete, working, tested code
✅ REQUIRE before/after code comparison
✅ REQUIRE actual performance measurements
✅ REQUIRE comprehensive documentation
✅ REQUIRE verification of all acceptance criteria
✅ REQUIRE honest assessment of what was actually done

════════════════════════════════════════════════════════════════════

SPECIAL INSTRUCTIONS FOR BMAD MASTER:

1. WORKFLOW ORCHESTRATION:
   ├─ Use dev-story workflow from BMM for implementation
   ├─ Use code-review workflow for validation
   ├─ Apply quick-dev workflow if any issues arise
   └─ Document all workflow steps used

2. AGENT COORDINATION:
   ├─ Coordinate with development agent for implementation
   ├─ Coordinate with review agent for validation
   ├─ Escalate to architect if design issues found
   └─ Report all coordination steps

3. QUALITY GATES:
   ├─ NO code committed until Phase 3 tests pass
   ├─ NO performance claims without Phase 4 measurements
   ├─ NO completion reported until ALL acceptance criteria met
   ├─ NO shortcuts on validation steps
   └─ Quality gates must be met 100%, not 90%

4. REPORTING HONESTY:
   ├─ Report actual measurements, not estimates
   ├─ Report actual code changes, not claimed changes
   ├─ Report actual test results, not assumed results
   ├─ Report actual problems encountered and how resolved
   └─ NO false positives or misleading completion claims

════════════════════════════════════════════════════════════════════

AFTER TASK 5 COMPLETION:

Once all phases complete and verified:
  1. Generate final TASK-005-COMPLETION-REPORT.md
  2. Update EXECUTION-SUMMARY.md with Task 5 results
  3. Mark Task 6 as verified complete (already done)
  4. Prepare for STORY-001 final closure

Then ask user: "Task 5 complete. Proceed to generate STORY-001 final report?"

════════════════════════════════════════════════════════════════════

REFERENCE DOCUMENTS:
  • @_bmad-output\TASK-005-PREPARATION.md
  • @_bmad-output\TASK-004-FINAL-SUMMARY.md (previous task pattern)
  • docs\PERFORMANCE_REVIEW_FINDINGS.md (original analysis)
  • OutlookOkan\Models\GenerateCheckList.cs (target code)

BEGIN: Execute Phase 1 (AUDIT) now
```

---

## 📝 **How to Use This Prompt**

### **In New Amp Thread:**

1. **Create new thread** in Amp
2. **Paste the entire prompt above** (from "sử dụng agent..." to "...now")
3. **Send to agent**
4. **Monitor execution** - agent will report progress for each phase
5. **Verify completeness** - check all acceptance criteria met

---

## ✅ **Key Safety Features in This Prompt**

**Anti-Shortcut Measures:**
- ✅ 4 phases MUST execute in order
- ✅ Each phase has specific output requirements
- ✅ Mandatory verification gates
- ✅ Explicit "DO NOT accept" conditions
- ✅ Explicit "REQUIRE" complete conditions

**Quality Enforcement:**
- ✅ Code must pass unit tests
- ✅ Build must succeed
- ✅ Measurements must be actual, not estimated
- ✅ Performance claims must be verified
- ✅ Completion only when ALL criteria met

**Honest Reporting:**
- ✅ Report actual code, not claimed code
- ✅ Report actual measurements, not estimates
- ✅ Report problems encountered and solutions
- ✅ No false positives
- ✅ No misleading completion claims

---

## 🎯 **What This Prompt Ensures**

| Aspect | How Ensured |
|--------|------------|
| **Complete Code** | Phase 2 requires full implementation, Phase 3 tests it |
| **No Shell Code** | Phase 1 audit + Phase 2 implementation both documented |
| **Verified Work** | Phase 3 validation + Phase 4 measurement both mandatory |
| **Honest Reporting** | "Validation Requirements" section explicitly forbids false claims |
| **Quality Gates** | Each phase must complete before next, verification gates enforced |

---

## 📊 **Expected Outcome**

After agent completes this prompt:

✅ **Phase 1:** Audit report showing all string operations found  
✅ **Phase 2:** Code changes with before/after comparison  
✅ **Phase 3:** Test validation confirming no breaks  
✅ **Phase 4:** Performance measurement with actual numbers  
✅ **Final:** Completion report with all criteria verified  

**Status:** STORY-001 = 83% complete (ready for Task 6 closure + final report)

---

**Prepared By:** BMad Master Executor  
**Date:** 2026-01-22  
**Status:** ✅ READY TO USE IN NEXT THREAD
