# 🔥 Performance Code Review Findings (Phase 2)
**Reviewer:** Antigravity (Backend Specialist)
**Target:** Full Codebase (Handlers, ViewModels, Services)
**Focus:** Speed Turning & Micro-optimizations

## 🔴 CRITICAL (Algorithm & Complexity)

### 1. O(n) Lookup in Whitelists (`CsvFileHandler` & `SettingsService`)
- **Location:** `Handlers/CsvFileHandler.cs`, `Services/SettingsService.cs`
- **Problem:** `ReadCsv` returns `List<T>`. All Whitelist checks (thousands of iterations) use `List.Contains` or `IEnumerable.Any`, which are O(n).
- **Impact:** Exponential slowdown as whitelist size grows.
- **Fix:** Refactor `SettingsService` to store Whitelists as `HashSet<string>`. Change `ReadCsv` to return `IEnumerable` to allow immediate `.ToHashSet()` projection.

### 2. Redundant Iterations in UI Logic (`ConfirmationWindowViewModel`)
- **Location:** `ViewModels/ConfirmationWindowViewModel.cs` -> `ToggleSendButton`
- **Problem:**
  ```csharp
  var isToAddressesCompleteChecked = ToAddresses.Count(x => x.IsChecked) == ToAddresses.Count;
  // Repeated 5 times for different lists
  ```
  This iterates the *entire* list 5 separate times on every UI toggle.
- **Fix:** Use `.All(x => x.IsChecked)`. It returns `false` immediately upon finding the first unchecked item (Short-circuiting).

## 🟡 MEDIUM (Memory & GC)

### 3. String Allocations in CheckList Logic
- **Location:** `Models/GenerateCheckList.cs` (Line 337: `Replace("\r\n\r\n", "\r\n")`)
- **Problem:** Creating new string objects for simple replacements on large bodies.
- **Fix:** Use `Regex` (Compiled).

### 4. Heavy LINQ in Domain Counting
- **Location:** `GenerateCheckList.cs` -> `CountRecipientExternalDomains`
- **Problem:**
  ```csharp
  displayNameAndRecipient.To.Select(...).Where(...).Substring(...)
  ```
  Creates intermediate iterators.
- **Fix:** Combine validation logic.

## 🟢 LOW (Micro-optimizations)

### 5. ObservableCollection Batching
- **Location:** `ConfirmationWindowViewModel.cs`
- **Problem:** Adding items one-by-one fires events repeatedly.
- **Fix:** initialize with list directly.
