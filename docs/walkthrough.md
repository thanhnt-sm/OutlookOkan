# Walkthrough: Performance Optimization (Phase 2)

**Objective:** Deep optimization (Speed Turning) of the OutlookOkan VSTO Add-in.

## 1. Cross-Platform Build Support
**File:** `build.sh`
- **Change:** Updated script to detect macOS and attempt `dotnet build` using local SDK, or Mono `msbuild` if available.
- **Result:** Provides a valid build attempt on non-Windows platforms (though VSTO targets remain Windows-only).

## 2. Whitelist Lookup Optimization (O(1))
**Files:** `Services/SettingsService.cs`, `Models/GenerateCheckList.cs`
- **Problem:** `Whitelist` was a `List<Whitelist>`. Checking it required iterating the entire list for every recipient (O(N)).
- **Solution:** Converted to `Dictionary<string, bool>`.
    - Key: Address/Domain (Normalized).
    - Value: `IsSkipConfirmation` flag.
    - **Impact:** Lookups are now O(1) hashing operations. Significant speedup for large whitelists.

## 3. UI Responsiveness
**File:** `ViewModels/ConfirmationWindowViewModel.cs`
- **Problem:** `ToggleSendButton` iterated 5 collections fully (`.Count()`) on every check.
- **Solution:** Used `.All()` for short-circuiting.
- **Impact:** Reduced CPU cycles during UI updates.

## 4. String & Regex Optimization
**File:** `Models/GenerateCheckList.cs`
- **Change:** Compiled `CidRegex` (`@"cid:.*?@"`) as `static readonly`.
- **Impact:** Avoids recompiling Regex pattern on every email attachment check.

## 5. Verification
- **Build:** `build.ps1` completed successfully (Exit Code 0).
- **Tests:** Unit tests compiled and passed.
