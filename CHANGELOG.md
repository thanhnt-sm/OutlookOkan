# Changelog

All notable changes to OutlookOkan will be documented in this file.

## [2.8.3] - 2026-03-04

### Fixed
- **Add-in Auto-Disable**: Fixed issue where Outlook automatically disables OutlookOkan after a few days of usage
  - Root cause: Outlook Resiliency Manager detects unhandled exceptions and COM memory leaks, then disables the add-in
  - Added `ThisAddIn_Shutdown` handler to properly release COM objects (`_currentExplorer`, `_inspectors`, `_mapiNamespace`, `_currentMailItem`) and unregister event handlers when Outlook closes
  - Wrapped `CurrentExplorer_SelectionChange` and `OpenOutboxItemInspector` event handlers in try-catch to prevent crashes from triggering Resiliency disable
  - Added `EnsureResiliencyProtection()` at startup: writes `DoNotDisableAddinList` registry key and clears `CrashingAddinList`/`DisabledItems` entries
  - Added `ResetResiliencyKeys()` in installer `CustomAction` to clear disable history on install/reinstall

### Technical Details
- **Registry keys modified** (HKCU):
  - `Software\Microsoft\Office\{16.0,15.0}\Outlook\Resiliency\DoNotDisableAddinList\OutlookOkan` = 1
  - Clears: `CrashingAddinList`, `DisabledItems`
- **Files changed**:
  - `OutlookOkan/ThisAddIn.cs` (+198 lines)
  - `SetupCustomAction/CustomAction.cs` (+49 lines)

## [2.8.2] - Previous Release

### Notes
- Previous stable release before resiliency fix
