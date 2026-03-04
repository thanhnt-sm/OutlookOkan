# OutlookOkan Resiliency Fix (v2.8.3)

## Problem

OutlookOkan VSTO add-in tự động bị Outlook disable sau vài ngày sử dụng.
User phải vào **File → Options → Add-Ins → COM Add-ins** tích chọn lại OutlookOkan.

## Root Cause Analysis

Outlook có cơ chế **Resiliency Manager** giám sát tất cả add-in:

| Trigger | Outlook Action |
|---------|---------------|
| Add-in crash (unhandled exception) | Ghi vào `CrashingAddinList`, disable nếu lặp lại |
| Add-in gây slow startup (>1000ms) | Cảnh báo user, có thể disable |
| Memory leak từ COM objects | Ghi nhận unstable, disable |

### 5 Lỗi Cụ Thể

| # | Lỗi | Impact |
|---|------|--------|
| 1 | Không có `ThisAddIn_Shutdown` | COM objects không bao giờ được release → memory leak |
| 2 | `CurrentExplorer_SelectionChange` không có try-catch | Bất kỳ exception nào → crash → Resiliency disable |
| 3 | `OpenOutboxItemInspector` không có try-catch | Crash khi mở Outbox item → Resiliency disable |
| 4 | Không có Resiliency Registry protection | Outlook tự do disable mà không bị ngăn cản |
| 5 | Installer không reset Resiliency keys | Cài lại vẫn bị disable do keys cũ |

## Solution

### 1. ThisAddIn_Shutdown (COM Cleanup)

```csharp
private void ThisAddIn_Shutdown(object sender, EventArgs e)
{
    // Unregister events
    Application.ItemSend -= Application_ItemSend;
    _inspectors.NewInspector -= OpenOutboxItemInspector;
    _currentExplorer.SelectionChange -= CurrentExplorer_SelectionChange;
    _currentMailItem.BeforeAttachmentRead -= BeforeAttachmentRead;

    // Release COM objects
    SafeReleaseCom(_currentMailItem);
    SafeReleaseCom(_mapiNamespace);
    SafeReleaseCom(_currentExplorer);
    SafeReleaseCom(_inspectors);
}
```

### 2. Try-Catch Protection

Tất cả event handlers (`CurrentExplorer_SelectionChange`, `OpenOutboxItemInspector`) được wrap trong try-catch:
- Ngăn unhandled exceptions crash add-in
- Log lỗi qua `Debug.WriteLine` để debug khi cần

### 3. Resiliency Registry Protection

Mỗi khi add-in startup, ghi registry keys:

```
HKCU\Software\Microsoft\Office\{16.0,15.0}\Outlook\Resiliency\
├── DoNotDisableAddinList\OutlookOkan = 1  (DWORD)
├── CrashingAddinList  → xóa entries của OutlookOkan
└── DisabledItems      → xóa entries của OutlookOkan
```

### 4. Installer Reset

`CustomAction.cs` gọi `ResetResiliencyKeys()` khi install:
- Ghi `DoNotDisableAddinList`
- Xóa toàn bộ `CrashingAddinList` và `DisabledItems` keys

## Files Changed

| File | Lines | Description |
|------|-------|-------------|
| `OutlookOkan/ThisAddIn.cs` | +198 | Shutdown, try-catch, resiliency protection |
| `SetupCustomAction/CustomAction.cs` | +49 | Installer resiliency reset |
| `version` | 2.8.2 → 2.8.3 | Version bump |
| `OutlookOkan.csproj` | 2.8.2 → 2.8.3 | Version bump |

## Verification

After deploying v2.8.3:

1. **Check Registry**: Open `regedit` →
   `HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList`
   → Confirm `OutlookOkan` = 1

2. **Check Add-in Active**: Outlook → File → Options → Add-Ins
   → OutlookOkan should be in "Active Application Add-ins"

3. **Long-term Test**: Use Outlook normally for several days
   → Add-in should remain active without manual re-enabling
