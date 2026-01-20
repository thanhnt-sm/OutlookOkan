# Kiến Trúc OutlookOkan

## Tổng Quan

**OutlookOkan** (おかん for Outlook) là một VSTO Add-in cho Microsoft Outlook, được phát triển bằng C#/.NET Framework 4.6.2. Mục đích chính là **ngăn ngừa gửi email nhầm** bằng cách hiển thị cửa sổ xác nhận trước khi gửi email.

> [!NOTE]
> **VSTO** (Visual Studio Tools for Office) cho phép tích hợp sâu với Outlook thông qua COM interop.

---

## Cấu Trúc Thư Mục

```
OutlookOkan/
├── 📁 Handlers/          # Xử lý file (CSV, Mail Header, Office, PDF, ZIP)
├── 📁 Helpers/           # Native methods helper
├── 📁 Models/            # Business logic chính
├── 📁 Properties/        # Resources & Settings
├── 📁 Services/          # Dịch vụ (đa ngôn ngữ)
├── 📁 Types/             # Data models (29 types)
├── 📁 ViewModels/        # MVVM ViewModels
├── 📁 Views/             # WPF Windows
├── 📄 ThisAddIn.cs       # Entry point (858 dòng)
└── 📄 Ribbon.cs          # Ribbon integration
```

---

## Kiến Trúc Tổng Quan

```mermaid
graph TB
    subgraph Outlook["Microsoft Outlook"]
        OE[Outlook Events]
    end
    
    subgraph AddIn["OutlookOkan Add-in"]
        TA[ThisAddIn<br/>Entry Point]
        GC[GenerateCheckList<br/>Core Logic]
        
        subgraph UI["UI Layer - MVVM"]
            CW[ConfirmationWindow]
            SW[SettingsWindow]
            AW[AboutWindow]
            VM[ViewModels]
        end
        
        subgraph Data["Data Layer"]
            CSV[CsvFileHandler]
            Types[Types/Models]
        end
        
        subgraph Handlers["Handlers"]
            MH[MailHeaderHandler]
            OF[OfficeFileHandler]
            ZF[ZipFileHandler]
            PDF[PdfFileHandler]
        end
    end
    
    OE --> TA
    TA --> GC
    GC --> CSV
    GC --> Types
    GC --> Handlers
    TA --> CW
    CW --> VM
    VM --> Types
    
    style TA fill:#ff6b6b,color:#fff
    style GC fill:#4ecdc4,color:#fff
    style CW fill:#45b7d1,color:#fff
```

---

## Các Thành Phần Chính

### 1. Entry Point - `ThisAddIn.cs`

**Chức năng**: Điểm vào chính của Add-in, xử lý các sự kiện từ Outlook.

| Event Handler | Mô Tả |
|:-------------|:------|
| `ThisAddIn_Startup` | Khởi tạo Add-in, load settings |
| `Application_ItemSend` | **Quan trọng nhất** - Chặn gửi email để kiểm tra |
| `CurrentExplorer_SelectionChange` | Phân tích email đã chọn |
| `BeforeAttachmentRead` | Cảnh báo trước khi mở attachment |

### 2. Core Logic - `GenerateCheckList.cs`

**Chức năng**: Xử lý business logic chính (2383 dòng code).

```mermaid
flowchart TD
    A[Nhận MailItem] --> B[Lấy Sender & Domain]
    B --> C[Kiểm tra Recipients]
    C --> D[Kiểm tra Attachments]
    D --> E[Kiểm tra Keywords]
    E --> F[Kiểm tra Whitelist]
    F --> G{Có vi phạm?}
    G -->|Có| H[Tạo Alerts]
    G -->|Không| I[Đánh dấu Checked]
    H --> J[Trả về CheckList]
    I --> J
```

**Các phương thức chính:**

| Method | Chức năng |
|:-------|:---------|
| `GenerateCheckListFromMail()` | Phương thức chính, tạo CheckList từ email |
| `GetSenderAndSenderDomain()` | Lấy thông tin người gửi |
| `GetNameAndRecipient()` | Phân tích danh sách người nhận |
| `CountRecipientExternalDomains()` | Đếm domain bên ngoài |

### 3. UI Layer - Views & ViewModels

**Pattern**: MVVM (Model-View-ViewModel)

```mermaid
classDiagram
    class ConfirmationWindow {
        +ShowDialog()
        -DataContext: ConfirmationWindowViewModel
    }
    class ConfirmationWindowViewModel {
        +CheckList CheckList
        +bool CanSend
        +ICommand SendCommand
        +ICommand CancelCommand
    }
    class SettingsWindow {
        +ShowDialog()
    }
    class SettingsWindowViewModel {
        +GeneralSetting Settings
        +ICommand SaveCommand
        +ICommand ImportCommand
        +ICommand ExportCommand
    }
    
    ConfirmationWindow --> ConfirmationWindowViewModel
    SettingsWindow --> SettingsWindowViewModel
```

### 4. Handlers

| Handler | Chức năng |
|:--------|:---------|
| `CsvFileHandler` | Đọc/ghi settings từ CSV files |
| `MailHeaderHandler` | Phân tích SPF, DKIM, DMARC |
| `OfficeFileHandler` | Kiểm tra macro trong Office files |
| `ZipFileHandler` | Kiểm tra ZIP có mã hóa/lnk files |
| `PdfFileHandler` | Xử lý PDF files |

---

## Luồng Xử Lý Gửi Email

```mermaid
sequenceDiagram
    actor User as Người dùng
    participant Outlook as Outlook
    participant TA as ThisAddIn
    participant GC as GenerateCheckList
    participant CW as ConfirmationWindow
    
    User->>Outlook: Click "Send"
    Outlook->>TA: Application_ItemSend()
    
    Note over TA: Load Settings từ CSV
    
    TA->>GC: GenerateCheckListFromMail()
    
    activate GC
    GC->>GC: Kiểm tra Recipients
    GC->>GC: Kiểm tra Attachments
    GC->>GC: Kiểm tra Keywords
    GC-->>TA: CheckList object
    deactivate GC
    
    alt Có lỗi nghiêm trọng (IsCanNotSendMail)
        TA->>User: Hiển thị thông báo lỗi
        TA->>Outlook: cancel = true
    else Cần xác nhận
        TA->>CW: ShowDialog(CheckList)
        CW->>User: Hiển thị cửa sổ xác nhận
        
        alt User chọn OK (sau khi check hết)
            CW-->>TA: true
            TA->>Outlook: Cho phép gửi
        else User chọn Cancel
            CW-->>TA: false
            TA->>Outlook: cancel = true
        end
    else Không cần xác nhận (Whitelist)
        TA->>Outlook: Cho phép gửi
    end
```

---

## Cấu Hình

Settings được lưu trữ dưới dạng **CSV files** tại:
```
%APPDATA%\Noraneko\OutlookOkan\
```

| File | Mô Tả |
|:-----|:------|
| `GeneralSetting.csv` | Cài đặt chung |
| `Whitelist.csv` | Danh sách cho phép |
| `InternalDomainList.csv` | Domain nội bộ |
| `AlertKeywordAndMessageList.csv` | Từ khóa cảnh báo |
| `AutoCcBccRecipientList.csv` | Tự động CC/BCC |
| `DeferredDeliveryMinutesList.csv` | Gửi trễ |

---

## Bảo Mật

### Phân Tích Email Nhận (Received Mail Security)

```mermaid
graph LR
    A[Email nhận] --> B{Phân tích Header}
    B --> C[SPF Check]
    B --> D[DKIM Check]
    B --> E[DMARC Check]
    C --> F{Kết quả}
    D --> F
    E --> F
    F -->|FAIL| G[Hiển thị cảnh báo]
    F -->|PASS| H[Không cảnh báo]
```

### Kiểm Tra Attachment

- ✅ ZIP có mã hóa
- ✅ File .lnk trong ZIP
- ✅ File .one (OneNote) trong ZIP
- ✅ Macro trong Office files (.docm, .xlsm, .pptm)

---

## Dependencies

| Package | Version | Mô Tả |
|:--------|:--------|:------|
| CsvHelper | 15.0.5 | Đọc/ghi CSV |
| Microsoft.Office.Interop.Outlook | 15.0.4797.1003 | Outlook COM |
| Microsoft.Office.Interop.Word | 15.0.4797.1003 | Word COM |
| SharpCompress | 0.37.2 | Xử lý ZIP |

---

## Đa Ngôn Ngữ

Add-in hỗ trợ **10 ngôn ngữ** thông qua `ResourceService`:

- 🇯🇵 Tiếng Nhật (mặc định)
- 🇺🇸 Tiếng Anh
- 🇨🇳 Tiếng Trung (Giản thể & Phồn thể)
- 🇰🇷 Tiếng Hàn
- 🇩🇪 Tiếng Đức
- 🇫🇷 Tiếng Pháp
- 🇪🇸 Tiếng Tây Ban Nha
- 🇵🇹 Tiếng Bồ Đào Nha
- 🇮🇹 Tiếng Ý
