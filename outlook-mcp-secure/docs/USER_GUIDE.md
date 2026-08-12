# Hướng Dẫn Sử Dụng Claude-Outlook MCP

> Dành cho người dùng không chuyên kỹ thuật.
> Phiên bản server: 2.1.0 — Cập nhật: 2026-06-24

---

## 1. Giới Thiệu — Claude-Outlook MCP Là Gì, Làm Được Gì

**Claude-Outlook MCP** là một cầu nối (bridge) cho phép Claude Code CLI (công cụ chat AI của Anthropic trên dòng lệnh) đọc và làm việc với hộp thư Outlook Desktop của bạn ngay trong cửa sổ chat.

Thay vì mở Outlook rồi tự tay tìm kiếm email, bạn có thể gõ yêu cầu thẳng vào Claude bằng tiếng Việt tự nhiên — ví dụ "tóm tắt email từ sếp tuần này" — và Claude sẽ tự động đọc, tìm kiếm, tóm tắt cho bạn.

### Những gì công cụ này làm được

- **Đọc email** — xem danh sách email, đọc nội dung đầy đủ của từng email
- **Tìm kiếm email** — tìm theo từ khóa, theo người gửi, theo ngày, theo tiêu đề
- **Liệt kê thư mục** — xem toàn bộ cấu trúc thư mục Outlook của bạn
- **Soạn email nháp** (tùy chọn, mặc định tắt) — mở sẵn cửa sổ soạn thảo với nội dung đã điền, bạn xem lại rồi tự nhấn Gửi
- **Trả lời email nháp** (tùy chọn, mặc định tắt) — soạn sẵn câu trả lời, bạn kiểm tra rồi tự nhấn Gửi

### Những gì công cụ này KHÔNG làm

- **Không tự động gửi email** — mọi email đều cần bạn nhấn Gửi thủ công trong Outlook
- **Không xóa, không di chuyển, không đánh dấu** email
- **Không truy cập thư mục ngoài danh sách cho phép** (xem mục 7)
- **Không chạy khi Outlook đóng** — Outlook Desktop phải đang mở trước

---

## 2. Yêu Cầu

Trước khi cài đặt, đảm bảo máy tính của bạn đáp ứng các điều kiện sau:

| Yêu cầu | Chi tiết |
|---|---|
| **Python 3.11 trở lên** | Ngôn ngữ lập trình nền tảng của server. Tải tại [python.org](https://python.org) |
| **Windows 10/11** | Công cụ chỉ chạy trên Windows (dùng COM — công nghệ tích hợp phần mềm của Windows) |
| **Microsoft Outlook Desktop** | Phiên bản Outlook cài đặt trên máy (không hỗ trợ Outlook Web) — phải đang **mở và đăng nhập** khi dùng |
| **Claude Code CLI** | Công cụ chat AI trên dòng lệnh. Tải tại [claude.ai/download](https://claude.ai/download) |
| **Kết nối Internet** | Cần thiết cho Claude Code CLI để giao tiếp với Anthropic |

### Kiểm tra phiên bản Python

Mở PowerShell hoặc Command Prompt, gõ:

```
python --version
```

Kết quả phải là `Python 3.11.x` trở lên. Nếu thấy phiên bản cũ hơn, hãy tải Python mới tại [python.org/downloads](https://python.org/downloads).

---

## 3. Cài Đặt — Từng Bước Chạy setup.ps1

Script `setup.ps1` tự động cài đặt toàn bộ môi trường cho bạn — **không cần quyền Admin**.

### Bước 1: Mở PowerShell

Nhấn phím `Windows`, gõ `PowerShell`, rồi nhấn Enter.

### Bước 2: Di chuyển đến thư mục dự án

```powershell
cd "D:\100.Software\Github\OutlookOkan\outlook-mcp-secure"
```

*(Thay đường dẫn trên bằng vị trí thực tế bạn đã lưu thư mục dự án.)*

### Bước 3: Cho phép chạy script PowerShell

Nếu máy chưa từng chạy script `.ps1`, chạy lệnh sau một lần:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Nhập `Y` khi được hỏi xác nhận.

### Bước 4: Chạy script cài đặt

```powershell
.\setup.ps1
```

Script sẽ tự động thực hiện 7 bước:

1. **Kiểm tra Python** — đảm bảo phiên bản đủ yêu cầu
2. **Tạo môi trường ảo** (virtual environment) tại thư mục `.\venv\` — ngăn cách các thư viện của dự án này với máy tính
3. **Nâng cấp pip** — công cụ quản lý thư viện Python
4. **Cài đặt thư viện** từ file `requirements.txt`
5. **Cài đặt pywin32** — thư viện kết nối Python với Outlook
6. **Hỏi lưu API Key** — nếu chọn "y", wizard sẽ hướng dẫn lưu Anthropic API Key vào Windows Credential Manager (kho lưu mật khẩu bảo mật của Windows, **hoàn toàn không lưu vào file**)
7. **Hiển thị bước tiếp theo** — hướng dẫn thêm server vào Claude

### Bước 5: Tạo file cấu hình

Nếu chưa có file `config.toml` trong thư mục dự án, chạy:

```powershell
.\venv\Scripts\python.exe server.py --setup
```

Wizard sẽ tự tạo file `config.toml` từ mẫu có sẵn. Mở file đó và điền tên tài khoản Outlook của bạn (xem mục 7 để biết chi tiết cách chỉnh).

---

## 4. Kết Nối Claude — Cách Thêm MCP Server Vào Claude Code CLI

Sau khi cài đặt xong, bạn cần báo cho Claude Code CLI biết server này tồn tại.

### Cách A — Dùng lệnh (khuyến nghị, đơn giản nhất)

Mở PowerShell trong thư mục dự án, chạy:

```powershell
claude mcp add outlook -- .\venv\Scripts\python.exe server.py
```

Lệnh này tự động đăng ký server với tên `outlook` vào Claude Code CLI. Chỉ cần chạy **một lần duy nhất**.

### Cách B — Chỉnh tay file cấu hình

Nếu Cách A không hoạt động, bạn có thể chỉnh tay:

1. Mở file `claude-mcp.json` trong thư mục dự án
2. Thay chữ `ABSOLUTE_PATH_TO_THIS_DIR` bằng đường dẫn thực tế đến thư mục dự án (ví dụ: `D:\\100.Software\\Github\\OutlookOkan\\outlook-mcp-secure`)
3. Mở file `C:\Users\<tên_user>\.claude\config.json`
4. Thêm nội dung phần `"mcpServers"` từ `claude-mcp.json` vào file config đó

### Kiểm tra kết nối

Mở Outlook Desktop trước, rồi mở Claude Code CLI và gõ:

```
Liệt kê các thư mục email Outlook của tôi
```

Nếu Claude trả về danh sách thư mục — kết nối thành công.

---

## 5. Cách Dùng — Ví Dụ Câu Hỏi Thực Tế

Sau khi kết nối thành công, bạn dùng Claude bình thường — chỉ cần gõ yêu cầu bằng tiếng Việt tự nhiên.

> **Lưu ý quan trọng:** Outlook Desktop phải đang **mở** trước khi gõ bất kỳ yêu cầu nào liên quan đến email.

### Đọc email mới nhất

```
Đọc 10 email mới nhất trong Inbox
```

```
Cho tôi xem 5 email chưa đọc trong Hộp thư đến
```

Claude sẽ trả về danh sách email với tiêu đề, người gửi, ngày nhận. Mỗi email có một mã nhận dạng nội bộ (Entry ID) để dùng cho các lệnh tiếp theo.

### Tìm email theo người gửi

```
Tìm email từ nguyen@example.com tuần này
```

```
Tìm tất cả email từ sếp trong tháng 6
```

```
Tìm email có từ khóa "báo cáo quý 2" trong tiêu đề
```

### Đọc nội dung đầy đủ một email cụ thể

Sau khi có danh sách, bạn có thể nhờ Claude đọc chi tiết một email:

```
Đọc nội dung email thứ 3 trong danh sách vừa rồi
```

Hoặc nếu biết Entry ID (mã nhận dạng nội bộ của email, dạng chuỗi ký tự dài):

```
Tóm tắt email có EntryID 00000000ABCD1234...
```

### Tóm tắt và phân tích

```
Tóm tắt 5 email quan trọng nhất trong Inbox hôm nay
```

```
Email nào từ khách hàng chưa được trả lời trong tuần này?
```

```
Liệt kê tất cả email có file đính kèm trong tháng này
```

### Soạn email nháp (chỉ khi đã bật tính năng này trong config.toml)

```
Soạn email cho abc@example.com với tiêu đề "Báo cáo tháng 6" và nội dung: Kính gửi anh/chị, xin gửi báo cáo tháng 6 như đã hẹn...
```

```
Soạn email trả lời email vừa đọc, nội dung: Cảm ơn anh, tôi sẽ xem xét và phản hồi trước thứ Sáu.
```

> Claude sẽ mở cửa sổ soạn thảo trong Outlook với nội dung đã điền sẵn. **Bạn phải tự nhấn nút Gửi** — Claude không bao giờ gửi thay bạn.

---

## 6. Bảo Mật — Dữ Liệu Đi Đâu, Ai Thấy Gì

Phần này giải thích đơn giản cách dữ liệu email của bạn được xử lý.

### Dữ liệu email đi đâu?

Khi bạn nhờ Claude đọc email:

1. Claude Code CLI **gửi yêu cầu** đến MCP server chạy trên máy tính của bạn
2. Server **đọc email từ Outlook** (chạy hoàn toàn trên máy bạn, không qua internet)
3. Nội dung email **được gửi đến Anthropic** (nhà phát triển Claude) qua kết nối mã hóa HTTPS để Claude hiểu và trả lời
4. Anthropic **xử lý và trả về câu trả lời** — không lưu nội dung email vĩnh viễn theo chính sách riêng tư của họ

### AI thấy gì từ email của bạn?

Claude **chỉ thấy** nội dung email trong các thư mục bạn đã cho phép trong cấu hình (xem mục 7). Mặc định chỉ là thư mục Inbox.

### Audit Log là gì?

Mọi lần Claude đọc email đều được ghi vào **nhật ký kiểm toán** (audit log — file ghi lại lịch sử hoạt động) lưu tại:

```
C:\Users\<tên_user>\AppData\Roaming\ClaudeOutlookMCP\audit.jsonl
```

File này ghi **siêu dữ liệu** (metadata — thông tin mô tả, không phải nội dung): "ai gọi tool nào, lúc mấy giờ, trả về bao nhiêu kết quả" — **không ghi nội dung email**.

### API Key được lưu ở đâu?

Anthropic API Key (chìa khóa để dùng dịch vụ Claude) được lưu trong **Windows Credential Manager** — kho mật khẩu bảo mật tích hợp sẵn trong Windows, giống nơi Windows lưu mật khẩu Wi-Fi và tài khoản. **Hoàn toàn không lưu vào file nào trong thư mục dự án.**

### Tóm tắt các biện pháp bảo vệ

- Server chỉ kết nối với Claude qua **cổng chuẩn stdin/stdout** — không mở cổng mạng nào
- Chỉ đọc email trong **danh sách thư mục được phép** — không tự ý đọc thư mục khác
- Mặc định ở **chế độ chỉ đọc** — không thể soạn hay gửi email trừ khi bạn tắt chế độ này
- Không lưu nội dung email vào bất kỳ file nào trên máy bạn

---

## 7. Cấu Hình — Cách Chỉnh file config.toml

File `config.toml` trong thư mục dự án là nơi bạn tùy chỉnh hoạt động của server.

### Vị trí file

```
<thư mục dự án>\config.toml
```

Nếu file chưa có, chạy `.\venv\Scripts\python.exe server.py --setup` để tạo từ mẫu.

### Các cài đặt quan trọng nhất

Mở file bằng Notepad hoặc bất kỳ trình soạn thảo văn bản nào:

#### Đặt tên tài khoản Outlook

```toml
[outlook]
account_name = "your.email@company.com"
```

Thay bằng địa chỉ email hiển thị trong Outlook của bạn. Đây là trường **bắt buộc**.

#### Thêm thư mục vào danh sách cho phép

```toml
[security]
allowed_folders = [
    "Inbox",
    "Hộp thư đến",
    "Sent Items",
    "Projects/Khách hàng A",
]
```

- Chỉ các thư mục có tên trong danh sách này mới được Claude đọc
- Hỗ trợ thư mục con bằng dấu gạch chéo: `"Inbox/Dự án"`
- Hỗ trợ ký tự đại diện một cấp: `"Inbox/*"` (nhưng **không** hỗ trợ `**` đệ quy)
- Hỗ trợ cả tên tiếng Việt và tiếng Anh

#### Bật tính năng soạn email nháp

Mặc định tắt để an toàn. Để bật:

```toml
[security]
read_only_mode = false
```

> **Lưu ý:** Ngay cả khi bật, Claude chỉ mở cửa sổ soạn thảo với nội dung đã điền — bạn vẫn phải tự nhấn Gửi trong Outlook.

#### Kiểm tra cấu hình sau khi chỉnh

Sau khi lưu file, chạy lệnh sau để xác nhận cấu hình không có lỗi:

```powershell
.\venv\Scripts\python.exe config.py
```

Nếu thấy thông báo "Cấu hình hợp lệ!" là thành công.

---

## 8. Troubleshooting — Lỗi Thường Gặp Và Cách Sửa

### Lỗi: "Outlook không đang chạy"

**Triệu chứng:** Claude báo lỗi khi gọi bất kỳ lệnh email nào.

**Nguyên nhân:** Outlook Desktop chưa mở hoặc đang bị treo (không phản hồi).

**Cách sửa:**
1. Mở Outlook Desktop và đăng nhập
2. Đảm bảo Outlook **không đang tải** hay **đang đồng bộ** — chờ đến khi ổn định
3. Thử lại lệnh trong Claude

---

### Lỗi: "Thư mục không được phép truy cập"

**Triệu chứng:** Claude nói thư mục không có trong danh sách cho phép.

**Nguyên nhân:** Thư mục bạn muốn đọc chưa được thêm vào `allowed_folders` trong `config.toml`.

**Cách sửa:**
1. Mở `config.toml`
2. Thêm tên thư mục vào mục `allowed_folders`
3. Lưu file
4. Khởi động lại Claude Code CLI (đóng rồi mở lại)

---

### Lỗi: "Chế độ chỉ đọc đang bật" khi soạn email

**Triệu chứng:** Claude không thể soạn email nháp dù bạn đã yêu cầu.

**Nguyên nhân:** `read_only_mode = true` trong `config.toml` (mặc định an toàn).

**Cách sửa:**
1. Mở `config.toml`
2. Trong phần `[security]`, đổi `read_only_mode = true` thành `read_only_mode = false`
3. Lưu file và khởi động lại Claude

---

### Lỗi: "Không tìm thấy Python 3.11+"

**Triệu chứng:** Script `setup.ps1` dừng ở bước đầu tiên.

**Cách sửa:**
1. Tải Python 3.11 hoặc mới hơn tại [python.org/downloads](https://python.org/downloads)
2. Khi cài, **tích vào ô "Add Python to PATH"** — bước này rất quan trọng
3. Đóng PowerShell, mở lại, chạy lại `setup.ps1`

---

### Lỗi: "pip install thất bại" khi chạy setup.ps1

**Triệu chứng:** Bước 4 của `setup.ps1` báo lỗi cài thư viện.

**Cách sửa phổ biến:**
- Kiểm tra kết nối internet
- Thử chạy lại `setup.ps1` — đôi khi do mạng không ổn định
- Nếu vẫn lỗi, chạy thủ công:
  ```powershell
  .\venv\Scripts\pip.exe install -r requirements.txt
  ```
  Và đọc thông báo lỗi cụ thể

---

### Lỗi: Server không kết nối được với Claude

**Triệu chứng:** Các lệnh email không hoạt động dù Outlook đang mở.

**Cách kiểm tra:**
1. Xác nhận đã chạy `claude mcp add outlook` thành công
2. Khởi động lại Claude Code CLI
3. Chạy lệnh: `Liệt kê thư mục email Outlook` — nếu thấy danh sách thư mục là đã kết nối

---

### Quá nhiều yêu cầu — bị chặn tạm thời

**Triệu chứng:** Claude báo "Quá nhiều yêu cầu".

**Nguyên nhân:** Server giới hạn 60 lần gọi mỗi phút để bảo vệ hệ thống.

**Cách sửa:** Đợi vài giây rồi thử lại.

---

## 9. Giới Hạn — Những Điều Cần Biết

### Chế độ chỉ đọc mặc định

Khi mới cài đặt, server mặc định ở **chế độ chỉ đọc** (read-only mode). Nghĩa là:

- Chỉ có thể đọc và tìm kiếm email
- **Không thể** soạn hoặc trả lời email qua Claude
- Để bật tính năng soạn thảo, cần chỉnh `config.toml` (xem mục 7)

### Claude không tự gửi email

Đây là **giới hạn cố ý** vì lý do an toàn. Ngay cả khi bạn bật chế độ soạn thảo:

- Claude **chỉ mở cửa sổ soạn email** trong Outlook với nội dung đã điền
- Bạn **phải tự đọc lại và nhấn nút Gửi** trong Outlook
- Không có cách nào để Claude gửi email mà không có hành động xác nhận của bạn

### Giới hạn số lượng email mỗi lần đọc

Mặc định mỗi lần đọc tối đa 50 email (có thể tăng lên đến 200 trong `config.toml`). Nếu cần nhiều hơn, hãy dùng tính năng tìm kiếm để lọc bớt trước.

### Giới hạn độ dài nội dung email

Nội dung email được cắt bớt ở 10.000 ký tự (khoảng 5-6 trang A4). Email dài hơn sẽ bị cắt nhưng vẫn có thể đọc phần chính.

### Chỉ hỗ trợ Outlook Desktop trên Windows

Công cụ này **không hỗ trợ**:
- Outlook Web (phiên bản trên trình duyệt)
- Outlook trên Mac
- Gmail, Yahoo Mail hoặc các dịch vụ email khác
- Outlook đang chạy trong máy ảo (virtual machine) trên cùng máy

### Server phải khởi động lại sau khi chỉnh config.toml

Sau mỗi lần thay đổi `config.toml`, cần đóng Claude Code CLI và mở lại để server đọc cấu hình mới.
