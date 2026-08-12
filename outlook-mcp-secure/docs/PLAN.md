# PLAN.md — Kế Hoạch Triển Khai Claude-Outlook MCP Secure

> **Phiên bản:** 1.0  
> **Ngày:** 2026-06-24  
> **Lead Architect:** Tổng hợp từ Security Analyst + Solutions Architect + Red Team  
> **Mục tiêu:** Tích hợp Claude AI với Outlook Desktop qua MCP server Python cục bộ, bảo mật cấp ngân hàng.

---

## 1. Tổng Quan Kiến Trúc

Hệ thống này cho phép Claude AI trợ lý đọc, tìm kiếm và soạn email trong Outlook Desktop của người dùng thông qua giao thức MCP (Model Context Protocol — giao thức kết nối AI với các công cụ bên ngoài), mà **không** cần kết nối trực tiếp vào mail server IMAP/SMTP.

Luồng tổng thể:

```
Claude Code CLI  -->  server.py (MCP stdio)  -->  outlook_com.py (COM Bridge)  -->  Outlook.exe
```

Mọi thao tác gửi email VẪN đi qua Outlook Desktop — Claude chỉ mở cửa sổ soạn thảo, người dùng tự xác nhận và nhấn Gửi.

---

## 2. Cấu Trúc Thư Mục

```
outlook-mcp-secure/
├── server.py                  # Điểm vào MCP server
├── config.py                  # Đọc và validate config.toml
├── outlook_com.py             # COM Bridge — giao tiếp với Outlook
├── config.toml                # Cấu hình (người dùng chỉnh sửa)
├── config.toml.example        # Mẫu cấu hình (commit vào git)
├── requirements.txt           # Dependencies với version và hash cố định
├── requirements.lock          # Lock file với SHA-256 hash (pip-compile)
├── setup.ps1                  # Script cài đặt và khởi tạo
├── claude-mcp.json            # Cấu hình MCP cho Claude Code CLI
├── security/
│   ├── __init__.py            # Re-export các lớp bảo mật
│   ├── credential.py          # Windows Credential Manager wrapper
│   ├── audit.py               # Audit logger JSON Lines
│   └── validator.py           # Input validation và sanitization
├── tools/
│   ├── __init__.py            # Re-export tool handlers
│   ├── list_folders.py        # Tool: list_folders
│   ├── read_email.py          # Tool: list_emails + read_email
│   ├── search.py              # Tool: search_emails
│   └── compose.py             # Tool: compose_draft + reply_draft
└── docs/
    ├── PLAN.md                # File này
    ├── ARCHITECTURE.md        # Kiến trúc chi tiết
    ├── SECURITY.md            # Chính sách bảo mật
    ├── COMPONENTS.md          # Mô tả từng thành phần
    └── USER_GUIDE.md          # Hướng dẫn người dùng
```

---

## 3. Các Giai Đoạn Triển Khai

### Giai Đoạn 1 — Hạ Tầng Nền Tảng (Ngày 1–2)

**Mục tiêu:** Dựng khung cơ bản, không có logic nghiệp vụ.

Công việc cần làm:

1. Tạo Python venv portable tại `.venv/` bên trong thư mục dự án
2. Viết `requirements.txt` với version cố định (xem mục 7)
3. Viết `config.toml.example` với cấu trúc đầy đủ (xem mục 6)
4. Viết `config.py` — đọc TOML, validate bằng dataclass, singleton
5. Viết `setup.ps1` — tự động hóa bước cài đặt ban đầu
6. Viết `claude-mcp.json` — khai báo MCP server cho Claude Code CLI

**Tiêu chí hoàn thành:**
- `python server.py` khởi động không lỗi
- Config load thành công, log ra `server_start` entry

---

### Giai Đoạn 2 — Lớp Bảo Mật (Ngày 3–4)

**Mục tiêu:** Hoàn thiện toàn bộ security layer trước khi viết logic nghiệp vụ.

Công việc cần làm:

1. **`security/audit.py`** — AuditLogger ghi JSON Lines
   - Log rotation theo ngày
   - File ACL chỉ Owner + SYSTEM
   - Không log email content (kiểm tra kỹ)
   
2. **`security/credential.py`** — CredentialManager
   - Force WinVaultKeyring backend, raise lỗi nếu không phải Windows Credential Manager
   - Hàm `get_api_key()` — không cache, gọi keyring mỗi lần
   - Audit log mỗi lần truy cập credential
   
3. **`security/validator.py`** — InputValidator
   - `validate_folder_name()` — allowlist exact match
   - `validate_email_id()` — hex string, max 256 ký tự
   - `validate_search_query()` — max 200 ký tự, strip injection patterns
   - `validate_email_address()` — RFC5322 regex
   - `sanitize_string()` — loại bỏ null bytes và control characters

**Tiêu chí hoàn thành:**
- Unit test cho tất cả validator functions
- Audit log file được tạo với đúng ACL
- Cố tình gọi `get_api_key()` khi chưa setup — phải raise `CredentialNotFoundError` rõ ràng

---

### Giai Đoạn 3 — COM Bridge (Ngày 5–7)

**Mục tiêu:** Lớp giao tiếp Outlook COM an toàn, đúng threading model.

Công việc cần làm:

1. **`outlook_com.py`** — OutlookCOMBridge
   - STA thread setup với `pythoncom.CoInitialize()`
   - Context manager `__enter__`/`__exit__` với `ReleaseComObject`
   - Task dispatch qua `concurrent.futures.ThreadPoolExecutor(max_workers=1)`
   - Retry logic cho `RPC_E_CALL_REJECTED` (tối đa 3 lần, backoff 0.5s/1s/2s)
   - Kiểm tra Outlook đang chạy trước khi connect (dùng `GetActiveObject`, không `Dispatch`)
   - Sanitize lỗi COM trước khi trả ra ngoài

**Các method của OutlookCOMBridge:**
- `get_namespace()` — lấy MAPI namespace
- `get_folder_by_allowlist_name(name)` — resolve folder theo allowlist
- `list_mail_items(folder, limit, offset)` — liệt kê MailItems
- `get_mail_by_entry_id(entry_id)` — đọc email theo EntryID
- `create_draft(to, cc, subject, body, importance)` — tạo draft + Display()
- `create_reply_draft(entry_id, body, reply_all, additional_cc)` — reply + Display()

**Tiêu chí hoàn thành:**
- Chạy được `list_folders` từ Python shell, không crash Outlook
- Sau 1000 lần gọi liên tiếp, Outlook process memory không tăng (không leak COM objects)
- Khi Outlook đóng, server trả về lỗi rõ ràng thay vì crash

---

### Giai Đoạn 4 — MCP Tool Handlers (Ngày 8–10)

**Mục tiêu:** Cài đặt 6 tools chức năng hoàn chỉnh.

Thứ tự cài đặt (từ đơn giản đến phức tạp):

1. `list_folders` — đơn giản nhất, không cần params
2. `list_emails` — cần folder validation
3. `read_email` — cần entry_id validation
4. `search_emails` — cần DASL filter construction
5. `compose_draft` — cần read_only check, email validation
6. `reply_draft` — cần entry_id + read_only check

**Tiêu chí hoàn thành:**
- Mỗi tool có unit test riêng với mock COM objects
- `grep -r "\.Send()" tools/` phải trả về 0 kết quả
- `grep -r "imaplib\|smtplib" .` phải trả về 0 kết quả

---

### Giai Đoạn 5 — MCP Server Integration (Ngày 11–12)

**Mục tiêu:** Kết nối tất cả thành phần, chạy thử với Claude Code CLI.

Công việc cần làm:

1. Viết `server.py` — MCP Server với stdio transport
2. Đăng ký 6 tools với JSON Schema đầy đủ
3. Implement `handle_call_tool()` với:
   - Validation layer (gọi InputValidator)
   - Audit logging (trước và sau khi execute)
   - COM thread dispatch (dùng `run_in_executor`)
   - Error sanitization
4. Thêm rate limiting (60 calls/phút/session)
5. Thêm timeout (30 giây/call)

**Tiêu chí hoàn thành:**
- Claude Code CLI gọi được `list_folders` và nhận kết quả thực tế
- Audit log ghi đúng format, không có email content
- Khi gọi `compose_draft`, cửa sổ Outlook mở ra, KHÔNG tự gửi

---

### Giai Đoạn 6 — Kiểm Thử Bảo Mật (Ngày 13–14)

**Mục tiêu:** Verify tất cả security controls hoạt động như thiết kế.

Checklist kiểm thử:

| Kịch bản tấn công | Kết quả kỳ vọng |
|---|---|
| Gọi tool với `folder_path = "../Contacts"` | Bị từ chối, log `block_reason=not_in_allowlist` |
| Gọi `compose_draft` khi `read_only=true` | Bị từ chối, log `block_reason=read_only_mode` |
| Email chứa `[IGNORE PREVIOUS INSTRUCTIONS]` trong subject | Tool trả về dữ liệu thô, Claude không bị inject |
| Gọi `search_emails` với query dài 1000 ký tự | Bị từ chối, log `block_reason=invalid_input` |
| Gọi 100 tools trong 1 phút | Bị throttle sau call thứ 60 |
| Xóa/sửa audit log file trong khi server chạy | Server phát hiện ACL thay đổi, refuse to write |
| `entry_id` chứa ký tự không phải hex | Bị từ chối, log validation error |
| Thử truy cập PST path trực tiếp qua COM | Không có code path nào cho phép điều này |

---

## 4. MCP Tools — Đặc Tả Chính Thức

### 4.1 list_folders

**Mô tả:** Liệt kê các thư mục email mà Claude được phép truy cập.

**Tham số đầu vào:**

| Tên | Kiểu | Bắt buộc | Giá trị mặc định | Mô tả |
|---|---|---|---|---|
| `include_subfolders` | boolean | Không | `false` | Có liệt kê thư mục con không |

**Kết quả trả về:**
```json
{
  "folders": [
    {
      "name": "Inbox",
      "path": "Inbox",
      "unread_count": 5,
      "total_count": 120
    }
  ]
}
```

**Ràng buộc bảo mật:**
- Filter theo `allowed_folders` trong config trước khi trả về
- Không bao giờ expose đường dẫn PST file
- Audit log: ghi `folder_count`, không ghi tên folder cụ thể nếu không có trong allowlist

---

### 4.2 list_emails

**Mô tả:** Lấy danh sách email trong một thư mục (chỉ metadata, không nội dung).

**Tham số đầu vào:**

| Tên | Kiểu | Bắt buộc | Giá trị mặc định | Giới hạn | Mô tả |
|---|---|---|---|---|---|
| `folder_path` | string | Có | — | max 260 ký tự | Tên thư mục (phải có trong allowlist) |
| `limit` | integer | Không | 20 | tối đa 100 | Số email tối đa mỗi lần |
| `offset` | integer | Không | 0 | — | Phân trang |
| `unread_only` | boolean | Không | `false` | — | Chỉ lấy email chưa đọc |

**Kết quả trả về:**
```json
{
  "emails": [
    {
      "entry_id": "00000000ABC123...",
      "subject": "Báo cáo tháng 6",
      "sender_name": "Nguyễn Văn A",
      "sender_email": "a@example.com",
      "received_time": "2026-06-24T09:30:00+07:00",
      "has_attachment": true,
      "is_read": false,
      "size_kb": 45
    }
  ],
  "total": 120
}
```

**Ràng buộc bảo mật:**
- `validate_folder_name()` kiểm tra `folder_path` trước khi COM lookup
- `limit` bị cap tại `config.security.max_results`
- Audit log: ghi `folder_name`, `items_returned`, không ghi subject/sender

---

### 4.3 read_email

**Mô tả:** Đọc nội dung đầy đủ của một email.

**Tham số đầu vào:**

| Tên | Kiểu | Bắt buộc | Giới hạn | Mô tả |
|---|---|---|---|---|
| `entry_id` | string | Có | hex string, max 256 ký tự | Outlook Entry ID từ `list_emails` |

**Kết quả trả về:**
```json
{
  "subject": "Báo cáo tháng 6",
  "sender_name": "Nguyễn Văn A",
  "sender_email": "a@example.com",
  "to_recipients": ["b@example.com"],
  "cc_recipients": [],
  "received_time": "2026-06-24T09:30:00+07:00",
  "body_text": "Kính gửi...",
  "attachments": [
    {
      "name": "report.pdf",
      "size_kb": 512,
      "extension": ".pdf"
    }
  ]
}
```

**Ràng buộc bảo mật:**
- `validate_email_id()` — reject nếu không phải hex string hợp lệ
- Folder chứa email phải thuộc allowlist
- Body HTML được strip thành plain text trước khi trả về
- Email content KHÔNG ghi vào audit log (chỉ ghi `entry_id` truncated 8 ký tự)

---

### 4.4 search_emails

**Mô tả:** Tìm kiếm email trong các thư mục được phép.

**Tham số đầu vào:**

| Tên | Kiểu | Bắt buộc | Giới hạn | Mô tả |
|---|---|---|---|---|
| `query` | string | Có | max 200 ký tự | Từ khóa tìm kiếm |
| `folder_path` | string | Không | max 260 ký tự | Giới hạn trong thư mục này |
| `search_in` | enum | Không | — | `subject`, `body`, `sender`, `all` (mặc định: `subject`) |
| `date_from` | string (date) | Không | — | Từ ngày (YYYY-MM-DD) |
| `date_to` | string (date) | Không | — | Đến ngày (YYYY-MM-DD) |
| `limit` | integer | Không | tối đa 50 | Số kết quả tối đa |

**Kết quả trả về:**
```json
{
  "results": [
    {
      "entry_id": "00000000DEF456...",
      "subject": "Báo cáo Q2",
      "sender_email": "a@example.com",
      "received_time": "2026-06-20T14:00:00+07:00",
      "folder_path": "Inbox",
      "snippet": "...nội dung liên quan đến từ khóa..."
    }
  ],
  "total_found": 3
}
```

**Ràng buộc bảo mật:**
- `validate_search_query()` — strip SQL/DASL injection, ký tự điều khiển
- Chỉ tìm trong allowlist folders
- Audit log: ghi `SHA256(query)` không ghi query plaintext
- Dùng `Items.Restrict()` với DASL filter (không vòng lặp Python thủ công)

---

### 4.5 compose_draft

**Mô tả:** Tạo email nháp và mở cửa sổ Outlook để người dùng xem lại.

**Tham số đầu vào:**

| Tên | Kiểu | Bắt buộc | Giới hạn | Mô tả |
|---|---|---|---|---|
| `to` | array[string] | Có | tối đa 50 địa chỉ | Danh sách người nhận |
| `cc` | array[string] | Không | tối đa 50 địa chỉ | Danh sách CC |
| `subject` | string | Có | max 500 ký tự | Tiêu đề email |
| `body` | string | Có | max 50.000 ký tự | Nội dung email |
| `importance` | enum | Không | — | `low`, `normal`, `high` (mặc định: `normal`) |

**Kết quả trả về:**
```json
{
  "status": "draft_opened",
  "message": "Cửa sổ soạn email đã mở trong Outlook. Vui lòng xem lại và nhấn Send để gửi.",
  "draft_entry_id": "00000000GHI789..."
}
```

**Ràng buộc bảo mật:**
- Yêu cầu `read_only_mode = false` trong config
- Mỗi địa chỉ email được `validate_email_address()` 
- Chỉ gọi `.Display()` — TUYỆT ĐỐI KHÔNG gọi `.Send()`
- Audit log: ghi `SHA256(to_addresses)`, `SHA256(subject)` — không ghi body

---

### 4.6 reply_draft

**Mô tả:** Tạo bản trả lời cho email và mở cửa sổ Outlook.

**Tham số đầu vào:**

| Tên | Kiểu | Bắt buộc | Giới hạn | Mô tả |
|---|---|---|---|---|
| `entry_id` | string | Có | hex, max 256 ký tự | Entry ID email cần trả lời |
| `body` | string | Có | max 50.000 ký tự | Nội dung phản hồi |
| `reply_all` | boolean | Không | — | `true` = Reply All (mặc định: `false`) |
| `additional_cc` | array[string] | Không | tối đa 20 địa chỉ | CC thêm |

**Kết quả trả về:**
```json
{
  "status": "reply_opened",
  "message": "Cửa sổ trả lời đã mở trong Outlook. Vui lòng xem lại và nhấn Send để gửi.",
  "reply_entry_id": "00000000JKL012..."
}
```

**Ràng buộc bảo mật:**
- Yêu cầu `read_only_mode = false` trong config
- `validate_email_id()` cho `entry_id`
- Email gốc phải thuộc allowlist folders
- Gọi `.Reply()` hoặc `.ReplyAll()` rồi `.Display()` — KHÔNG `.Send()`
- Audit log: ghi `SHA256(entry_id)`, `action_type`

---

## 5. Code Patterns Bắt Buộc

### 5.1 COM STA Threading — Mô Hình Bắt Buộc

```python
# outlook_com.py — Toàn bộ COM operation trong 1 thread duy nhất
import pythoncom
import win32com.client
import gc
import threading
from concurrent.futures import Future
from queue import Queue

# Hàng đợi task cho STA thread
_com_task_queue: Queue = Queue()

def _sta_worker():
    """
    Vòng lặp chính của STA thread — xử lý tất cả COM operations.
    Bắt buộc gọi CoInitialize trước bất kỳ win32com call nào.
    """
    # Bước 1: Khởi tạo COM STA apartment cho thread này
    pythoncom.CoInitialize()
    try:
        while True:
            # Bước 2: Lấy task từ hàng đợi (blocking)
            task = _com_task_queue.get()
            # Bước 3: Sentinel value None = thoát vòng lặp
            if task is None:
                break
            func, args, kwargs, future = task
            try:
                # Bước 4: Thực thi COM operation và set kết quả
                result = func(*args, **kwargs)
                future.set_result(result)
            except Exception as e:
                # Bước 5: Set exception để caller nhận được
                future.set_exception(e)
    finally:
        # Bước 6: Dọn dẹp COM apartment khi thread kết thúc
        pythoncom.CoUninitialize()

# Khởi động STA thread khi module load
_sta_thread = threading.Thread(
    target=_sta_worker,
    name='OutlookCOMThread',
    daemon=True
)
_sta_thread.start()

def dispatch_to_sta(func, *args, **kwargs) -> Future:
    """
    Gửi function vào STA thread để thực thi an toàn với COM.
    Trả về Future — caller await kết quả.
    """
    future = Future()
    _com_task_queue.put((func, args, kwargs, future))
    return future
```

### 5.2 Context Manager Giải Phóng COM Objects

```python
class OutlookCOMBridge:
    """
    Context manager bọc toàn bộ Outlook COM lifecycle.
    Tự động giải phóng COM objects khi exit.
    """

    def __init__(self):
        # Danh sách COM objects cần release khi kết thúc
        self._refs = []

    def __enter__(self):
        # Kết nối đến Outlook đang chạy — KHÔNG tạo instance mới
        try:
            self._app = win32com.client.GetActiveObject('Outlook.Application')
            self._refs.append(self._app)
            return self
        except Exception:
            raise OutlookNotRunningError(
                'Outlook không đang chạy. Vui lòng mở Outlook trước.'
            )

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Giải phóng theo thứ tự ngược (item -> folder -> namespace -> app)
        for ref in reversed(self._refs):
            try:
                win32com.client.ReleaseComObject(ref)
            except Exception:
                pass
        self._refs.clear()
        # Buộc Python garbage collector thu hồi COM references còn sót
        gc.collect()
        # Không suppress exception — trả về False
        return False
```

### 5.3 MCP Server Dispatch Pattern

```python
# server.py — Xử lý tool call từ Claude

import asyncio
from concurrent.futures import ThreadPoolExecutor

# Executor với max_workers=1 đảm bảo chỉ 1 COM thread duy nhất
_com_executor = ThreadPoolExecutor(
    max_workers=1,
    thread_name_prefix='outlook-com'
)

async def run_in_com_thread(func, *args, **kwargs):
    """
    Chạy COM operation trong COM thread, await từ asyncio event loop.
    Bắt buộc dùng hàm này — KHÔNG gọi win32com trực tiếp từ coroutine.
    """
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(
        _com_executor,
        lambda: func(*args, **kwargs)
    )

@app.call_tool()
async def handle_call_tool(name: str, arguments: dict):
    """
    Điểm vào duy nhất cho tất cả tool calls từ Claude.
    Thứ tự bắt buộc: validate -> audit -> dispatch -> audit -> return.
    """
    start_time = time.monotonic()
    try:
        # Bước 1: Validate input trước tất cả
        validated_args = await validate_tool_args(name, arguments)
        # Bước 2: Log bắt đầu call
        audit.log_start(name, validated_args)
        # Bước 3: Dispatch sang COM thread
        result = await run_in_com_thread(
            TOOL_DISPATCH[name], validated_args
        )
        # Bước 4: Log kết quả thành công
        duration_ms = int((time.monotonic() - start_time) * 1000)
        audit.log_success(name, result.get('_meta', {}), duration_ms)
        return [TextContent(type='text', text=json.dumps(result))]
    except ValidationError as e:
        audit.log_blocked(name, str(e))
        return [TextContent(type='text', text=json.dumps({'error': str(e)}))]
    except OutlookOperationError as e:
        audit.log_error(name, 'outlook_error')
        return [TextContent(type='text', text=json.dumps({'error': str(e)}))]
```

### 5.4 DASL Filter cho Search

```python
# tools/search.py — Dùng DASL thay vì vòng lặp Python thủ công

def build_dasl_filter(query: str, search_in: str,
                      date_from: str | None, date_to: str | None) -> str:
    """
    Tạo DASL filter string cho Outlook Items.Restrict().
    DASL (DAV Searching and Locating) — ngôn ngữ tìm kiếm của Outlook MAPI.
    Query đã được validate và sanitize trước khi vào hàm này.
    """
    # Escape dấu nháy đơn để tránh injection vào DASL query
    safe_query = query.replace("'", "''")

    # Bước 1: Xây dựng điều kiện tìm kiếm theo trường
    if search_in == 'subject':
        content_filter = (
            f'"urn:schemas:httpmail:subject" LIKE \'%{safe_query}%\''
        )
    elif search_in == 'sender':
        content_filter = (
            f'"urn:schemas:httpmail:fromemail" LIKE \'%{safe_query}%\''
        )
    elif search_in == 'body':
        content_filter = (
            f'"urn:schemas:httpmail:textdescription" LIKE \'%{safe_query}%\''
        )
    else:  # 'all'
        content_filter = (
            f'"urn:schemas:httpmail:subject" LIKE \'%{safe_query}%\' OR '
            f'"urn:schemas:httpmail:textdescription" LIKE \'%{safe_query}%\''
        )

    # Bước 2: Thêm bộ lọc ngày nếu có
    date_filters = []
    if date_from:
        date_filters.append(
            f'"urn:schemas:httpmail:datereceived" >= \'{date_from}T00:00:00Z\''
        )
    if date_to:
        date_filters.append(
            f'"urn:schemas:httpmail:datereceived" <= \'{date_to}T23:59:59Z\''
        )

    # Bước 3: Ghép thành DASL hoàn chỉnh
    all_conditions = [f'({content_filter})'] + date_filters
    return '@SQL=' + ' AND '.join(all_conditions)
```

### 5.5 Lỗi COM Không Expose Ra Ngoài

```python
# outlook_com.py — Sanitize lỗi trước khi trả về tool handler

import pywintypes

def _safe_com_call(operation_name: str, func, *args):
    """
    Wrapper bắt lỗi COM và chuyển thành thông báo an toàn.
    Không expose COM error code, HRESULT, hay stack trace ra ngoài.
    """
    try:
        return func(*args)
    except pywintypes.error as e:
        # Log chi tiết nội bộ để debug — KHÔNG trả về caller
        _internal_logger.debug(
            'COM error in %s: hresult=0x%08X, strerror=%s',
            operation_name, e.winerror, e.strerror
        )
        # Thông báo an toàn cho người dùng
        raise OutlookOperationError(
            f'Không thể thực hiện "{operation_name}". '
            'Đảm bảo Outlook đang chạy và thử lại.'
        )
    except pythoncom.error as e:
        _internal_logger.debug('pythoncom error in %s: %s', operation_name, e)
        raise OutlookOperationError(
            'Mất kết nối với Outlook. Khởi động lại Outlook và thử lại.'
        )
```

---

## 6. config.toml — Schema Đầy Đủ

```toml
# config.toml — Cấu hình Claude-Outlook MCP Secure
# KHÔNG chứa credentials — API key lưu trong Windows Credential Manager

[outlook]
# Tên tài khoản email Outlook
account_name = "thanhnt@softmart.net.vn"
# Tên hiển thị của PST file (bỏ trống nếu dùng mailbox mặc định)
pst_display_name = ""

[security]
# Chế độ chỉ đọc (true = không cho compose/reply, KHUYẾN NGHỊ GIỮ true)
read_only_mode = true

# Danh sách thư mục Claude được phép truy cập
# Dùng tên tiếng Anh chuẩn — server tự resolve sang ngôn ngữ OS
# Hỗ trợ path lồng nhau: "Inbox/Projects"
# Hỗ trợ wildcard 1 cấp: "Inbox/*" (không hỗ trợ "Inbox/**")
allowed_folders = [
    "Inbox",
    "Sent Items",
    "Drafts",
]

# Số kết quả tối đa cho search và list (tối đa tuyệt đối: 200)
max_results = 50

# Số recipients tối đa mỗi email draft
max_recipients_per_draft = 20

# Độ dài tối đa của Entry ID (hex string)
entry_id_max_length = 256

[audit]
# Thư mục chứa audit log (tương đối so với thư mục dự án)
log_dir = "logs"

# Số ngày giữ log (banking compliance: 90 ngày)
retain_days = 90

# Thuật toán hash cho params trong log
hash_algorithm = "sha256"

[limits]
# Giới hạn độ dài các input
search_query_max_length = 200
email_body_max_length = 50000
subject_max_length = 500
list_emails_default_limit = 20
list_emails_max_limit = 100

# Rate limiting
max_calls_per_minute = 60

# Timeout cho mỗi COM operation (giây)
com_operation_timeout_seconds = 30
```

---

## 7. requirements.txt — Dependencies Cố Định

```
# Claude-Outlook MCP Secure — Dependencies
# Tất cả version phải PIN CHÍNH XÁC (==) để tránh supply chain attack

# MCP SDK — giao thức kết nối Claude với tools
mcp==1.0.0

# Outlook COM automation trên Windows
pywin32==306

# Windows Credential Manager integration
keyring==25.0.0

# TOML config parsing (Python 3.11+ có sẵn tomllib trong stdlib)
# tomli là backport cho Python < 3.11
tomli==2.0.1

# HTML strip để làm sạch email body trước khi trả về Claude
beautifulsoup4==4.12.3

# [Dev dependencies — không cần trong production]
pytest==8.2.0
pytest-asyncio==0.23.6
```

Sau khi cài đặt, chạy `pip-compile --generate-hashes requirements.txt -o requirements.lock` để tạo lock file với SHA-256 hash.

---

## 8. Audit Log — Format JSON Lines

**Vị trí file:** `%APPDATA%\OutlookMCPSecure\audit\audit-YYYY-MM-DD.jsonl`

**Ví dụ các loại log entry:**

```jsonl
{"ts":"2026-06-24T09:00:00.123456+07:00","event":"server_start","session_id":"a1b2c3d4-...","version":"1.0.0","read_only":true,"allowlist_count":3,"com_backend":"win32com"}
{"ts":"2026-06-24T09:01:15.456789+07:00","session_id":"a1b2c3d4-...","tool":"list_folders","params":{"include_subfolders":false},"status":"ok","duration_ms":145,"items_returned":3}
{"ts":"2026-06-24T09:02:30.789012+07:00","session_id":"a1b2c3d4-...","tool":"list_emails","params":{"folder":"Inbox","limit":20},"status":"ok","duration_ms":234,"items_returned":20}
{"ts":"2026-06-24T09:03:45.012345+07:00","session_id":"a1b2c3d4-...","tool":"read_email","params":{"entry_id_prefix":"00000000"},"status":"ok","duration_ms":89,"items_returned":1}
{"ts":"2026-06-24T09:04:00.111222+07:00","session_id":"a1b2c3d4-...","tool":"search_emails","params":{"query_hash":"sha256:abc123...","folder":"Inbox","search_in":"subject"},"status":"ok","duration_ms":312,"items_returned":5}
{"ts":"2026-06-24T09:05:10.333444+07:00","session_id":"a1b2c3d4-...","tool":"list_emails","params":{"folder":"../Contacts","limit":20},"status":"blocked","block_reason":"not_in_allowlist","risk_level":"medium","duration_ms":2}
{"ts":"2026-06-24T09:06:20.555666+07:00","session_id":"a1b2c3d4-...","tool":"compose_draft","params":{},"status":"blocked","block_reason":"read_only_mode","risk_level":"low","duration_ms":1}
{"ts":"2026-06-24T09:07:30.777888+07:00","event":"credential_access","session_id":"a1b2c3d4-...","credential_type":"anthropic_api_key","status":"ok"}
{"ts":"2026-06-24T09:08:00.999000+07:00","event":"integrity_check","session_id":"a1b2c3d4-...","entries_since_last":100,"sha256_last_10_entries":"def456..."}
{"ts":"2026-06-24T17:00:00.000000+07:00","event":"server_stop","session_id":"a1b2c3d4-...","total_calls":145,"errors":2}
```

**Quy tắc bắt buộc:**
- Mỗi dòng là một JSON object độc lập, kết thúc bằng `\n`
- Encoding UTF-8, không BOM
- KHÔNG ghi: subject, body, sender, recipient, attachment name, API key, stack trace
- Mở file với `mode='a'` (append-only)
- Mỗi 100 entries ghi 1 checksum entry để phát hiện giả mạo

---

## 9. Folder Allowlist — Logic Validation

### Chuỗi Xử Lý

```
Tên folder từ tool params
    -> strip whitespace
    -> Unicode NFC normalize
    -> casefold() (lowercase unicode-aware)
    -> So sánh với allowed_folders trong config
    -> Nếu pass: resolve qua COM để lấy folder object
    -> Verify COM object .Name property == tên kỳ vọng (chống giả mạo)
    -> Nếu fail: log blocked entry, raise FolderNotAllowedError
```

### Xử Lý Trường Hợp Đặc Biệt

| Trường hợp | Giải pháp |
|---|---|
| Folder tên tiếng Việt (Unicode) | Dùng `casefold()` + NFC normalize trước khi so sánh |
| Tên mặc định theo ngôn ngữ OS (Inbox → "Hộp thư đến") | Dùng `GetDefaultFolder(olFolderInbox)` thay vì tìm theo tên |
| Trùng tên ở các level khác nhau | Allowlist dùng path notation: `"Inbox/Projects"` |
| Wildcard một cấp | Hỗ trợ `"Inbox/*"` — chỉ một level, không recursive |
| Null bytes trong tên folder | `validate_folder_name()` reject ngay lập tức |
| Path absolute (C:\\...) | Reject nếu chứa `://`, `:\\`, hay `../` |

### Cấu Hình Mẫu

```toml
[security]
allowed_folders = [
    "Inbox",           # Hộp thư đến chính
    "Inbox/*",         # Tất cả subfolder trực tiếp của Inbox
    "Inbox/Projects",  # Subfolder cụ thể
    "Sent Items",      # Thư đã gửi
    "Drafts",          # Thư nháp
]
```

---

## 10. Security Controls — Checklist Không Được Bỏ

### Controls Bắt Buộc (Không Đàm Phán)

- [ ] **TUYỆT ĐỐI KHÔNG** gọi `.Send()` trong bất kỳ file nào — kiểm tra bằng `grep -r "\.Send()" tools/`
- [ ] **TUYỆT ĐỐI KHÔNG** import `imaplib`, `smtplib` trong bất kỳ file nào
- [ ] **TUYỆT ĐỐI KHÔNG** ghi email content vào audit log
- [ ] **TUYỆT ĐỐI KHÔNG** lưu API key vào file hay environment variable
- [ ] **TUYỆT ĐỐI KHÔNG** gọi win32com từ asyncio coroutine hay ThreadPoolExecutor thread thông thường
- [ ] **TUYỆT ĐỐI KHÔNG** expose COM object ra ngoài OutlookCOMBridge
- [ ] **TUYỆT ĐỐI KHÔNG** bind server trên interface nào khác 127.0.0.1 (với stdio transport: không bind)
- [ ] **TUYỆT ĐỐI KHÔNG** dùng `eval()`, `exec()`, hay `pickle` với data từ email
- [ ] **TUYỆT ĐỐI KHÔNG** cho phép `allowed_folders` rỗng khi `read_only=false`
- [ ] **TUYỆT ĐỐI KHÔNG** dùng `Dispatch('Outlook.Application')` — phải dùng `GetActiveObject()`

### Controls Nên Có (Khuyến Nghị Cao)

- [ ] Rate limiting: 60 calls/phút/session
- [ ] Timeout: 30 giây/COM operation
- [ ] Integrity check log mỗi 100 entries
- [ ] ACL verification khi server khởi động
- [ ] Hash verification cho requirements.lock trước khi install
- [ ] Process parent check khi startup (verify caller là Claude Code CLI)

---

## 11. Hướng Dẫn Cài Đặt Nhanh

### Bước 1 — Chuẩn Bị

```powershell
# Chạy từ thư mục outlook-mcp-secure/
.\setup.ps1
```

Script sẽ tự động:
1. Tạo Python venv tại `.venv/`
2. Cài đặt dependencies từ `requirements.txt`
3. Tạo `config.toml` từ mẫu
4. Tạo thư mục `logs/` với ACL đúng
5. Lưu Anthropic API key vào Windows Credential Manager (yêu cầu nhập tay)

### Bước 2 — Cấu Hình

Mở `config.toml` và chỉnh sửa:
- `account_name` — địa chỉ email Outlook của bạn
- `allowed_folders` — danh sách thư mục Claude được phép đọc
- `read_only_mode` — giữ `true` cho đến khi tin tưởng hệ thống

### Bước 3 — Kết Nối Claude Code CLI

Copy hoặc merge nội dung `claude-mcp.json` vào file `.claude/mcp.json` trong thư mục làm việc của Claude Code.

### Bước 4 — Kiểm Tra

```powershell
# Kiểm tra server khởi động bình thường
.venv\Scripts\python.exe server.py

# Trong Claude Code, thử:
# "Liệt kê các thư mục email"
# "Đọc 5 email mới nhất trong Inbox"
```

---

## 12. Rủi Ro Đã Biết Và Cách Giảm Thiểu

| Rủi ro | Mức độ | Biện pháp giảm thiểu |
|---|---|---|
| Prompt injection qua email content | Nghiêm trọng | Wrap email trong JSON neutral container, system prompt cứng, chỉ `.Display()` không `.Send()` |
| COM object leak gây Outlook crash | Trung bình | Context manager bắt buộc, `ReleaseComObject` + `gc.collect()` sau mỗi call |
| Audit log bị giả mạo | Trung bình | File ACL chỉ Owner+SYSTEM, checksum entry mỗi 100 dòng |
| DPAPI side-channel đọc API key | Trung bình | Service name ngẫu nhiên, key rotation định kỳ |
| Supply chain attack qua pip | Trung bình | Pin exact version + SHA-256 hash trong requirements.lock |
| COM threading violation crash | Cao | STA thread duy nhất, ThreadPoolExecutor max_workers=1 |
| Unicode lookalike folder name spoof | Thấp | casefold() + NFC normalize + verify COM .Name sau resolve |

---

## 13. Tiêu Chí Hoàn Thành Dự Án

Hệ thống được coi là sẵn sàng đưa vào sử dụng khi:

1. **Chức năng:** Tất cả 6 tools hoạt động đúng với Outlook Desktop thực tế
2. **Bảo mật:** Tất cả 10 checklist "Không Được Bỏ" đã được verify bằng grep/test
3. **Audit:** Log ghi đúng format, không có email content, ACL được set đúng
4. **Threading:** Chạy 1000 tool calls liên tiếp không crash Outlook, không leak memory
5. **Error handling:** Mọi lỗi COM đều được sanitize trước khi trả về Claude
6. **Tài liệu:** USER_GUIDE.md đủ để người không kỹ thuật cài đặt thành công

---

*Tài liệu này được tổng hợp từ phân tích của Security Analyst, Solutions Architect và Red Team. Mọi thay đổi lớn về kiến trúc cần review lại toàn bộ threat model.*
