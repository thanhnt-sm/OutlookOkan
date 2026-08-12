# COMPONENTS.md — Mô Tả Thành Phần Hệ Thống

> **Phiên bản:** 1.0
> **Ngày:** 2026-06-24
> **Mục tiêu:** Mô tả chi tiết từng file, mục đích, dependencies và interfaces của hệ thống Claude-Outlook MCP Secure.

---

## 1. Bảng Thành Phần Tổng Quan

| File / Thư Mục | Mục Đích | Phụ Thuộc Vào | Được Dùng Bởi |
|---|---|---|---|
| `server.py` | Điểm vào MCP server, điều phối toàn bộ tool calls | `config.py`, `security/`, `tools/`, `outlook_com.py` | Claude Code CLI (qua stdio) |
| `config.py` | Đọc và validate cấu hình từ `config.toml` | `tomllib` (stdlib) hoặc `tomli` | Tất cả các module |
| `outlook_com.py` | COM Bridge — giao tiếp an toàn với Outlook Desktop | `pywin32`, `pythoncom`, `pywintypes` | Tất cả `tools/*.py` |
| `security/credential.py` | Lưu và đọc API key từ Windows Credential Manager | `keyring`, `security/audit.py` | `server.py` |
| `security/audit.py` | Ghi audit log JSON Lines, rotation, checksum | stdlib (`json`, `pathlib`, `datetime`) | `server.py`, tất cả `tools/*.py` |
| `security/validator.py` | Validate và sanitize toàn bộ input từ Claude | stdlib (`re`, `unicodedata`) | `server.py`, tất cả `tools/*.py` |
| `security/__init__.py` | Re-export các lớp bảo mật | `credential.py`, `audit.py`, `validator.py` | Các module import `security` |
| `tools/list_folders.py` | Tool `list_folders` — liệt kê thư mục được phép | `outlook_com.py`, `security/audit.py` | `server.py` |
| `tools/read_email.py` | Tool `list_emails` + `read_email` — đọc email | `outlook_com.py`, `security/validator.py`, `beautifulsoup4` | `server.py` |
| `tools/search.py` | Tool `search_emails` — tìm kiếm qua DASL | `outlook_com.py`, `security/validator.py` | `server.py` |
| `tools/compose.py` | Tool `compose_draft` + `reply_draft` — soạn email | `outlook_com.py`, `security/validator.py` | `server.py` |
| `tools/__init__.py` | Re-export tool handlers | Tất cả `tools/*.py` | `server.py` |
| `config.toml` | File cấu hình người dùng chỉnh sửa | — | `config.py` |
| `config.toml.example` | Mẫu cấu hình (commit vào git) | — | Người dùng tham khảo |
| `requirements.txt` | Dependencies với version PIN chính xác | — | `pip install` |
| `requirements.lock` | Lock file với SHA-256 hash (pip-compile) | — | `pip install --require-hashes` |
| `setup.ps1` | Script cài đặt và khởi tạo môi trường | PowerShell 5.1+ | Người dùng chạy 1 lần |
| `claude-mcp.json` | Khai báo MCP server cho Claude Code CLI | — | Claude Code CLI |

---

## 2. server.py — Entry Point, Lifecycle

### Mục Đích

`server.py` là điểm vào duy nhất của hệ thống. File này:
- Khởi động MCP server với transport `stdio` (đọc/ghi qua stdin/stdout)
- Đăng ký 6 tools với Claude Code CLI kèm JSON Schema đầy đủ
- Là "bộ điều phối trung tâm" — nhận tool call từ Claude, validate, dispatch sang COM thread, trả kết quả về

### Lifecycle (Vòng Đời)

```
Khởi động
    -> Đọc config.toml (config.py)
    -> Kiểm tra Outlook đang chạy
    -> Khởi tạo AuditLogger, đặt ACL file log
    -> Ghi log entry "server_start"
    -> Đăng ký 6 tools với MCP SDK
    -> Bắt đầu lắng nghe stdin (asyncio event loop)

Nhận tool call từ Claude
    -> validate_tool_args(name, arguments)  [security/validator.py]
    -> audit.log_start(name, safe_params)   [security/audit.py]
    -> run_in_com_thread(tool_handler, args) [asyncio executor]
    -> audit.log_result(name, meta, duration_ms)
    -> Trả về JSON result qua stdout

Tắt server (Ctrl+C hoặc stdin đóng)
    -> Ghi log entry "server_stop" (tổng số calls, lỗi)
    -> Dừng STA COM thread (gửi sentinel None vào queue)
    -> Đợi thread kết thúc
    -> Thoát
```

### Các Thành Phần Nội Bộ

| Thành Phần | Kiểu | Mô Tả |
|---|---|---|
| `app` | `mcp.Server` | Instance MCP server từ SDK |
| `_com_executor` | `ThreadPoolExecutor(max_workers=1)` | Executor duy nhất cho COM thread |
| `TOOL_DISPATCH` | `dict[str, Callable]` | Map tên tool → hàm handler tương ứng |
| `_rate_limiter` | `RateLimiter` | Giới hạn 60 calls/phút/session |
| `handle_call_tool()` | `async def` | Dispatcher chính, xử lý mọi tool call |
| `run_in_com_thread()` | `async def` | Wrapper chạy hàm trong COM thread an toàn |
| `validate_tool_args()` | `async def` | Gọi InputValidator theo từng tool |

### Hàm Chính

```python
async def handle_call_tool(name: str, arguments: dict) -> list[TextContent]:
    """
    Điểm vào duy nhất cho tất cả tool calls từ Claude.
    Thứ tự bắt buộc: validate -> audit_start -> dispatch -> audit_result -> return.
    Ném ValidationError nếu input không hợp lệ.
    Ném OutlookOperationError nếu COM thất bại.
    Luôn trả về TextContent JSON — không bao giờ raise ra ngoài.
    """
```

```python
async def run_in_com_thread(func: Callable, *args, **kwargs) -> Any:
    """
    Chạy hàm COM trong executor thread duy nhất.
    Bắt buộc dùng hàm này — KHÔNG gọi win32com trực tiếp từ coroutine.
    Timeout sau config.limits.com_operation_timeout_seconds giây.
    """
```

---

## 3. config.py — Settings, Config.toml Reference Đầy Đủ

### Mục Đích

`config.py` đọc `config.toml`, validate kiểu dữ liệu, và cung cấp singleton `AppConfig` cho toàn bộ hệ thống. Không module nào đọc TOML trực tiếp — tất cả phải qua `config.py`.

### Dataclasses (Cấu Trúc Dữ Liệu)

```python
@dataclass(frozen=True)
class OutlookConfig:
    """Cấu hình kết nối Outlook"""
    account_name: str           # Địa chỉ email tài khoản Outlook
    pst_display_name: str = ""  # Tên PST file (bỏ trống = mailbox mặc định)

@dataclass(frozen=True)
class SecurityConfig:
    """Cấu hình bảo mật và quyền truy cập"""
    read_only_mode: bool              # True = không cho compose/reply
    allowed_folders: list[str]        # Danh sách thư mục được phép
    max_results: int                  # Số kết quả tối đa (cap cứng)
    max_recipients_per_draft: int     # Số recipients tối đa
    entry_id_max_length: int          # Độ dài EntryID tối đa

@dataclass(frozen=True)
class AuditConfig:
    """Cấu hình audit logging"""
    log_dir: str          # Thư mục chứa log file
    retain_days: int      # Số ngày giữ log
    hash_algorithm: str   # Thuật toán hash (mặc định: sha256)

@dataclass(frozen=True)
class LimitsConfig:
    """Giới hạn kích thước và tốc độ"""
    search_query_max_length: int        # Max 200
    email_body_max_length: int          # Max 50000
    subject_max_length: int             # Max 500
    list_emails_default_limit: int      # Mặc định 20
    list_emails_max_limit: int          # Tối đa 100
    max_calls_per_minute: int           # Rate limit 60
    com_operation_timeout_seconds: int  # Timeout 30 giây

@dataclass(frozen=True)
class AppConfig:
    """Config tổng hợp toàn bộ hệ thống — singleton"""
    outlook: OutlookConfig
    security: SecurityConfig
    audit: AuditConfig
    limits: LimitsConfig
```

### Hàm Chính

```python
def load_config(config_path: Path | None = None) -> AppConfig:
    """
    Đọc config.toml, validate và trả về AppConfig singleton.
    Ném ConfigError nếu file không tìm thấy hoặc giá trị không hợp lệ.
    Gọi một lần khi khởi động server — không gọi lại trong quá trình chạy.
    config_path mặc định là thư mục chứa server.py / config.toml.
    """

def get_config() -> AppConfig:
    """
    Lấy AppConfig singleton đã load.
    Ném RuntimeError nếu load_config() chưa được gọi.
    """
```

### Config.toml Reference — Bảng Đầy Đủ

| Section | Key | Kiểu | Mặc Định | Mô Tả |
|---|---|---|---|---|
| `[outlook]` | `account_name` | string | (bắt buộc) | Địa chỉ email tài khoản Outlook |
| `[outlook]` | `pst_display_name` | string | `""` | Tên hiển thị PST file, bỏ trống = mailbox mặc định |
| `[security]` | `read_only_mode` | boolean | `true` | Chặn compose/reply khi `true` — khuyến nghị giữ `true` |
| `[security]` | `allowed_folders` | array[string] | `["Inbox"]` | Danh sách thư mục Claude được phép truy cập |
| `[security]` | `max_results` | integer | `50` | Giới hạn kết quả trả về, tối đa tuyệt đối 200 |
| `[security]` | `max_recipients_per_draft` | integer | `20` | Số địa chỉ nhận tối đa khi compose |
| `[security]` | `entry_id_max_length` | integer | `256` | Độ dài Entry ID tối đa (hex string) |
| `[audit]` | `log_dir` | string | `"logs"` | Thư mục chứa audit log (tương đối với thư mục dự án) |
| `[audit]` | `retain_days` | integer | `90` | Số ngày giữ log (banking compliance: 90 ngày) |
| `[audit]` | `hash_algorithm` | string | `"sha256"` | Thuật toán hash cho params trong log |
| `[limits]` | `search_query_max_length` | integer | `200` | Độ dài tối đa query tìm kiếm |
| `[limits]` | `email_body_max_length` | integer | `50000` | Độ dài tối đa body email khi compose |
| `[limits]` | `subject_max_length` | integer | `500` | Độ dài tối đa tiêu đề email |
| `[limits]` | `list_emails_default_limit` | integer | `20` | Số email mặc định mỗi lần list |
| `[limits]` | `list_emails_max_limit` | integer | `100` | Số email tối đa mỗi lần list |
| `[limits]` | `max_calls_per_minute` | integer | `60` | Rate limit — số tool calls tối đa mỗi phút |
| `[limits]` | `com_operation_timeout_seconds` | integer | `30` | Timeout cho mỗi COM operation (giây) |

---

## 4. outlook_com.py — COM Wrapper, Dataclasses, Methods

### Mục Đích

`outlook_com.py` bọc toàn bộ việc giao tiếp với Outlook Desktop thông qua COM (Component Object Model — giao thức tích hợp ứng dụng Windows). File này đảm bảo:
- Tất cả COM calls chạy trong một STA thread (Single-Threaded Apartment — mô hình threading COM) duy nhất
- COM objects được giải phóng đúng thứ tự sau mỗi operation
- Lỗi COM được sanitize (làm sạch) trước khi trả ra ngoài

### Dataclasses Kết Quả

```python
@dataclass
class FolderInfo:
    """Thông tin một thư mục email"""
    name: str           # Tên thư mục hiển thị
    path: str           # Đường dẫn logic (không phải đường dẫn PST)
    unread_count: int   # Số email chưa đọc
    total_count: int    # Tổng số email

@dataclass
class EmailSummary:
    """Thông tin tóm tắt email (không có body) — dùng trong list_emails"""
    entry_id: str         # Outlook Entry ID dạng hex
    subject: str          # Tiêu đề email
    sender_name: str      # Tên người gửi
    sender_email: str     # Email người gửi
    received_time: str    # Thời gian nhận (ISO 8601)
    has_attachment: bool  # Có đính kèm không
    is_read: bool         # Đã đọc chưa
    size_kb: int          # Kích thước email (KB)

@dataclass
class AttachmentInfo:
    """Thông tin file đính kèm"""
    name: str        # Tên file
    size_kb: int     # Kích thước (KB)
    extension: str   # Phần mở rộng (.pdf, .xlsx...)

@dataclass
class EmailDetail:
    """Nội dung đầy đủ email — dùng trong read_email"""
    subject: str                    # Tiêu đề
    sender_name: str                # Tên người gửi
    sender_email: str               # Email người gửi
    to_recipients: list[str]        # Danh sách người nhận chính
    cc_recipients: list[str]        # Danh sách CC
    received_time: str              # Thời gian nhận (ISO 8601)
    body_text: str                  # Body đã strip HTML thành plain text
    attachments: list[AttachmentInfo]  # Danh sách file đính kèm

@dataclass
class DraftResult:
    """Kết quả tạo draft email"""
    status: str       # "draft_opened" hoặc "reply_opened"
    message: str      # Thông báo tiếng Việt cho người dùng
    draft_entry_id: str  # Entry ID của draft vừa tạo
```

### STA Thread Setup

```python
# Hàng đợi dùng để gửi tasks vào STA thread từ asyncio
_com_task_queue: Queue = Queue()

def _sta_worker():
    """
    Vòng lặp chính của STA thread — xử lý TẤT CẢ COM operations.
    Phải gọi CoInitialize() TRƯỚC bất kỳ win32com call nào.
    Chạy mãi cho đến khi nhận sentinel value (None).
    """

# STA thread được khởi động khi module load
_sta_thread: threading.Thread  # daemon=True, name='OutlookCOMThread'

def dispatch_to_sta(func: Callable, *args, **kwargs) -> concurrent.futures.Future:
    """
    Gửi function vào STA thread để thực thi an toàn với COM.
    Trả về Future — caller dùng future.result(timeout=...) để lấy kết quả.
    """
```

### Class OutlookCOMBridge

Context manager (trình quản lý ngữ cảnh) bọc toàn bộ lifecycle của một COM session. Mỗi tool call tạo một instance mới.

```python
class OutlookCOMBridge:
    """
    Context manager cho một phiên làm việc với Outlook COM.
    Tự động giải phóng tất cả COM objects khi thoát (kể cả khi có exception).
    KHÔNG tạo instance Outlook mới — chỉ kết nối đến Outlook đang chạy.
    """

    def __enter__(self) -> 'OutlookCOMBridge':
        """
        Kết nối đến Outlook đang chạy qua GetActiveObject().
        Ném OutlookNotRunningError nếu Outlook chưa được mở.
        """

    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        """
        Giải phóng tất cả COM objects theo thứ tự ngược.
        Gọi gc.collect() sau khi release để thu hồi references còn sót.
        Trả về False — không suppress exceptions.
        """
```

### Các Method Chính

```python
def get_namespace(self) -> win32com.client.Dispatch:
    """
    Lấy MAPI namespace (không gian làm việc chính của Outlook).
    Kết quả được thêm vào self._refs để tự động release sau.
    """

def get_folder_by_allowlist_name(self, name: str) -> win32com.client.Dispatch:
    """
    Resolve thư mục theo tên đã qua allowlist validation.
    Xử lý tên mặc định theo ngôn ngữ OS (dùng GetDefaultFolder cho Inbox, Sent, Drafts).
    Xử lý path lồng nhau: "Inbox/Projects".
    Xử lý wildcard một cấp: "Inbox/*".
    Ném FolderNotFoundError nếu thư mục không tồn tại trong Outlook.
    """

def list_mail_items(self, folder, limit: int, offset: int,
                    unread_only: bool = False) -> tuple[list[EmailSummary], int]:
    """
    Liệt kê emails trong folder với phân trang (paging).
    Trả về (danh_sách_email, tổng_số_email).
    Sắp xếp: mới nhất trước (sort by ReceivedTime descending).
    """

def get_mail_by_entry_id(self, entry_id: str) -> EmailDetail:
    """
    Đọc email đầy đủ theo Entry ID.
    Tự động strip HTML từ body bằng BeautifulSoup.
    Verify thư mục chứa email thuộc allowlist trước khi trả về.
    Ném EmailNotFoundError nếu Entry ID không hợp lệ.
    Ném FolderNotAllowedError nếu email nằm ngoài allowlist.
    """

def search_items(self, folders: list, dasl_filter: str,
                 limit: int) -> list[EmailSummary]:
    """
    Tìm kiếm emails dùng Items.Restrict() với DASL filter.
    Chỉ tìm trong folders thuộc allowlist.
    Không dùng vòng lặp Python thủ công — dùng DASL để Outlook filter.
    """

def create_draft(self, to: list[str], cc: list[str], subject: str,
                 body: str, importance: str) -> DraftResult:
    """
    Tạo MailItem nháp và mở cửa sổ Outlook bằng Display().
    TUYỆT ĐỐI KHÔNG gọi Send().
    Ném ReadOnlyModeError nếu config.security.read_only_mode = true.
    """

def create_reply_draft(self, entry_id: str, body: str,
                       reply_all: bool, additional_cc: list[str]) -> DraftResult:
    """
    Tạo reply cho email gốc và mở cửa sổ Outlook bằng Display().
    Gọi Reply() hoặc ReplyAll() rồi mới Display() — KHÔNG Send().
    Ném ReadOnlyModeError nếu config.security.read_only_mode = true.
    """
```

### Error Classes (Lớp Lỗi)

```python
class OutlookError(Exception):           # Lớp cha tất cả lỗi Outlook
class OutlookNotRunningError(OutlookError):   # Outlook chưa được mở
class FolderNotFoundError(OutlookError):      # Thư mục không tồn tại
class FolderNotAllowedError(OutlookError):    # Thư mục ngoài allowlist
class EmailNotFoundError(OutlookError):       # Entry ID không tìm thấy
class OutlookOperationError(OutlookError):    # COM operation thất bại (đã sanitize)
class ReadOnlyModeError(OutlookError):        # Cố compose khi read_only=true
```

### Hàm Tiện Ích Nội Bộ

```python
def _safe_com_call(operation_name: str, func: Callable, *args) -> Any:
    """
    Wrapper bắt lỗi COM (pywintypes.error, pythoncom.error).
    Log chi tiết nội bộ để debug — KHÔNG expose HRESULT ra ngoài.
    Chuyển mọi COM exception thành OutlookOperationError an toàn.
    """

def _strip_html(html_body: str) -> str:
    """
    Dùng BeautifulSoup để strip HTML tags, giữ lại text thuần túy.
    Xử lý encoding edge cases (tiếng Việt, emoji).
    """
```

---

## 5. security/ — Chi Tiết Từng File

### 5.1 security/credential.py

**Mục đích:** Quản lý API key Anthropic thông qua Windows Credential Manager (kho lưu trữ mật khẩu tích hợp của Windows). API key **không bao giờ** được lưu trong file, biến môi trường, hay bộ nhớ cache.

**Hàm chính:**

```python
class CredentialManager:
    """
    Wrapper cho Windows Credential Manager (DPAPI backend).
    Bắt buộc dùng WinVaultKeyring — raise lỗi ngay nếu chạy không phải Windows.
    """

    SERVICE_NAME: ClassVar[str] = "OutlookMCPSecure"
    USERNAME: ClassVar[str] = "anthropic_api_key"

    def __init__(self, audit: AuditLogger):
        """
        Kiểm tra keyring backend là WinVaultKeyring.
        Ném CredentialBackendError nếu không phải Windows Credential Manager.
        """

    def get_api_key(self) -> str:
        """
        Đọc API key từ Windows Credential Manager.
        KHÔNG cache kết quả — gọi keyring mỗi lần.
        Ghi audit log mỗi lần truy cập (event: credential_access).
        Ném CredentialNotFoundError nếu chưa setup.
        """

    def set_api_key(self, api_key: str) -> None:
        """
        Lưu API key vào Windows Credential Manager.
        Chỉ gọi từ setup.ps1 lúc cài đặt lần đầu — không gọi từ server.
        Ném ValueError nếu api_key rỗng hoặc không đúng định dạng.
        """

    def delete_api_key(self) -> None:
        """
        Xóa API key khỏi Windows Credential Manager.
        Dùng khi uninstall hoặc rotate key.
        """
```

**Error Classes:**

```python
class CredentialError(Exception):               # Lớp cha
class CredentialNotFoundError(CredentialError):  # API key chưa được setup
class CredentialBackendError(CredentialError):   # Không phải WinVaultKeyring
```

---

### 5.2 security/audit.py

**Mục đích:** Ghi audit log (nhật ký kiểm toán) dạng JSON Lines (mỗi dòng một JSON object) với đầy đủ thông tin về mọi tool call, theo dõi bất thường.

**Format File Log:**
- Vị trí: `%APPDATA%\OutlookMCPSecure\audit\audit-YYYY-MM-DD.jsonl`
- Encoding: UTF-8 không BOM
- Mỗi dòng: một JSON object kết thúc bằng `\n`
- File ACL: chỉ Owner + SYSTEM có quyền ghi

**Hàm chính:**

```python
class AuditLogger:
    """
    Ghi audit log JSON Lines với rotation theo ngày.
    Thread-safe — dùng threading.Lock để tránh log bị lẫn lộn.
    Tự động set ACL cho file log khi tạo mới.
    """

    def __init__(self, config: AuditConfig, session_id: str):
        """
        Tạo thư mục log nếu chưa có.
        Set ACL: chỉ Owner + SYSTEM có quyền ghi (dùng icacls).
        Xác minh ACL ngay sau khi set.
        """

    def log_server_start(self, version: str, read_only: bool,
                         allowlist_count: int) -> None:
        """Ghi entry server_start khi khởi động"""

    def log_server_stop(self, total_calls: int, error_count: int) -> None:
        """Ghi entry server_stop khi tắt"""

    def log_tool_start(self, tool_name: str, safe_params: dict) -> None:
        """
        Ghi entry bắt đầu tool call.
        safe_params chỉ chứa params an toàn (không có content email).
        """

    def log_tool_success(self, tool_name: str, meta: dict,
                         duration_ms: int) -> None:
        """Ghi entry tool call thành công với thời gian thực thi"""

    def log_tool_blocked(self, tool_name: str, block_reason: str,
                         risk_level: str = "medium") -> None:
        """
        Ghi entry bị từ chối (blocked).
        risk_level: "low" | "medium" | "high"
        """

    def log_tool_error(self, tool_name: str, error_type: str) -> None:
        """Ghi entry lỗi — không ghi exception message có thể chứa data nhạy cảm"""

    def log_credential_access(self, status: str) -> None:
        """Ghi mỗi lần đọc credential từ Windows Credential Manager"""

    def _write_entry(self, entry: dict) -> None:
        """
        Ghi một JSON entry vào file (thread-safe, append-only).
        Mỗi 100 entries ghi thêm 1 integrity checksum entry.
        """

    def _rotate_if_needed(self) -> None:
        """Tạo file mới nếu qua ngày mới (rotation theo ngày)"""

    def _cleanup_old_logs(self) -> None:
        """Xóa file log cũ hơn retain_days ngày"""
```

**Cấu Trúc JSON Entry:**

```json
{
  "ts": "2026-06-24T09:01:15.456789+07:00",
  "session_id": "a1b2c3d4-...",
  "tool": "list_emails",
  "params": {"folder": "Inbox", "limit": 20},
  "status": "ok",
  "duration_ms": 234,
  "items_returned": 20
}
```

**Quy Tắc Bắt Buộc — Không Ghi Vào Log:**
- Subject email
- Body email
- Tên/địa chỉ người gửi, người nhận
- Tên file đính kèm
- API key dưới mọi hình thức
- Stack trace Python (chỉ ghi error_type)

---

### 5.3 security/validator.py

**Mục đích:** Validate và sanitize (làm sạch) toàn bộ input từ Claude trước khi truyền vào bất kỳ operation nào. Đây là "cổng bảo vệ" đầu tiên.

**Hàm chính:**

```python
class InputValidator:
    """
    Validate và sanitize input từ MCP tool calls.
    Tất cả hàm raise ValidationError nếu input không hợp lệ.
    """

    def __init__(self, config: SecurityConfig):
        """Nhận config để biết allowlist, giới hạn kích thước..."""

    def validate_folder_name(self, folder_path: str) -> str:
        """
        Kiểm tra folder_path hợp lệ và thuộc allowlist.
        Các bước xử lý theo thứ tự:
          1. Loại bỏ null bytes và ký tự điều khiển
          2. Chuẩn hóa Unicode NFC
          3. casefold() (lowercase unicode-aware, xử lý đúng tiếng Việt)
          4. Kiểm tra không chứa "../", "://", ":\\"
          5. So sánh exact match với allowed_folders (đã casefold)
          6. Trả về tên gốc (chưa casefold) nếu pass
        Ném FolderNotAllowedError nếu không thuộc allowlist.
        """

    def validate_email_id(self, entry_id: str) -> str:
        """
        Kiểm tra entry_id là hex string hợp lệ.
        Regex: ^[0-9A-Fa-f]+$ và độ dài <= entry_id_max_length.
        Strip null bytes trước khi check.
        Ném ValidationError nếu không hợp lệ.
        """

    def validate_search_query(self, query: str) -> str:
        """
        Làm sạch query tìm kiếm.
        - Giới hạn độ dài theo config.limits.search_query_max_length
        - Strip null bytes và ký tự điều khiển
        - Escape dấu nháy đơn (để tránh injection vào DASL)
        - Không cho phép các pattern DASL injection: OR, AND, LIKE ở đầu query
        Trả về query đã sanitize.
        """

    def validate_email_address(self, email: str) -> str:
        """
        Kiểm tra địa chỉ email hợp lệ theo RFC5322 (chuẩn định dạng email).
        Regex pattern chuẩn, không quá nghiêm ngặt (chấp nhận unicode domain).
        Ném ValidationError nếu không hợp lệ.
        """

    def validate_email_list(self, emails: list[str],
                            max_count: int | None = None) -> list[str]:
        """
        Validate danh sách địa chỉ email.
        max_count mặc định lấy từ config.security.max_recipients_per_draft.
        Ném ValidationError nếu list rỗng hoặc vượt max_count.
        """

    def validate_body(self, body: str) -> str:
        """
        Kiểm tra body email không vượt config.limits.email_body_max_length.
        Strip null bytes.
        Ném ValidationError nếu vượt giới hạn.
        """

    def validate_subject(self, subject: str) -> str:
        """
        Kiểm tra tiêu đề không vượt config.limits.subject_max_length.
        Strip null bytes và ký tự xuống dòng.
        """

    @staticmethod
    def sanitize_string(text: str) -> str:
        """
        Loại bỏ null bytes (\\x00) và control characters (\\x01-\\x1f, trừ \\t, \\n, \\r).
        Dùng cho mọi string input trước các bước validate cụ thể.
        """
```

**Error Classes:**

```python
class ValidationError(Exception):          # Lớp cha mọi lỗi validate
class FolderNotAllowedError(ValidationError):  # Folder ngoài allowlist
```

---

## 6. tools/ — Mỗi Tool File, Mỗi Function

### 6.1 tools/list_folders.py

**Tool:** `list_folders`
**Mô tả:** Liệt kê các thư mục email mà Claude được phép truy cập.

```python
def handle_list_folders(args: dict, bridge: OutlookCOMBridge,
                        config: AppConfig) -> dict:
    """
    Xử lý tool call list_folders.
    Nhận: args = {"include_subfolders": bool}
    Trả về: {"folders": [FolderInfo dạng dict]}
    
    Luồng xử lý:
      1. Lấy include_subfolders từ args (mặc định False)
      2. Với mỗi folder trong config.security.allowed_folders:
         a. Gọi bridge.get_folder_by_allowlist_name(name)
         b. Nếu include_subfolders=True, đệ quy một cấp con
      3. Trả về danh sách FolderInfo
    Không bao giờ trả về folder ngoài allowlist, dù COM trả về nhiều hơn.
    """
```

**JSON Schema Tool:**
```json
{
  "type": "object",
  "properties": {
    "include_subfolders": {"type": "boolean", "default": false}
  },
  "required": [],
  "additionalProperties": false
}
```

---

### 6.2 tools/read_email.py

File này chứa **hai** tool handlers: `list_emails` và `read_email`.

**Tool: `list_emails`**

```python
def handle_list_emails(args: dict, validator: InputValidator,
                       bridge: OutlookCOMBridge, config: AppConfig) -> dict:
    """
    Xử lý tool call list_emails.
    Nhận: args = {"folder_path": str, "limit": int, "offset": int, "unread_only": bool}
    Trả về: {"emails": [EmailSummary dạng dict], "total": int}
    
    Luồng xử lý:
      1. validate_folder_name(folder_path)
      2. Cap limit tại config.security.max_results (không tin giá trị client)
      3. Gọi bridge.list_mail_items(folder, limit, offset, unread_only)
      4. Chuyển EmailSummary thành dict để serialize JSON
    Audit params: folder_name, items_returned — KHÔNG ghi subject/sender.
    """
```

**JSON Schema Tool `list_emails`:**
```json
{
  "type": "object",
  "properties": {
    "folder_path": {"type": "string", "maxLength": 260},
    "limit": {"type": "integer", "default": 20, "maximum": 100},
    "offset": {"type": "integer", "default": 0},
    "unread_only": {"type": "boolean", "default": false}
  },
  "required": ["folder_path"],
  "additionalProperties": false
}
```

**Tool: `read_email`**

```python
def handle_read_email(args: dict, validator: InputValidator,
                      bridge: OutlookCOMBridge) -> dict:
    """
    Xử lý tool call read_email.
    Nhận: args = {"entry_id": str}
    Trả về: EmailDetail dạng dict (body đã strip HTML)
    
    Luồng xử lý:
      1. validate_email_id(entry_id)
      2. Gọi bridge.get_mail_by_entry_id(entry_id)
         (Bridge tự verify folder thuộc allowlist)
      3. Chuyển EmailDetail thành dict
    Audit params: entry_id[:8] — KHÔNG ghi body/subject/sender.
    """
```

**JSON Schema Tool `read_email`:**
```json
{
  "type": "object",
  "properties": {
    "entry_id": {
      "type": "string",
      "pattern": "^[0-9A-Fa-f]+$",
      "maxLength": 256
    }
  },
  "required": ["entry_id"],
  "additionalProperties": false
}
```

---

### 6.3 tools/search.py

**Tool:** `search_emails`
**Mô tả:** Tìm kiếm email trong các thư mục được phép bằng DASL filter.

```python
def handle_search_emails(args: dict, validator: InputValidator,
                         bridge: OutlookCOMBridge, config: AppConfig) -> dict:
    """
    Xử lý tool call search_emails.
    Nhận: args = {"query", "folder_path"?, "search_in", "date_from"?, "date_to"?, "limit"}
    Trả về: {"results": [EmailSummary + folder_path + snippet], "total_found": int}
    
    Luồng xử lý:
      1. validate_search_query(query) — sanitize DASL injection
      2. Nếu có folder_path: validate_folder_name(folder_path)
         Nếu không có: dùng tất cả allowed_folders
      3. validate_date(date_from), validate_date(date_to) nếu có
      4. Cap limit tại config.security.max_results
      5. Xây dựng DASL filter bằng build_dasl_filter()
      6. Gọi bridge.search_items(folders, dasl_filter, limit)
      7. Tạo snippet từ body (không expose toàn bộ body)
    Audit params: SHA256(query), folder_name — KHÔNG ghi query plaintext.
    """

def build_dasl_filter(query: str, search_in: str,
                      date_from: str | None, date_to: str | None) -> str:
    """
    Tạo DASL filter string cho Outlook Items.Restrict().
    DASL (DAV Searching and Locating) — ngôn ngữ tìm kiếm nội bộ của Outlook MAPI.
    Query đã được validate và escape single quotes trước khi gọi hàm này.
    
    search_in = "subject":  lọc theo urn:schemas:httpmail:subject
    search_in = "sender":   lọc theo urn:schemas:httpmail:fromemail
    search_in = "body":     lọc theo urn:schemas:httpmail:textdescription
    search_in = "all":      OR kết hợp subject và body
    
    Tự động thêm bộ lọc ngày nếu date_from hoặc date_to được cung cấp.
    """
```

**JSON Schema Tool `search_emails`:**
```json
{
  "type": "object",
  "properties": {
    "query": {"type": "string", "maxLength": 200},
    "folder_path": {"type": "string", "maxLength": 260},
    "search_in": {
      "type": "string",
      "enum": ["subject", "body", "sender", "all"],
      "default": "subject"
    },
    "date_from": {"type": "string", "format": "date"},
    "date_to": {"type": "string", "format": "date"},
    "limit": {"type": "integer", "default": 20, "maximum": 50}
  },
  "required": ["query"],
  "additionalProperties": false
}
```

---

### 6.4 tools/compose.py

File này chứa **hai** tool handlers: `compose_draft` và `reply_draft`. Cả hai đều bị khóa hoàn toàn khi `config.security.read_only_mode = true`.

**Tool: `compose_draft`**

```python
def handle_compose_draft(args: dict, validator: InputValidator,
                         bridge: OutlookCOMBridge, config: AppConfig) -> dict:
    """
    Xử lý tool call compose_draft.
    Nhận: args = {"to": list[str], "cc"?: list[str], "subject": str,
                  "body": str, "importance"?: str}
    Trả về: DraftResult dạng dict
    
    Luồng xử lý:
      1. Kiểm tra read_only_mode = false, nếu true ném ReadOnlyModeError
      2. validate_email_list(to) — tối đa max_recipients_per_draft địa chỉ
      3. validate_email_list(cc) nếu có
      4. validate_subject(subject)
      5. validate_body(body)
      6. Kiểm tra importance trong ["low", "normal", "high"]
      7. Gọi bridge.create_draft(to, cc, subject, body, importance)
         Bridge gọi Display() — KHÔNG Send()
    Audit params: SHA256(to_addresses), SHA256(subject) — KHÔNG ghi body.
    """
```

**JSON Schema Tool `compose_draft`:**
```json
{
  "type": "object",
  "properties": {
    "to": {
      "type": "array",
      "items": {"type": "string"},
      "minItems": 1,
      "maxItems": 50
    },
    "cc": {
      "type": "array",
      "items": {"type": "string"},
      "maxItems": 50
    },
    "subject": {"type": "string", "maxLength": 500},
    "body": {"type": "string", "maxLength": 50000},
    "importance": {
      "type": "string",
      "enum": ["low", "normal", "high"],
      "default": "normal"
    }
  },
  "required": ["to", "subject", "body"],
  "additionalProperties": false
}
```

**Tool: `reply_draft`**

```python
def handle_reply_draft(args: dict, validator: InputValidator,
                       bridge: OutlookCOMBridge, config: AppConfig) -> dict:
    """
    Xử lý tool call reply_draft.
    Nhận: args = {"entry_id": str, "body": str,
                  "reply_all"?: bool, "additional_cc"?: list[str]}
    Trả về: DraftResult dạng dict
    
    Luồng xử lý:
      1. Kiểm tra read_only_mode = false, nếu true ném ReadOnlyModeError
      2. validate_email_id(entry_id)
      3. validate_body(body)
      4. validate_email_list(additional_cc) nếu có (tối đa 20 địa chỉ)
      5. Gọi bridge.create_reply_draft(entry_id, body, reply_all, additional_cc)
         Bridge gọi Reply() hoặc ReplyAll() rồi Display() — KHÔNG Send()
    Audit params: SHA256(entry_id), action_type — KHÔNG ghi body.
    """
```

**JSON Schema Tool `reply_draft`:**
```json
{
  "type": "object",
  "properties": {
    "entry_id": {
      "type": "string",
      "pattern": "^[0-9A-Fa-f]+$",
      "maxLength": 256
    },
    "body": {"type": "string", "maxLength": 50000},
    "reply_all": {"type": "boolean", "default": false},
    "additional_cc": {
      "type": "array",
      "items": {"type": "string"},
      "maxItems": 20
    }
  },
  "required": ["entry_id", "body"],
  "additionalProperties": false
}
```

---

## 7. Configuration Reference — Bảng Tất Cả Settings

Đây là bảng tham chiếu đầy đủ tất cả settings trong `config.toml`, bao gồm giá trị mặc định, giới hạn tối đa và mô tả chi tiết.

### Nhóm [outlook]

| Setting | Kiểu | Bắt Buộc | Mặc Định | Mô Tả |
|---|---|---|---|---|
| `account_name` | string | Có | — | Địa chỉ email tài khoản Outlook (ví dụ: `thanhnt@softmart.net.vn`) |
| `pst_display_name` | string | Không | `""` | Tên PST file nếu dùng nhiều hộp thư. Bỏ trống = dùng mailbox mặc định |

### Nhóm [security]

| Setting | Kiểu | Bắt Buộc | Mặc Định | Giới Hạn | Mô Tả |
|---|---|---|---|---|---|
| `read_only_mode` | boolean | Không | `true` | — | `true` = khóa toàn bộ compose/reply. **Khuyến nghị giữ `true`** cho đến khi tin tưởng hệ thống |
| `allowed_folders` | array[string] | Có | — | — | Danh sách thư mục Claude được phép đọc. Hỗ trợ path lồng nhau (`Inbox/Projects`) và wildcard một cấp (`Inbox/*`) |
| `max_results` | integer | Không | `50` | 1–200 | Số kết quả tối đa server trả về, bất kể client yêu cầu bao nhiêu |
| `max_recipients_per_draft` | integer | Không | `20` | 1–50 | Số địa chỉ email nhận tối đa khi compose draft |
| `entry_id_max_length` | integer | Không | `256` | 64–512 | Độ dài tối đa Entry ID (Outlook dùng hex string, thực tế ~120 ký tự) |

### Nhóm [audit]

| Setting | Kiểu | Bắt Buộc | Mặc Định | Mô Tả |
|---|---|---|---|---|
| `log_dir` | string | Không | `"logs"` | Thư mục chứa audit log. Đường dẫn tương đối so với thư mục dự án |
| `retain_days` | integer | Không | `90` | Số ngày giữ log. 90 ngày theo tiêu chuẩn banking compliance (tuân thủ ngân hàng) |
| `hash_algorithm` | string | Không | `"sha256"` | Thuật toán hash cho params nhạy cảm trong log. Chỉ hỗ trợ `sha256` |

### Nhóm [limits]

| Setting | Kiểu | Bắt Buộc | Mặc Định | Giới Hạn | Mô Tả |
|---|---|---|---|---|---|
| `search_query_max_length` | integer | Không | `200` | 50–500 | Độ dài tối đa chuỗi tìm kiếm |
| `email_body_max_length` | integer | Không | `50000` | 1000–200000 | Độ dài tối đa body email khi compose/reply |
| `subject_max_length` | integer | Không | `500` | 50–998 | Độ dài tối đa tiêu đề email (RFC 5321 giới hạn 998) |
| `list_emails_default_limit` | integer | Không | `20` | 1–100 | Số email trả về mặc định khi client không chỉ định |
| `list_emails_max_limit` | integer | Không | `100` | 1–200 | Số email tối đa một lần gọi list_emails |
| `max_calls_per_minute` | integer | Không | `60` | 10–300 | Rate limiting (giới hạn tốc độ) — số tool calls tối đa mỗi phút |
| `com_operation_timeout_seconds` | integer | Không | `30` | 5–120 | Timeout cho mỗi COM operation. Nếu Outlook không phản hồi trong thời gian này, server trả lỗi |

---

## 8. Dependencies — requirements.txt Giải Thích Từng Package

Tất cả version phải PIN CHÍNH XÁC (dùng `==`) để tránh supply chain attack (tấn công chuỗi cung ứng — hacker chèn mã độc vào phiên bản mới của thư viện).

### Dependency Bắt Buộc (Production)

| Package | Version | Mục Đích | Lý Do Chọn |
|---|---|---|---|
| `mcp` | `==1.0.0` | MCP SDK — giao thức kết nối Claude với tools bên ngoài | Official SDK từ Anthropic, hỗ trợ stdio transport |
| `pywin32` | `==306` | Outlook COM automation — giao tiếp với Outlook Desktop qua Windows COM | Thư viện chuẩn nhất cho COM trên Python Windows. Cung cấp `win32com.client`, `pythoncom`, `pywintypes` |
| `keyring` | `==25.0.0` | Windows Credential Manager integration — lưu/đọc API key an toàn | Hỗ trợ WinVaultKeyring backend (DPAPI encryption) trên Windows |
| `tomli` | `==2.0.1` | TOML config parsing — đọc `config.toml` | Backport cho Python < 3.11. Python >= 3.11 dùng `tomllib` stdlib, không cần package này |
| `beautifulsoup4` | `==4.12.3` | HTML strip — làm sạch body email trước khi trả về Claude | An toàn hơn regex thủ công, xử lý đúng encoding và edge cases HTML |

### Dependency Development (Không Cần Trong Production)

| Package | Version | Mục Đích |
|---|---|---|
| `pytest` | `==8.2.0` | Framework viết và chạy unit test |
| `pytest-asyncio` | `==0.23.6` | Hỗ trợ test async functions (cần để test `server.py`) |

### Quy Trình Tạo Lock File

Sau khi cài đặt `pip-tools`, chạy lệnh sau để tạo lock file với SHA-256 hash cho từng package:

```powershell
pip-compile --generate-hashes requirements.txt -o requirements.lock
```

Khi cài đặt, dùng lock file để đảm bảo đúng hash:

```powershell
pip install --require-hashes -r requirements.lock
```

### Thư Viện Chuẩn Python Sử Dụng (Không Cần Cài Thêm)

| Module Stdlib | Dùng Trong | Mục Đích |
|---|---|---|
| `tomllib` | `config.py` | Đọc TOML (Python >= 3.11 — dùng `tomli` nếu Python cũ hơn) |
| `asyncio` | `server.py` | Event loop cho MCP server async |
| `concurrent.futures` | `server.py`, `outlook_com.py` | `ThreadPoolExecutor` và `Future` cho COM thread |
| `threading` | `outlook_com.py`, `security/audit.py` | `Thread` cho STA worker, `Lock` cho audit logger |
| `queue` | `outlook_com.py` | `Queue` cho COM task queue |
| `json` | `security/audit.py`, `server.py` | Serialize/deserialize JSON |
| `pathlib` | `config.py`, `security/audit.py` | Xử lý đường dẫn file an toàn |
| `re` | `security/validator.py` | Regex validate entry_id, email address |
| `unicodedata` | `security/validator.py` | NFC normalize tên thư mục |
| `hashlib` | `security/audit.py`, `security/validator.py` | SHA-256 hash cho audit log |
| `datetime` | `security/audit.py`, `outlook_com.py` | Timestamp, log rotation |
| `dataclasses` | `config.py`, `outlook_com.py` | Dataclass cho config, kết quả trả về |
| `gc` | `outlook_com.py` | Garbage collection sau khi release COM objects |
| `logging` | Tất cả | Internal debug logger (không phải audit log) |
| `time` | `server.py` | Đo thời gian thực thi (duration_ms) |
| `uuid` | `server.py` | Tạo session_id duy nhất mỗi lần khởi động |

---

## Sơ Đồ Phụ Thuộc Giữa Các Thành Phần

```
Claude Code CLI
      |
      | stdio (MCP protocol)
      |
  server.py
  ├── config.py ─────────────────── config.toml
  ├── security/
  │   ├── credential.py ─────────── Windows Credential Manager
  │   ├── audit.py ──────────────── logs/audit-YYYY-MM-DD.jsonl
  │   └── validator.py
  ├── tools/
  │   ├── list_folders.py
  │   ├── read_email.py
  │   ├── search.py
  │   └── compose.py
  └── outlook_com.py ─────────────── Outlook.exe (COM / MAPI)
```

Nguyên tắc phụ thuộc một chiều:
- `server.py` biết tất cả, nhưng không module nào import `server.py`
- `outlook_com.py` không biết đến security layer — chỉ làm việc với COM
- `security/validator.py` không import `outlook_com.py` — validate thuần túy input
- `tools/*.py` không giao tiếp trực tiếp với nhau — tất cả đi qua `server.py`

---

*Tài liệu này mô tả thiết kế chi tiết của hệ thống Claude-Outlook MCP Secure. Mọi thay đổi interface phải cập nhật tài liệu này đồng thời.*
