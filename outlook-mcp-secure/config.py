"""
config.py — Quản lý cấu hình cho Claude-Outlook MCP Secure

Mục đích:
    Module này chịu trách nhiệm đọc, kiểm tra và cung cấp cấu hình cho toàn bộ
    hệ thống MCP server. Cấu hình có thể lấy từ file config.toml (người dùng
    tùy chỉnh) hoặc dùng giá trị mặc định an toàn nếu không có file config.

Cơ chế hoạt động:
    1. Định nghĩa lớp Config với tất cả các tham số cấu hình và giá trị mặc định
    2. Hàm load_config() đọc config.toml nếu file tồn tại, rồi ghi đè lên mặc định
    3. Sau khi load, gọi validate() để kiểm tra ràng buộc bảo mật
    4. Tạo thư mục lưu audit log nếu chưa tồn tại

Bảo mật:
    - KHÔNG lưu credentials trong file config — dùng Windows Credential Manager
    - allowed_folders phải được chỉ định rõ (không cho phép wildcard **)
    - Khi read_only_mode = True, bất kỳ thao tác ghi nào đều bị từ chối

Phụ thuộc:
    - tomllib (Python 3.11+ built-in) — đọc file TOML
    - dataclasses — định nghĩa cấu trúc dữ liệu cấu hình
    - pathlib — xử lý đường dẫn file an toàn
"""

from __future__ import annotations

import logging
import os
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

# Dùng tomllib (Python 3.11+ built-in) để đọc file TOML
try:
    import tomllib  # Python 3.11+
except ImportError:
    # Fallback cho Python 3.10 trở xuống — cần cài tomli từ pip
    try:
        import tomli as tomllib  # type: ignore[no-reuse-def]
    except ImportError as exc:
        raise ImportError(
            "Không tìm thấy thư viện đọc TOML. "
            "Python 3.11+ có sẵn tomllib. "
            "Python 3.10 trở xuống cần chạy: pip install tomli"
        ) from exc

# Logger nội bộ — chỉ ghi ra console/file log của server, không trả về Claude
_logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# Các hằng số giới hạn an toàn — không cho phép vượt quá dù config.toml chỉ định
# ─────────────────────────────────────────────────────────────────────────────

# Số lượng email tối đa tuyệt đối được trả về trong một lần gọi
_HARD_CAP_MAX_EMAILS: int = 200

# Số ký tự tối đa tuyệt đối của nội dung email được trả về
_HARD_CAP_MAX_BODY_CHARS: int = 100_000

# Số lượng tối đa người nhận trong một bản nháp
_HARD_CAP_MAX_RECIPIENTS: int = 100

# Số lần gọi tool tối đa mỗi phút (rate limiting)
_HARD_CAP_MAX_CALLS_PER_MINUTE: int = 120

# Độ dài tối đa của entry_id email (hex string)
_HARD_CAP_ENTRY_ID_MAX_LENGTH: int = 512

# Độ dài tối đa của search query
_HARD_CAP_SEARCH_QUERY_MAX_LENGTH: int = 500


# ─────────────────────────────────────────────────────────────────────────────
# Các lớp ngoại lệ dành riêng cho config
# ─────────────────────────────────────────────────────────────────────────────

class ConfigError(Exception):
    """Lỗi cấu hình — xảy ra khi config.toml có giá trị không hợp lệ hoặc thiếu bắt buộc."""


class ConfigValidationError(ConfigError):
    """Lỗi kiểm tra cấu hình — vi phạm ràng buộc bảo mật hoặc logic nghiệp vụ."""


# ─────────────────────────────────────────────────────────────────────────────
# Lớp cấu hình chính
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class Config:
    """
    Lớp cấu hình trung tâm — chứa tất cả tham số điều khiển hành vi của MCP server.

    Giá trị mặc định được thiết kế theo nguyên tắc "an toàn nhất" (secure by default):
    - READ_ONLY_MODE = True: mặc định chỉ đọc, không soạn/gửi email
    - ALLOWED_FOLDERS giới hạn rõ ràng: chỉ Inbox và Hộp thư đến
    - Không có credential nào được lưu trong lớp này

    Attributes:
        ACCOUNT_NAME: Tên tài khoản email trong Outlook (hiển thị trong Outlook Account Settings)
        ALLOWED_FOLDERS: Danh sách tên thư mục Claude được phép truy cập
        MAX_EMAILS_PER_REQUEST: Số email tối đa trả về trong một lần gọi list_emails
        MAX_BODY_CHARS: Số ký tự tối đa của nội dung email trả về (cắt bớt nếu vượt)
        AUDIT_LOG_PATH: Đường dẫn file audit log JSON Lines
        READ_ONLY_MODE: Khi True, chặn toàn bộ thao tác soạn/gửi email
        KEYRING_SERVICE: Tên dịch vụ trong Windows Credential Manager
        MAX_RECIPIENTS_PER_DRAFT: Số người nhận tối đa trong một bản nháp email
        MAX_CALLS_PER_MINUTE: Giới hạn số lần gọi tool mỗi phút (rate limiting)
        AUDIT_RETAIN_DAYS: Số ngày giữ lại file audit log trước khi xóa tự động
        AUDIT_HASH_ALGORITHM: Thuật toán hash dùng để che giấu dữ liệu nhạy cảm trong log
        SEARCH_QUERY_MAX_LENGTH: Độ dài tối đa của câu truy vấn tìm kiếm
        ENTRY_ID_MAX_LENGTH: Độ dài tối đa của entry_id email (mã định danh nội bộ Outlook)
        COM_OPERATION_TIMEOUT_SECONDS: Timeout (giây) cho mỗi thao tác COM với Outlook
        PST_DISPLAY_NAME: Tên hiển thị của file PST (nếu dùng PST, để trống nếu không dùng)
        LOG_LEVEL: Mức độ ghi log nội bộ (DEBUG/INFO/WARNING/ERROR)
    """

    # ── Thông tin tài khoản Outlook ──
    # Tên tài khoản email — phải khớp với tên hiển thị trong Outlook
    ACCOUNT_NAME: str = "thanhnt@softmart.net.vn"

    # Tên hiển thị của PST file (để trống nếu chỉ dùng IMAP account)
    PST_DISPLAY_NAME: str = ""

    # ── Kiểm soát truy cập thư mục ──
    # Danh sách thư mục được phép truy cập — PHẢI được chỉ định rõ ràng
    # Hỗ trợ cả tên tiếng Việt (Hộp thư đến) và tên tiếng Anh (Inbox)
    # Hỗ trợ path lồng nhau: "Inbox/Projects"
    # Hỗ trợ wildcard 1 cấp: "Inbox/*" (không hỗ trợ ** đệ quy)
    ALLOWED_FOLDERS: list[str] = field(
        default_factory=lambda: ["Hộp thư đến", "Inbox"]
    )

    # ── Giới hạn dữ liệu trả về ──
    # Số email tối đa trả về trong một lần gọi list_emails
    MAX_EMAILS_PER_REQUEST: int = 50

    # Số ký tự tối đa của nội dung email — phần vượt quá sẽ bị cắt bớt
    MAX_BODY_CHARS: int = 10_000

    # ── Đường dẫn audit log ──
    # Dùng %APPDATA% để tránh lưu trong thư mục source code
    # Mặc định: C:\Users\<user>\AppData\Roaming\ClaudeOutlookMCP\audit.jsonl
    AUDIT_LOG_PATH: Path = field(
        default_factory=lambda: Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
        / "ClaudeOutlookMCP"
        / "audit.jsonl"
    )

    # ── Chế độ bảo mật ──
    # TRUE (mặc định và khuyến nghị): chỉ đọc email, chặn compose/reply
    # FALSE: cho phép soạn nháp email (vẫn cần người dùng xác nhận mới gửi)
    READ_ONLY_MODE: bool = True

    # ── Windows Credential Manager ──
    # Tên dịch vụ (service name) dùng khi lưu/đọc credential từ Windows Vault
    KEYRING_SERVICE: str = "ClaudeOutlookMCP"

    # ── Giới hạn soạn email ──
    # Số người nhận tối đa trong một bản nháp (to + cc gộp lại)
    MAX_RECIPIENTS_PER_DRAFT: int = 20

    # ── Rate limiting — giới hạn tần suất gọi tool ──
    # Số lần gọi tool tối đa mỗi phút trước khi trả lỗi RateLimitError
    MAX_CALLS_PER_MINUTE: int = 60

    # ── Lưu trữ audit log ──
    # Số ngày giữ file audit log (banking compliance thường yêu cầu 90 ngày)
    AUDIT_RETAIN_DAYS: int = 90

    # Thuật toán hash để che giấu dữ liệu nhạy cảm trong log (sha256 hoặc sha512)
    AUDIT_HASH_ALGORITHM: str = "sha256"

    # ── Giới hạn input ──
    # Độ dài tối đa của câu truy vấn tìm kiếm email
    SEARCH_QUERY_MAX_LENGTH: int = 200

    # Độ dài tối đa của entry_id — mã hex định danh nội bộ của Outlook
    ENTRY_ID_MAX_LENGTH: int = 256

    # ── COM operation timeout ──
    # Thời gian tối đa (giây) chờ mỗi thao tác COM với Outlook
    COM_OPERATION_TIMEOUT_SECONDS: int = 30

    # ── Logging nội bộ ──
    # Mức độ chi tiết của log nội bộ (không phải audit log)
    LOG_LEVEL: str = "INFO"

    # ── Giới hạn tìm kiếm ──
    # Số kết quả tối đa trả về từ search_emails
    SEARCH_MAX_RESULTS: int = 50

    # ── Phiên bản server ──
    # Dùng để ghi vào audit log khi server khởi động
    SERVER_VERSION: str = "1.0.0"


# ─────────────────────────────────────────────────────────────────────────────
# Hàm đọc và xây dựng Config từ file config.toml
# ─────────────────────────────────────────────────────────────────────────────

def load_from_toml(path: Path) -> dict[str, Any]:
    """
    Đọc file config.toml và trả về dictionary các cài đặt.

    Hàm này chỉ đọc và parse TOML — không validate. Gọi validate() sau khi build Config.

    Args:
        path: Đường dẫn tuyệt đối đến file config.toml

    Returns:
        Dictionary chứa tất cả cài đặt từ TOML, phẳng hoá từ các section lồng nhau.
        Trả về dict rỗng nếu file không tồn tại.

    Raises:
        ConfigError: Nếu file TOML có cú pháp sai hoặc không thể đọc được.
    """
    # Bước 1: Kiểm tra file có tồn tại không — nếu không thì dùng mặc định
    if not path.exists():
        _logger.info(
            "Không tìm thấy file config.toml tại '%s'. Dùng cấu hình mặc định.", path
        )
        return {}

    _logger.info("Đang đọc cấu hình từ file '%s'.", path)

    # Bước 2: Đọc và parse file TOML
    try:
        with open(path, "rb") as fh:  # TOML yêu cầu mở ở chế độ binary "rb"
            raw: dict[str, Any] = tomllib.load(fh)
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(
            f"File config.toml tại '{path}' có cú pháp sai: {exc}"
        ) from exc
    except OSError as exc:
        raise ConfigError(
            f"Không thể đọc file config.toml tại '{path}': {exc}"
        ) from exc

    # Bước 3: Phẳng hoá cấu trúc TOML lồng nhau thành dict phẳng
    # config.toml có các section [outlook], [security], [audit], [limits]
    # → chuyển thành dict với key viết HOA theo quy ước tên field của Config
    flattened: dict[str, Any] = {}

    # Section [outlook] — thông tin tài khoản
    if "outlook" in raw:
        outlook_section = raw["outlook"]
        _map_if_present(outlook_section, "account_name", flattened, "ACCOUNT_NAME")
        _map_if_present(outlook_section, "pst_display_name", flattened, "PST_DISPLAY_NAME")

    # Section [security] — cài đặt bảo mật
    if "security" in raw:
        sec = raw["security"]
        _map_if_present(sec, "read_only_mode", flattened, "READ_ONLY_MODE")
        _map_if_present(sec, "allowed_folders", flattened, "ALLOWED_FOLDERS")
        _map_if_present(sec, "max_results", flattened, "MAX_EMAILS_PER_REQUEST")
        _map_if_present(sec, "max_recipients_per_draft", flattened, "MAX_RECIPIENTS_PER_DRAFT")
        _map_if_present(sec, "entry_id_max_length", flattened, "ENTRY_ID_MAX_LENGTH")
        _map_if_present(sec, "keyring_service", flattened, "KEYRING_SERVICE")

    # Section [audit] — cài đặt audit logging
    if "audit" in raw:
        audit = raw["audit"]
        if "log_dir" in audit:
            # Chuyển log_dir (relative path) thành AUDIT_LOG_PATH tuyệt đối
            log_dir_raw = audit["log_dir"]
            flattened["AUDIT_LOG_PATH"] = _resolve_audit_log_path(log_dir_raw, path.parent)
        _map_if_present(audit, "retain_days", flattened, "AUDIT_RETAIN_DAYS")
        _map_if_present(audit, "hash_algorithm", flattened, "AUDIT_HASH_ALGORITHM")

    # Section [limits] — giới hạn tham số
    if "limits" in raw:
        limits = raw["limits"]
        _map_if_present(limits, "search_query_max_length", flattened, "SEARCH_QUERY_MAX_LENGTH")
        _map_if_present(limits, "email_body_max_length", flattened, "MAX_BODY_CHARS")
        _map_if_present(limits, "max_calls_per_minute", flattened, "MAX_CALLS_PER_MINUTE")
        _map_if_present(limits, "com_operation_timeout_seconds", flattened, "COM_OPERATION_TIMEOUT_SECONDS")
        _map_if_present(limits, "list_emails_max_limit", flattened, "MAX_EMAILS_PER_REQUEST")
        _map_if_present(limits, "search_max_results", flattened, "SEARCH_MAX_RESULTS")

    # Section [server] — thông tin server (tùy chọn)
    if "server" in raw:
        srv = raw["server"]
        _map_if_present(srv, "log_level", flattened, "LOG_LEVEL")
        _map_if_present(srv, "version", flattened, "SERVER_VERSION")

    # Bước 4: Kiểm tra không có section/key lạ ngoài schema đã biết
    # (theo nguyên tắc extra='forbid' — từ chối field không có trong schema)
    known_sections = {"outlook", "security", "audit", "limits", "server"}
    unknown_sections = set(raw.keys()) - known_sections
    if unknown_sections:
        raise ConfigError(
            f"File config.toml chứa section không hợp lệ: {sorted(unknown_sections)}. "
            f"Các section hợp lệ: {sorted(known_sections)}"
        )

    _logger.debug("Đã đọc thành công %d tham số cấu hình từ TOML.", len(flattened))
    return flattened


def _map_if_present(
    source: dict[str, Any],
    source_key: str,
    target: dict[str, Any],
    target_key: str,
) -> None:
    """
    Sao chép giá trị từ source dict sang target dict nếu key tồn tại.
    Hàm tiện ích nội bộ — không dùng trực tiếp từ bên ngoài module.
    """
    if source_key in source:
        target[target_key] = source[source_key]


def _resolve_audit_log_path(log_dir_raw: Any, config_dir: Path) -> Path:
    """
    Chuyển log_dir từ config.toml thành đường dẫn tuyệt đối đến file audit.jsonl.

    Nếu log_dir là đường dẫn tương đối, resolve so với thư mục chứa config.toml.
    Đảm bảo file audit.jsonl nằm bên trong thư mục dự án (chống path traversal).

    Args:
        log_dir_raw: Giá trị log_dir từ TOML (phải là string)
        config_dir: Thư mục chứa file config.toml

    Returns:
        Path tuyệt đối đến file audit.jsonl

    Raises:
        ConfigError: Nếu giá trị không hợp lệ hoặc cố vượt ra ngoài thư mục dự án
    """
    if not isinstance(log_dir_raw, str):
        raise ConfigError(
            f"Giá trị log_dir trong [audit] phải là string, nhận được: {type(log_dir_raw).__name__}"
        )

    log_dir_str = log_dir_raw.strip()

    # Kiểm tra không chứa null bytes — dấu hiệu tấn công
    if "\x00" in log_dir_str:
        raise ConfigError("Giá trị log_dir chứa null byte — không hợp lệ.")

    log_dir_path = Path(log_dir_str)

    # Chuyển thành đường dẫn tuyệt đối nếu là tương đối
    if not log_dir_path.is_absolute():
        log_dir_path = (config_dir / log_dir_path).resolve()
    else:
        log_dir_path = log_dir_path.resolve()

    # Kiểm tra path traversal — log_dir phải nằm trong thư mục dự án
    # Cho phép đường dẫn tuyệt đối tùy ý nếu người dùng chủ động chỉ định
    # (chỉ chặn khi là relative path cố vượt ra ngoài)
    if not log_dir_raw.startswith(("/", "\\")) and ":\\" not in log_dir_raw:
        # Là relative path — kiểm tra không vượt ra ngoài thư mục dự án
        try:
            log_dir_path.relative_to(config_dir.resolve())
        except ValueError:
            raise ConfigError(
                f"Giá trị log_dir '{log_dir_raw}' cố vượt ra ngoài thư mục dự án. "
                "Dùng đường dẫn tuyệt đối nếu muốn lưu log ở nơi khác."
            )

    return log_dir_path / "audit.jsonl"


# ─────────────────────────────────────────────────────────────────────────────
# Hàm validate Config sau khi load
# ─────────────────────────────────────────────────────────────────────────────

def validate(cfg: Config) -> None:
    """
    Kiểm tra toàn bộ ràng buộc bảo mật và logic nghiệp vụ của Config.

    Hàm này phải được gọi SAU khi build Config và TRƯỚC khi server bắt đầu nhận request.
    Nếu bất kỳ ràng buộc nào vi phạm, raise ConfigValidationError ngay lập tức.

    Args:
        cfg: Đối tượng Config cần kiểm tra

    Raises:
        ConfigValidationError: Nếu bất kỳ ràng buộc nào vi phạm
    """
    errors: list[str] = []  # Gom tất cả lỗi, báo một lần để dễ sửa

    # ── Kiểm tra ACCOUNT_NAME ──
    if not cfg.ACCOUNT_NAME or not cfg.ACCOUNT_NAME.strip():
        errors.append(
            "ACCOUNT_NAME (account_name trong [outlook]) không được để trống. "
            "Điền tên tài khoản email Outlook của bạn."
        )

    # ── Kiểm tra ALLOWED_FOLDERS ──
    if not isinstance(cfg.ALLOWED_FOLDERS, list) or len(cfg.ALLOWED_FOLDERS) == 0:
        errors.append(
            "ALLOWED_FOLDERS (allowed_folders trong [security]) không được rỗng. "
            "Chỉ định ít nhất một thư mục Claude được phép truy cập, ví dụ: ['Inbox']"
        )
    else:
        # Kiểm tra từng entry trong allowed_folders
        for i, folder in enumerate(cfg.ALLOWED_FOLDERS):
            folder_errors = _validate_folder_entry(folder, index=i)
            errors.extend(folder_errors)

    # ── Kiểm tra MAX_EMAILS_PER_REQUEST ──
    if not isinstance(cfg.MAX_EMAILS_PER_REQUEST, int) or cfg.MAX_EMAILS_PER_REQUEST <= 0:
        errors.append(
            f"MAX_EMAILS_PER_REQUEST phải là số nguyên dương, nhận được: {cfg.MAX_EMAILS_PER_REQUEST}"
        )
    elif cfg.MAX_EMAILS_PER_REQUEST > _HARD_CAP_MAX_EMAILS:
        # Clamp xuống hard cap — ghi warning nhưng không raise error
        _logger.warning(
            "MAX_EMAILS_PER_REQUEST=%d vượt hard cap %d. Tự động giảm xuống %d.",
            cfg.MAX_EMAILS_PER_REQUEST,
            _HARD_CAP_MAX_EMAILS,
            _HARD_CAP_MAX_EMAILS,
        )
        cfg.MAX_EMAILS_PER_REQUEST = _HARD_CAP_MAX_EMAILS

    # ── Kiểm tra MAX_BODY_CHARS ──
    if not isinstance(cfg.MAX_BODY_CHARS, int) or cfg.MAX_BODY_CHARS <= 0:
        errors.append(
            f"MAX_BODY_CHARS phải là số nguyên dương, nhận được: {cfg.MAX_BODY_CHARS}"
        )
    elif cfg.MAX_BODY_CHARS > _HARD_CAP_MAX_BODY_CHARS:
        _logger.warning(
            "MAX_BODY_CHARS=%d vượt hard cap %d. Tự động giảm xuống %d.",
            cfg.MAX_BODY_CHARS,
            _HARD_CAP_MAX_BODY_CHARS,
            _HARD_CAP_MAX_BODY_CHARS,
        )
        cfg.MAX_BODY_CHARS = _HARD_CAP_MAX_BODY_CHARS

    # ── Kiểm tra AUDIT_LOG_PATH ──
    if not isinstance(cfg.AUDIT_LOG_PATH, Path):
        errors.append(
            f"AUDIT_LOG_PATH phải là Path object, nhận được: {type(cfg.AUDIT_LOG_PATH).__name__}"
        )
    elif cfg.AUDIT_LOG_PATH.suffix.lower() not in {".jsonl", ".json", ".log"}:
        errors.append(
            f"AUDIT_LOG_PATH phải có đuôi .jsonl, .json hoặc .log — "
            f"nhận được: '{cfg.AUDIT_LOG_PATH.suffix}'"
        )

    # ── Kiểm tra READ_ONLY_MODE kết hợp ALLOWED_FOLDERS ──
    # SEC-07: allowed_folders rỗng + read_only=false bị reject
    if not cfg.READ_ONLY_MODE and (
        not cfg.ALLOWED_FOLDERS or len(cfg.ALLOWED_FOLDERS) == 0
    ):
        errors.append(
            "Không thể bật chế độ soạn email (read_only_mode = false) khi allowed_folders rỗng. "
            "Chỉ định ít nhất một thư mục được phép trước."
        )

    # ── Kiểm tra KEYRING_SERVICE ──
    if not cfg.KEYRING_SERVICE or not cfg.KEYRING_SERVICE.strip():
        errors.append(
            "KEYRING_SERVICE không được để trống. "
            "Đây là tên dịch vụ trong Windows Credential Manager."
        )

    # ── Kiểm tra MAX_RECIPIENTS_PER_DRAFT ──
    if not isinstance(cfg.MAX_RECIPIENTS_PER_DRAFT, int) or cfg.MAX_RECIPIENTS_PER_DRAFT <= 0:
        errors.append(
            f"MAX_RECIPIENTS_PER_DRAFT phải là số nguyên dương, nhận được: {cfg.MAX_RECIPIENTS_PER_DRAFT}"
        )
    elif cfg.MAX_RECIPIENTS_PER_DRAFT > _HARD_CAP_MAX_RECIPIENTS:
        _logger.warning(
            "MAX_RECIPIENTS_PER_DRAFT=%d vượt hard cap %d. Tự động giảm xuống %d.",
            cfg.MAX_RECIPIENTS_PER_DRAFT,
            _HARD_CAP_MAX_RECIPIENTS,
            _HARD_CAP_MAX_RECIPIENTS,
        )
        cfg.MAX_RECIPIENTS_PER_DRAFT = _HARD_CAP_MAX_RECIPIENTS

    # ── Kiểm tra MAX_CALLS_PER_MINUTE ──
    if not isinstance(cfg.MAX_CALLS_PER_MINUTE, int) or cfg.MAX_CALLS_PER_MINUTE <= 0:
        errors.append(
            f"MAX_CALLS_PER_MINUTE phải là số nguyên dương, nhận được: {cfg.MAX_CALLS_PER_MINUTE}"
        )
    elif cfg.MAX_CALLS_PER_MINUTE > _HARD_CAP_MAX_CALLS_PER_MINUTE:
        _logger.warning(
            "MAX_CALLS_PER_MINUTE=%d vượt hard cap %d. Tự động giảm xuống %d.",
            cfg.MAX_CALLS_PER_MINUTE,
            _HARD_CAP_MAX_CALLS_PER_MINUTE,
            _HARD_CAP_MAX_CALLS_PER_MINUTE,
        )
        cfg.MAX_CALLS_PER_MINUTE = _HARD_CAP_MAX_CALLS_PER_MINUTE

    # ── Kiểm tra AUDIT_RETAIN_DAYS ──
    if not isinstance(cfg.AUDIT_RETAIN_DAYS, int) or cfg.AUDIT_RETAIN_DAYS < 1:
        errors.append(
            f"AUDIT_RETAIN_DAYS phải là số nguyên >= 1, nhận được: {cfg.AUDIT_RETAIN_DAYS}"
        )

    # ── Kiểm tra AUDIT_HASH_ALGORITHM ──
    valid_hash_algos = {"sha256", "sha512", "sha3_256"}
    if cfg.AUDIT_HASH_ALGORITHM.lower() not in valid_hash_algos:
        errors.append(
            f"AUDIT_HASH_ALGORITHM phải là một trong {sorted(valid_hash_algos)}, "
            f"nhận được: '{cfg.AUDIT_HASH_ALGORITHM}'"
        )

    # ── Kiểm tra SEARCH_QUERY_MAX_LENGTH ──
    if (
        not isinstance(cfg.SEARCH_QUERY_MAX_LENGTH, int)
        or cfg.SEARCH_QUERY_MAX_LENGTH <= 0
        or cfg.SEARCH_QUERY_MAX_LENGTH > _HARD_CAP_SEARCH_QUERY_MAX_LENGTH
    ):
        errors.append(
            f"SEARCH_QUERY_MAX_LENGTH phải là số nguyên từ 1 đến {_HARD_CAP_SEARCH_QUERY_MAX_LENGTH}, "
            f"nhận được: {cfg.SEARCH_QUERY_MAX_LENGTH}"
        )

    # ── Kiểm tra ENTRY_ID_MAX_LENGTH ──
    if (
        not isinstance(cfg.ENTRY_ID_MAX_LENGTH, int)
        or cfg.ENTRY_ID_MAX_LENGTH <= 0
        or cfg.ENTRY_ID_MAX_LENGTH > _HARD_CAP_ENTRY_ID_MAX_LENGTH
    ):
        errors.append(
            f"ENTRY_ID_MAX_LENGTH phải là số nguyên từ 1 đến {_HARD_CAP_ENTRY_ID_MAX_LENGTH}, "
            f"nhận được: {cfg.ENTRY_ID_MAX_LENGTH}"
        )

    # ── Kiểm tra COM_OPERATION_TIMEOUT_SECONDS ──
    if not isinstance(cfg.COM_OPERATION_TIMEOUT_SECONDS, int) or cfg.COM_OPERATION_TIMEOUT_SECONDS < 1:
        errors.append(
            f"COM_OPERATION_TIMEOUT_SECONDS phải là số nguyên >= 1, "
            f"nhận được: {cfg.COM_OPERATION_TIMEOUT_SECONDS}"
        )

    # ── Kiểm tra LOG_LEVEL ──
    valid_log_levels = {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}
    if cfg.LOG_LEVEL.upper() not in valid_log_levels:
        errors.append(
            f"LOG_LEVEL phải là một trong {sorted(valid_log_levels)}, "
            f"nhận được: '{cfg.LOG_LEVEL}'"
        )

    # ── Tổng hợp và raise lỗi nếu có ──
    if errors:
        error_msg = (
            f"Tìm thấy {len(errors)} lỗi cấu hình:\n"
            + "\n".join(f"  [{i+1}] {err}" for i, err in enumerate(errors))
        )
        raise ConfigValidationError(error_msg)

    _logger.info(
        "Cấu hình hợp lệ. read_only=%s, allowed_folders=%d thư mục, account=%s",
        cfg.READ_ONLY_MODE,
        len(cfg.ALLOWED_FOLDERS),
        cfg.ACCOUNT_NAME,
    )


def _validate_folder_entry(folder: Any, index: int) -> list[str]:
    """
    Kiểm tra một entry trong danh sách allowed_folders.

    Quy tắc kiểm tra:
    - Phải là string
    - Không rỗng sau khi strip
    - Không chứa null bytes hoặc control characters
    - Không chứa path traversal (../, ..\\ , ://, :\\\\)
    - Không dùng wildcard đệ quy (**)
    - Độ dài hợp lý (tối đa 260 ký tự — Windows MAX_PATH)

    Args:
        folder: Giá trị entry cần kiểm tra
        index: Vị trí trong danh sách (dùng để báo lỗi rõ hơn)

    Returns:
        Danh sách các thông điệp lỗi (rỗng nếu hợp lệ)
    """
    errs: list[str] = []
    prefix = f"allowed_folders[{index}]"

    # Bước 1: Kiểm tra kiểu dữ liệu
    if not isinstance(folder, str):
        errs.append(f"{prefix}: phải là string, nhận được {type(folder).__name__}")
        return errs  # Không tiếp tục nếu không phải string

    # Bước 2: Kiểm tra không rỗng sau strip
    stripped = folder.strip()
    if not stripped:
        errs.append(f"{prefix}: không được là chuỗi rỗng hoặc chỉ có khoảng trắng")
        return errs

    # Bước 3: Kiểm tra độ dài
    if len(stripped) > 260:
        errs.append(
            f"{prefix}: quá dài ({len(stripped)} ký tự), tối đa 260 ký tự "
            f"(Windows MAX_PATH)"
        )

    # Bước 4: Kiểm tra null bytes và control characters
    for char_idx, ch in enumerate(stripped):
        char_code = ord(ch)
        if char_code == 0:
            errs.append(f"{prefix}: chứa null byte tại vị trí {char_idx}")
            break
        if 0x01 <= char_code <= 0x1F:  # Control characters ASCII
            errs.append(
                f"{prefix}: chứa control character (U+{char_code:04X}) tại vị trí {char_idx}"
            )
            break

    # Bước 5: Kiểm tra path traversal
    path_traversal_patterns = ["../", "..\\", "://", ":\\\\"]
    for pattern in path_traversal_patterns:
        if pattern in stripped:
            errs.append(
                f"{prefix}: chứa pattern path traversal '{pattern}' — không được phép"
            )

    # Bước 6: Kiểm tra wildcard đệ quy ** (không hỗ trợ)
    if "**" in stripped:
        errs.append(
            f"{prefix}: chứa wildcard đệ quy '**' — không hỗ trợ. "
            "Dùng '*' cho wildcard 1 cấp duy nhất."
        )

    # Bước 7: Kiểm tra không phải đường dẫn tuyệt đối
    # (Outlook folder names không được là absolute path)
    normalized = unicodedata.normalize("NFC", stripped)
    if normalized.startswith("/") or (len(normalized) >= 2 and normalized[1] == ":"):
        errs.append(
            f"{prefix}: không được là đường dẫn tuyệt đối ('{stripped}'). "
            "Chỉ dùng tên thư mục Outlook (ví dụ: 'Inbox', 'Inbox/Projects')."
        )

    return errs


# ─────────────────────────────────────────────────────────────────────────────
# Hàm tạo thư mục audit log
# ─────────────────────────────────────────────────────────────────────────────

def ensure_audit_log_directory(cfg: Config) -> None:
    """
    Tạo thư mục chứa file audit log nếu chưa tồn tại.

    Hàm này phải được gọi trước khi AuditLogger bắt đầu ghi log.
    Nếu không thể tạo thư mục (ví dụ: thiếu quyền), raise lỗi rõ ràng.

    Args:
        cfg: Config đã được validate

    Raises:
        ConfigError: Nếu không thể tạo thư mục do thiếu quyền hoặc lỗi hệ thống
    """
    audit_dir = cfg.AUDIT_LOG_PATH.parent

    # Bước 1: Kiểm tra thư mục đã tồn tại chưa
    if audit_dir.exists():
        if not audit_dir.is_dir():
            raise ConfigError(
                f"Đường dẫn audit log '{audit_dir}' đã tồn tại nhưng không phải thư mục. "
                "Xóa hoặc đổi tên file đó trước."
            )
        _logger.debug("Thư mục audit log đã tồn tại: '%s'", audit_dir)
        return

    # Bước 2: Tạo thư mục (bao gồm tất cả thư mục cha chưa tồn tại)
    try:
        audit_dir.mkdir(parents=True, exist_ok=True)
        _logger.info("Đã tạo thư mục audit log: '%s'", audit_dir)
    except OSError as exc:
        raise ConfigError(
            f"Không thể tạo thư mục audit log tại '{audit_dir}': {exc}\n"
            "Kiểm tra quyền ghi hoặc thay đổi đường dẫn trong config.toml."
        ) from exc


# ─────────────────────────────────────────────────────────────────────────────
# Hàm tiện ích public: load_config() — entry point chính
# ─────────────────────────────────────────────────────────────────────────────

def load_config(config_toml_path: Path | None = None) -> Config:
    """
    Hàm chính để tải cấu hình — đây là entry point duy nhất nên dùng.

    Quy trình:
    1. Bắt đầu từ Config mặc định (tất cả giá trị mặc định an toàn)
    2. Nếu config_toml_path được chỉ định và tồn tại, đọc và ghi đè lên mặc định
    3. Nếu không chỉ định path, tự động tìm config.toml trong thư mục hiện tại
    4. Validate toàn bộ cấu hình — raise nếu có lỗi
    5. Tạo thư mục audit log nếu chưa có
    6. Trả về Config đã được kiểm tra và sẵn sàng dùng

    Args:
        config_toml_path: Đường dẫn đến file config.toml.
                          Nếu None, tự tìm trong thư mục hiện tại.

    Returns:
        Config object đã được validate và sẵn sàng dùng

    Raises:
        ConfigError: Lỗi đọc file hoặc cú pháp TOML
        ConfigValidationError: Ràng buộc bảo mật bị vi phạm
    """
    # Bước 1: Xác định đường dẫn file config
    if config_toml_path is None:
        # Tự động tìm config.toml trong thư mục chứa config.py hiện tại
        config_toml_path = Path(__file__).parent / "config.toml"
        _logger.debug(
            "Không chỉ định đường dẫn config — tự động tìm tại '%s'.", config_toml_path
        )

    # Bước 2: Đọc TOML và phẳng hoá thành dict
    overrides = load_from_toml(config_toml_path)

    # Bước 3: Xây dựng Config với các giá trị ghi đè
    cfg = Config()

    for field_name, value in overrides.items():
        # Kiểm tra field tồn tại trong Config trước khi ghi đè
        if not hasattr(cfg, field_name):
            _logger.warning(
                "Tham số '%s' không tồn tại trong Config — bỏ qua.", field_name
            )
            continue

        # Ghi đè giá trị
        setattr(cfg, field_name, value)
        _logger.debug("Ghi đè tham số '%s' từ config.toml.", field_name)

    # Bước 4: Validate cấu hình đã gộp
    validate(cfg)

    # Bước 5: Tạo thư mục audit log
    ensure_audit_log_directory(cfg)

    _logger.info(
        "Config đã load thành công. "
        "Phiên bản server: %s | read_only=%s | %d thư mục được phép",
        cfg.SERVER_VERSION,
        cfg.READ_ONLY_MODE,
        len(cfg.ALLOWED_FOLDERS),
    )
    return cfg


# ─────────────────────────────────────────────────────────────────────────────
# Cho phép chạy trực tiếp để kiểm tra cấu hình: python config.py
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    logging.basicConfig(
        level=logging.DEBUG,
        format="%(levelname)-8s %(name)s: %(message)s",
    )

    print("=== Kiểm tra cấu hình Claude-Outlook MCP Secure ===\n")

    # Lấy đường dẫn config.toml từ tham số dòng lệnh nếu có
    toml_path: Path | None = Path(sys.argv[1]) if len(sys.argv) > 1 else None

    try:
        config = load_config(toml_path)
        print("Cấu hình hợp lệ! Tóm tắt:")
        print(f"  Tài khoản       : {config.ACCOUNT_NAME}")
        print(f"  Chế độ          : {'Chỉ đọc (read-only)' if config.READ_ONLY_MODE else 'Đọc và soạn thảo'}")
        print(f"  Thư mục cho phép: {config.ALLOWED_FOLDERS}")
        print(f"  Email tối đa/lần: {config.MAX_EMAILS_PER_REQUEST}")
        print(f"  Nội dung tối đa : {config.MAX_BODY_CHARS} ký tự")
        print(f"  Audit log       : {config.AUDIT_LOG_PATH}")
        print(f"  Rate limit      : {config.MAX_CALLS_PER_MINUTE} lần/phút")
        print(f"  Keyring service : {config.KEYRING_SERVICE}")
        print(f"\nPhiên bản server: {config.SERVER_VERSION}")
    except (ConfigError, ConfigValidationError) as exc:
        print(f"\nLỖI CẤU HÌNH:\n{exc}", file=sys.stderr)
        sys.exit(1)
