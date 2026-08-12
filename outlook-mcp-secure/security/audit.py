"""
Audit Logger — Ghi nhật ký kiểm toán cho mọi thao tác của MCP server.

Mục đích: Đảm bảo truy vết (audit trail) đầy đủ theo yêu cầu bảo mật cấp ngân hàng.
Mỗi tool call, sự kiện bảo mật, và trạng thái phiên làm việc đều được ghi vào file
JSON Lines (mỗi dòng là một JSON object độc lập) để dễ phân tích sau này.

Nguyên tắc bảo mật bắt buộc:
  - Tuyệt đối không ghi nội dung email (subject, body, sender, recipient)
  - Tuyệt đối không ghi credentials, API key, password, token
  - Các giá trị nhạy cảm được hash SHA256 hoặc loại bỏ hoàn toàn
  - Truy cập file log bị giới hạn ở mức quyền tối thiểu (0o600 khi có thể)
  - Thread-safe: dùng Lock để tránh ghi đè khi nhiều luồng chạy đồng thời
  - Tự động xoay vòng file (rotate) khi kích thước vượt 10 MB
"""

from __future__ import annotations

import hashlib
import hmac as _hmac_module
import json
import logging
import os
import stat
import subprocess
import sys
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

# Logger nội bộ — chỉ ghi vào stderr/console, không ghi vào audit file
_internal_logger = logging.getLogger(__name__)

# Danh sách từ khóa nhạy cảm trong tên key — các key này sẽ bị loại khỏi params khi ghi log
_SENSITIVE_KEY_FRAGMENTS = (
    "password",
    "credential",
    "key",
    "secret",
    "token",
    "auth",
    "api_key",
    "apikey",
    "passwd",
    "pwd",
)

# Giới hạn độ dài chuỗi trong params trước khi ghi (tránh ghi vô tình nội dung email)
_MAX_PARAM_VALUE_LENGTH = 200

# Ngưỡng kích thước file audit trước khi xoay vòng (10 MB)
_ROTATE_SIZE_BYTES = 10 * 1024 * 1024

# Số lượng entries tích lũy trước khi ghi integrity checkpoint
_INTEGRITY_CHECK_INTERVAL = 100

# Số lượng entries tối đa lưu trong buffer để tính hash chain checkpoint
_CHECKPOINT_WINDOW = 100

# Tên environment variable chứa HMAC key cho audit log
_AUDIT_HMAC_KEY_ENV = "OUTLOOK_MCP_AUDIT_KEY"


class AuditLogger:
    """
    Ghi nhật ký kiểm toán (audit trail) cho mọi thao tác của MCP server.

    Mỗi phiên làm việc (session) có một UUID riêng để phân biệt.
    File log theo định dạng JSON Lines, encoding UTF-8.
    Thread-safe nhờ Lock nội bộ.

    Cách dùng:
        audit = AuditLogger(Path("logs/audit.jsonl"))
        audit.log(tool="list_emails", action="call", params={...}, result="ok")
        audit.close()
    """

    def __init__(
        self,
        config: Any = None,
        session_id: str | None = None,
        server_version: str = "",
        log_path: Path | None = None,
    ) -> None:
        """
        Khởi tạo AuditLogger.

        Hỗ trợ hai cách gọi:
          - Cách 1 (từ server.py): AuditLogger(config=_config, session_id=SESSION_ID, server_version=SERVER_VERSION)
            → lấy đường dẫn log từ config.AUDIT_LOG_PATH (hoặc config.audit_log_path)
          - Cách 2 (trực tiếp): AuditLogger(log_path=Path("logs/audit.jsonl"))
            → dùng đường dẫn truyền trực tiếp

        Bước 1: Xác định đường dẫn file log từ config hoặc tham số log_path.
        Bước 2: Tạo thư mục chứa file log nếu chưa có.
        Bước 3: Đặt quyền truy cập thư mục ở mức tối thiểu (0o700).
        Bước 4: Lưu session_id và server_version để ghi vào mỗi entry.
        Bước 5: Mở file log để ghi thêm (append), không xóa log cũ.
        Bước 6: Ghi entry đánh dấu bắt đầu phiên.

        Args:
            config:         Object cấu hình từ config.py (có thuộc tính AUDIT_LOG_PATH hoặc audit_log_path).
            session_id:     UUID phiên làm việc do server.py tạo ra (dùng chung với server).
            server_version: Phiên bản server để ghi vào audit entries.
            log_path:       Đường dẫn trực tiếp đến file audit log (dùng khi không có config).
        """
        # Bước 1: Xác định đường dẫn file log
        resolved_path: Path
        if log_path is not None:
            # Cách gọi trực tiếp — dùng đường dẫn được truyền vào
            resolved_path = log_path
        elif config is not None:
            # Cách gọi từ server.py — lấy đường dẫn từ config
            # Thử các tên thuộc tính phổ biến theo thứ tự ưu tiên
            raw_path = (
                getattr(config, "AUDIT_LOG_PATH", None)
                or getattr(config, "audit_log_path", None)
                or getattr(config, "log_path", None)
            )
            if raw_path is None:
                # Fallback an toàn: ghi vào thư mục logs/ cạnh file config
                resolved_path = Path("logs") / "audit.jsonl"
                _internal_logger.warning(
                    "Không tìm thấy AUDIT_LOG_PATH trong config — dùng mặc định: %s",
                    resolved_path,
                )
            else:
                resolved_path = Path(raw_path)
        else:
            # Không có cả config lẫn log_path — dùng mặc định
            resolved_path = Path("logs") / "audit.jsonl"
            _internal_logger.warning(
                "AuditLogger khởi tạo không có config và log_path — dùng mặc định: %s",
                resolved_path,
            )

        # Bước 2: Đảm bảo thư mục tồn tại
        self._log_path = resolved_path
        log_dir = resolved_path.parent
        log_dir.mkdir(parents=True, exist_ok=True)

        # Bước 3: Giới hạn quyền truy cập thư mục và file về mức tối thiểu
        # F-AUD-07: Dùng icacls trên Windows (os.chmod không có hiệu lực thực sự trên Windows)
        # Trên Linux/macOS vẫn dùng os.chmod như cũ
        # Phân quyền thư mục log theo platform
        if sys.platform == "win32":
            try:
                result = subprocess.run(
                    ["icacls", str(log_dir), "/inheritance:r", "/grant:r",
                     os.environ.get("USERNAME", "Owner") + ":(OI)(CI)F",
                     "SYSTEM:(OI)(CI)F"],
                    capture_output=True, timeout=10
                )
                if result.returncode != 0:
                    _internal_logger.warning(
                        "icacls không thể set ACL cho logs/: %s. Audit log có thể đọc được bởi user khác.",
                        result.stderr.decode(errors="replace"))
            except Exception as e:
                _internal_logger.warning("Không thể set Windows ACL cho logs/: %s.", e)
        else:
            try:
                os.chmod(log_dir, stat.S_IRWXU)  # 0o700: chỉ owner đọc/ghi/thực thi
            except OSError:
                _internal_logger.debug("Không thể set chmod cho thư mục log")

        # Bước 4: Lưu session_id và server_version
        # Nếu session_id được truyền vào từ server.py thì dùng luôn (nhóm log cùng session với server)
        # Nếu không có thì tạo mới
        self._session_id: str = session_id if session_id else str(uuid.uuid4())
        self._server_version: str = server_version

        # Lock để đảm bảo thread-safety khi nhiều luồng ghi đồng thời
        self._lock: threading.Lock = threading.Lock()

        # F-AUD-01: Lấy HMAC key từ environment variable bắt buộc.
        # KHÔNG dùng session_id làm key vì session_id ghi rõ trong mỗi dòng log —
        # ai đọc file log là có thể tính lại toàn bộ HMAC (vô hiệu hóa tính năng chống tamper).
        env_key = os.environ.get(_AUDIT_HMAC_KEY_ENV, "").strip()
        if env_key:
            # Dùng key từ env var: hash thêm một lần để normalize độ dài và tránh weak keys
            self._hmac_key: bytes = hashlib.sha256(env_key.encode("utf-8")).digest()
        else:
            # Fallback: sinh key ngẫu nhiên mỗi phiên — không persist qua restart nhưng
            # đảm bảo key KHÔNG xuất hiện trong log (khác với session_id)
            import secrets as _secrets_mod
            self._hmac_key = _secrets_mod.token_bytes(32)
            _internal_logger.error(
                "OUTLOOK_MCP_AUDIT_KEY chưa được đặt. "
                "Audit log sẽ dùng key ngẫu nhiên mỗi phiên — không thể verify log qua restart. "
                "Đặt OUTLOOK_MCP_AUDIT_KEY để bảo vệ tính toàn vẹn của audit log lâu dài."
            )

        # F-AUD-06: Tạo fingerprint 16 hex chars để nhận diện key mà không lộ key thực
        # Fingerprint giúp phát hiện khi HMAC key bị thay đổi giữa các lần chạy server
        # (key rotation detection) mà không cần lưu key vào log
        self._key_fingerprint: str = hashlib.sha256(self._hmac_key).hexdigest()[:16]

        # Đếm số entries kể từ lần integrity check cuối
        self._entries_since_checkpoint: int = 0

        # Buffer lưu _CHECKPOINT_WINDOW entries gần nhất để tính hash chain checkpoint
        self._recent_entries: list[str] = []

        # F-AUD-04: Hash của checkpoint trước — dùng để xây dựng hash chain liên tục
        # Giá trị rỗng cho checkpoint đầu tiên trong phiên
        self._last_checkpoint_hash: str = ""

        # Bước 5: Mở file ghi thêm (append mode), không xóa dữ liệu cũ
        self._file = open(resolved_path, "a", encoding="utf-8", buffering=1)  # buffering=1: line-buffered

        # Đặt quyền file audit về 0o600: chỉ owner đọc/ghi
        try:
            os.chmod(resolved_path, stat.S_IRUSR | stat.S_IWUSR)  # 0o600
        except OSError:
            _internal_logger.debug("Không thể set chmod cho file log (bình thường trên Windows)")

        # Bước 6: Ghi entry đánh dấu bắt đầu phiên làm việc
        # F-AUD-06: Thêm key_fingerprint để phát hiện khi HMAC key bị thay đổi (key rotation)
        self._write_raw_entry({
            "event": "session_start",
            "session_id": self._session_id,
            "server_version": self._server_version,
            "log_path": str(resolved_path),
            "key_fingerprint": self._key_fingerprint,
        })

        _internal_logger.info("AuditLogger khởi động — session_id=%s", self._session_id)

    # ------------------------------------------------------------------ #
    #  Phương thức ghi log công khai (public)                             #
    # ------------------------------------------------------------------ #

    def log(
        self,
        tool: str,
        action: str,
        params: dict[str, Any],
        result: str,
        details: str = "",
        duration_ms: int | None = None,
        items_returned: int | None = None,
    ) -> None:
        """
        Ghi một lần gọi tool vào audit log.

        Params sẽ được làm sạch trước khi ghi:
          - Xóa các key nhạy cảm (password, key, token, ...)
          - Cắt ngắn chuỗi dài hơn 200 ký tự
          - Không bao giờ ghi nội dung email

        Args:
            tool:           Tên tool MCP (vd: "list_emails", "read_email").
            action:         Hành động cụ thể (vd: "call", "blocked", "error").
            params:         Tham số đầu vào của tool — sẽ được sanitize.
            result:         Kết quả ("ok", "blocked", "error", "SECURITY_EVENT").
            details:        Mô tả bổ sung, không chứa dữ liệu nhạy cảm.
            duration_ms:    Thời gian thực thi tính bằng mili giây (tùy chọn).
            items_returned: Số lượng kết quả trả về (tùy chọn, dùng cho list/search).
        """
        # Làm sạch params trước khi ghi
        safe_params = self._sanitize_params(params)

        entry: dict[str, Any] = {
            "session_id": self._session_id,
            "tool": tool,
            "action": action,
            "params": safe_params,
            "result": result,
        }

        # Thêm các trường tùy chọn chỉ khi có giá trị — tránh làm log thêm null
        if details:
            entry["details"] = details
        if duration_ms is not None:
            entry["duration_ms"] = duration_ms
        if items_returned is not None:
            entry["items_returned"] = items_returned

        self._write_raw_entry(entry)

    def log_security_event(self, event_type: str, details: str) -> None:
        """
        Ghi một sự kiện bảo mật quan trọng vào audit log.

        Dùng khi phát hiện hành vi đáng ngờ hoặc vi phạm chính sách bảo mật,
        ví dụ: cố truy cập folder ngoài allowlist, injection pattern bị phát hiện,
        rate limit bị vượt quá.

        Args:
            event_type: Loại sự kiện (vd: "folder_not_allowed", "injection_detected",
                        "rate_limit_exceeded", "read_only_mode_violation").
            details:    Mô tả chi tiết sự kiện — KHÔNG chứa dữ liệu nhạy cảm.
        """
        entry: dict[str, Any] = {
            "session_id": self._session_id,
            "tool": "SYSTEM",
            "action": "security_event",
            "params": {"event_type": event_type},
            "result": "SECURITY_EVENT",
            "details": details,
        }
        self._write_raw_entry(entry)
        _internal_logger.warning("SECURITY EVENT: %s — %s", event_type, details)

    def log_credential_access(self, credential_type: str, status: str) -> None:
        """
        Ghi lại việc truy cập thông tin xác thực (credentials).

        Chỉ ghi loại credential và trạng thái — tuyệt đối không ghi giá trị.

        Args:
            credential_type: Loại credential (vd: "anthropic_api_key", "outlook_password").
            status:          Trạng thái truy cập ("ok", "not_found", "error").
        """
        entry: dict[str, Any] = {
            "event": "credential_access",
            "session_id": self._session_id,
            "credential_type": credential_type,
            "status": status,
            # Tuyệt đối không ghi credential_value hay bất kỳ giá trị thực nào
        }
        self._write_raw_entry(entry)

    def log_tool_start(self, tool_name: str, raw_arguments: dict[str, Any]) -> None:
        """
        Ghi entry đánh dấu bắt đầu xử lý một tool call.

        Ghi TRƯỚC khi thực thi để đảm bảo không mất log nếu server crash giữa chừng.
        Params được sanitize để loại bỏ dữ liệu nhạy cảm trước khi ghi.

        Args:
            tool_name:     Tên tool đang được gọi (vd: "list_emails").
            raw_arguments: Tham số thô từ Claude — sẽ được sanitize trước khi ghi.
        """
        entry: dict[str, Any] = {
            "session_id": self._session_id,
            "tool": tool_name,
            "action": "call",
            "params": self._sanitize_params(raw_arguments),
            "result": "pending",
        }
        self._write_raw_entry(entry)

    def log_tool_success(
        self,
        tool_name: str,
        duration_ms: int | None = None,
        items_returned: int | None = None,
    ) -> None:
        """
        Ghi entry đánh dấu tool call thực thi thành công.

        Args:
            tool_name:      Tên tool đã thực thi thành công.
            duration_ms:    Thời gian thực thi tính bằng mili giây.
            items_returned: Số lượng kết quả trả về (nếu có, ví dụ số email trong list).
        """
        self.log(
            tool=tool_name,
            action="success",
            params={},
            result="ok",
            duration_ms=duration_ms,
            items_returned=items_returned,
        )

    def log_tool_blocked(
        self,
        tool_name: str,
        block_reason: str,
        duration_ms: int | None = None,
    ) -> None:
        """
        Ghi entry đánh dấu tool call bị chặn do vi phạm bảo mật hoặc validation.

        Args:
            tool_name:    Tên tool bị chặn.
            block_reason: Lý do chặn (vd: "not_in_allowlist", "validation_error", "rate_limit").
            duration_ms:  Thời gian xử lý trước khi bị chặn (mili giây).
        """
        self.log(
            tool=tool_name,
            action="blocked",
            params={"block_reason": block_reason},
            result="blocked",
            details=block_reason,
            duration_ms=duration_ms,
        )

    def log_tool_error(
        self,
        tool_name: str,
        error_code: str,
        duration_ms: int | None = None,
    ) -> None:
        """
        Ghi entry đánh dấu tool call thất bại do lỗi kỹ thuật (không phải bảo mật).

        Chỉ ghi tên loại lỗi (error_code = tên class exception), không ghi message
        gốc vì message có thể chứa dữ liệu nhạy cảm.

        Args:
            tool_name:   Tên tool gặp lỗi.
            error_code:  Tên loại exception (vd: "OutlookOperationError", "ValueError").
            duration_ms: Thời gian thực thi trước khi lỗi (mili giây).
        """
        self.log(
            tool=tool_name,
            action="error",
            params={"error_code": error_code},
            result="error",
            details=error_code,
            duration_ms=duration_ms,
        )

    def log_server_start(
        self,
        version: str | None = None,
        read_only: bool | None = None,
        allowlist_count: int | None = None,
        com_backend: str = "win32com",
    ) -> None:
        """
        Ghi entry đánh dấu server khởi động thành công.

        Có thể gọi không tham số (server.py tự động lấy từ _server_version đã lưu lúc __init__),
        hoặc truyền tường minh để ghi thêm thông tin.

        Args:
            version:         Phiên bản server (vd: "1.0.0"). Mặc định dùng server_version từ __init__.
            read_only:       Chế độ chỉ đọc có bật không. Bỏ qua nếu None.
            allowlist_count: Số lượng folder trong allowlist. Bỏ qua nếu None.
            com_backend:     Thư viện COM đang dùng (mặc định "win32com").
        """
        entry: dict[str, Any] = {
            "event": "server_start",
            "session_id": self._session_id,
            "version": version if version is not None else self._server_version,
            "com_backend": com_backend,
        }
        # Chỉ ghi các trường tùy chọn khi có giá trị — tránh ghi null không cần thiết
        if read_only is not None:
            entry["read_only"] = read_only
        if allowlist_count is not None:
            entry["allowlist_count"] = allowlist_count
        self._write_raw_entry(entry)

    def log_server_stop(self) -> None:
        """
        Ghi entry đánh dấu server tắt có kiểm soát (graceful shutdown).

        Gọi bởi server.py trong _teardown_components() trước khi đóng file log.
        """
        self._write_raw_entry({
            "event": "server_stop",
            "session_id": self._session_id,
            "version": self._server_version,
        })

    def flush(self) -> None:
        """
        Flush toàn bộ dữ liệu đang đệm xuống đĩa ngay lập tức.

        Gọi bởi server.py trong _teardown_components() sau khi ghi log_server_stop()
        để đảm bảo không mất dữ liệu khi server tắt.
        """
        try:
            with self._lock:
                self._file.flush()
                os.fsync(self._file.fileno())
        except OSError as exc:
            _internal_logger.error("Lỗi khi flush audit log: %s", exc)

    def close(self) -> None:
        """
        Đóng audit logger khi server tắt.

        Bước 1: Ghi entry đánh dấu kết thúc phiên làm việc.
        Bước 2: Flush buffer để đảm bảo không mất dữ liệu.
        Bước 3: Đóng file handle.
        """
        try:
            # Bước 1: Ghi entry kết thúc phiên
            self._write_raw_entry({
                "event": "session_end",
                "session_id": self._session_id,
            })

            # Bước 2: Flush toàn bộ buffer xuống đĩa
            with self._lock:
                self._file.flush()
                os.fsync(self._file.fileno())

        except OSError as exc:
            _internal_logger.error("Lỗi khi đóng audit log: %s", exc)
        finally:
            # Bước 3: Đóng file handle dù có lỗi hay không
            try:
                self._file.close()
            except OSError:
                pass
            _internal_logger.info("AuditLogger đã đóng — session_id=%s", self._session_id)

    # ------------------------------------------------------------------ #
    #  Phương thức nội bộ (private)                                       #
    # ------------------------------------------------------------------ #

    def _write_raw_entry(self, entry: dict[str, Any]) -> None:
        """
        Ghi một entry thô vào file log theo định dạng JSON Lines.

        Bước 1: Thêm timestamp UTC ISO8601 vào entry.
        Bước 2: Kiểm tra có cần xoay vòng file (rotate) không.
        Bước 3: Ghi JSON line + newline, flush ngay.
        Bước 4: Cập nhật bộ đếm để kiểm tra integrity checkpoint.

        Args:
            entry: Dict chứa nội dung cần ghi — phải đã được sanitize.
        """
        # Bước 1: Thêm timestamp UTC ISO8601 có timezone (+00:00)
        entry["ts"] = datetime.now(tz=timezone.utc).isoformat()

        # Thêm HMAC signature để phát hiện log tampering
        entry["hmac"] = self._sign_entry(entry)

        with self._lock:
            try:
                # Bước 2: Kiểm tra kích thước file, xoay vòng nếu cần
                self._rotate_if_needed()

                # Bước 3: Serialize và ghi xuống file, mỗi entry trên một dòng
                line = json.dumps(entry, ensure_ascii=False, separators=(",", ":"))
                self._file.write(line + "\n")
                self._file.flush()

                # Bước 4: Cập nhật buffer integrity và bộ đếm
                # F-AUD-04: Tăng window lưu entries lên _CHECKPOINT_WINDOW (100) thay vì 10
                self._recent_entries.append(line)
                if len(self._recent_entries) > _CHECKPOINT_WINDOW:
                    # Giữ tối đa _CHECKPOINT_WINDOW entries gần nhất để tính hash chain checkpoint
                    self._recent_entries.pop(0)

                self._entries_since_checkpoint += 1
                if self._entries_since_checkpoint >= _INTEGRITY_CHECK_INTERVAL:
                    # Ghi integrity checkpoint mà không dùng đệ quy (gọi thẳng file.write)
                    self._write_integrity_checkpoint_unlocked()

            except OSError as exc:
                # Không raise ra ngoài — lỗi ghi log không được làm server crash
                _internal_logger.error("Không thể ghi audit entry: %s", exc)

    def _write_integrity_checkpoint_unlocked(self) -> None:
        """
        Ghi integrity checkpoint sau mỗi _INTEGRITY_CHECK_INTERVAL entries.

        F-AUD-04: Checkpoint dùng hash chain để bao phủ TẤT CẢ entries từ checkpoint
        trước, không chỉ 10 entries cuối. Hash chain = SHA256(entries hiện tại + hash_checkpoint_trước)
        giúp phát hiện mọi sửa đổi dù ở vị trí nào trong chuỗi log.

        Phương thức này phải được gọi trong phạm vi lock đã giữ (_lock).
        Không gọi _write_raw_entry() ở đây để tránh deadlock.
        """
        # Tính hash chain để bao phủ tất cả entries từ checkpoint trước
        # Nối toàn bộ entries trong buffer + hash của checkpoint trước để tạo chain liên tục
        checkpoint_data = "\n".join(self._recent_entries) + self._last_checkpoint_hash
        checkpoint_hash = hashlib.sha256(checkpoint_data.encode("utf-8")).hexdigest()

        # Lưu lại hash này để dùng cho checkpoint kế tiếp (hash chain)
        self._last_checkpoint_hash = checkpoint_hash

        checkpoint = {
            "ts": datetime.now(tz=timezone.utc).isoformat(),
            "event": "integrity_check",
            "session_id": self._session_id,
            "entries_since_last": self._entries_since_checkpoint,
            "checkpoint_hash": f"sha256:{checkpoint_hash}",
            "key_fingerprint": self._key_fingerprint,
        }

        try:
            line = json.dumps(checkpoint, ensure_ascii=False, separators=(",", ":"))
            self._file.write(line + "\n")
            self._file.flush()
        except OSError as exc:
            _internal_logger.error("Không thể ghi integrity checkpoint: %s", exc)

        # Reset bộ đếm sau khi ghi checkpoint
        self._entries_since_checkpoint = 0

    def _sanitize_params(self, params: dict[str, Any]) -> dict[str, Any]:
        """
        Làm sạch params trước khi ghi vào audit log.

        Áp dụng 3 lớp lọc:
          Lớp 1 — Xóa key nhạy cảm: loại bỏ bất kỳ key nào có tên chứa từ như
                   'password', 'key', 'token', 'secret', 'credential', ...
          Lớp 2 — Cắt ngắn chuỗi dài: giới hạn 200 ký tự để tránh vô tình
                   ghi nội dung email (subject, snippet, body).
          Lớp 3 — Đệ quy: áp dụng cho cả dict lồng nhau.

        Args:
            params: Dict params gốc từ tool call.

        Returns:
            Dict đã được làm sạch, an toàn để ghi vào log.
        """
        if not isinstance(params, dict):
            # Nếu params không phải dict (hiếm gặp), trả về dict rỗng cho an toàn
            return {}

        safe: dict[str, Any] = {}

        for key, value in params.items():
            key_lower = str(key).lower()

            # Lớp 1: Bỏ qua key có tên chứa từ khóa nhạy cảm
            if any(fragment in key_lower for fragment in _SENSITIVE_KEY_FRAGMENTS):
                safe[key] = "[REDACTED]"
                continue

            # Lớp 2: Xử lý theo kiểu dữ liệu của value
            if isinstance(value, str):
                if len(value) > _MAX_PARAM_VALUE_LENGTH:
                    # Cắt ngắn và đánh dấu để dễ nhận ra trong log
                    safe[key] = value[:_MAX_PARAM_VALUE_LENGTH] + "...[truncated]"
                else:
                    safe[key] = value

            elif isinstance(value, dict):
                # Lớp 3: Đệ quy cho dict lồng nhau
                safe[key] = self._sanitize_params(value)

            elif isinstance(value, list):
                # Xử lý list: áp dụng sanitize cho từng phần tử
                safe[key] = self._sanitize_list(value)

            elif isinstance(value, (int, float, bool)) or value is None:
                # Kiểu nguyên thủy an toàn — giữ nguyên
                safe[key] = value

            else:
                # Kiểu không xác định — chuyển sang string và cắt ngắn
                str_value = str(value)
                if len(str_value) > _MAX_PARAM_VALUE_LENGTH:
                    safe[key] = str_value[:_MAX_PARAM_VALUE_LENGTH] + "...[truncated]"
                else:
                    safe[key] = str_value

        return safe

    def _sanitize_list(self, items: list[Any]) -> list[Any]:
        """
        Làm sạch danh sách params — hỗ trợ sanitize cho list trong params.

        Áp dụng cùng quy tắc truncate và redact như _sanitize_params
        nhưng cho từng phần tử trong list.

        Args:
            items: Danh sách cần làm sạch.

        Returns:
            Danh sách đã được làm sạch.
        """
        safe_items: list[Any] = []

        for item in items:
            if isinstance(item, str):
                if len(item) > _MAX_PARAM_VALUE_LENGTH:
                    safe_items.append(item[:_MAX_PARAM_VALUE_LENGTH] + "...[truncated]")
                else:
                    safe_items.append(item)
            elif isinstance(item, dict):
                safe_items.append(self._sanitize_params(item))
            elif isinstance(item, list):
                safe_items.append(self._sanitize_list(item))
            elif isinstance(item, (int, float, bool)) or item is None:
                safe_items.append(item)
            else:
                str_item = str(item)
                safe_items.append(
                    str_item[:_MAX_PARAM_VALUE_LENGTH] + "...[truncated]"
                    if len(str_item) > _MAX_PARAM_VALUE_LENGTH
                    else str_item
                )

        return safe_items

    def _sign_entry(self, entry: dict) -> str:
        """
        Tính HMAC-SHA256 cho toàn bộ audit entry để phát hiện chỉnh sửa log.

        F-AUD-03: Phải sign TẤT CẢ các trường, không phải chỉ 4 trường cốt lõi.
        Loại trừ trường "hmac" để tránh circular dependency (HMAC của chính nó).
        """
        # Tạo bản sao entry không có trường "hmac" để tránh circular dependency
        # BREAKING CHANGE v2.1.0: entries signed before v2.1.0 used 4-field HMAC
        # Dùng --legacy flag của verify-integrity.ps1 cho entries cũ
        fields_to_sign = {k: v for k, v in entry.items() if k != "hmac"}
        # Sort keys để đảm bảo canonical form nhất quán
        canonical = json.dumps(fields_to_sign, sort_keys=True, separators=(",", ":"),
                               ensure_ascii=False, default=str)
        return _hmac_module.new(
            self._hmac_key,
            canonical.encode("utf-8"),
            digestmod="sha256"
        ).hexdigest()

    def _rotate_if_needed(self) -> None:
        """
        Xoay vòng file log (log rotation) khi kích thước vượt 10 MB.

        Bước 1: Kiểm tra kích thước file hiện tại.
        Bước 2: Nếu vượt ngưỡng, đóng file hiện tại.
        Bước 3: Đổi tên sang audit-YYYYMMDD-HHMMSS.jsonl.
        Bước 4: Mở file mới với tên gốc để tiếp tục ghi.

        Phương thức này phải được gọi trong phạm vi lock đã giữ (_lock).

        Raises:
            OSError: Nếu không thể rename hoặc tạo file mới (lỗi nghiêm trọng).
        """
        try:
            current_size = self._log_path.stat().st_size
        except OSError:
            # File chưa tồn tại hoặc không đọc được — bỏ qua rotate
            return

        if current_size < _ROTATE_SIZE_BYTES:
            return

        # Bước 2: File đủ lớn để xoay vòng — đóng file hiện tại trước
        try:
            self._file.flush()
            self._file.close()
        except OSError as exc:
            _internal_logger.error("Lỗi khi flush/close trước khi rotate: %s", exc)

        # Bước 3: Tạo tên file archive với timestamp hiện tại
        now_str = datetime.now(tz=timezone.utc).strftime("%Y%m%d-%H%M%S")
        archive_name = self._log_path.parent / f"audit-{now_str}.jsonl"

        try:
            self._log_path.rename(archive_name)
            _internal_logger.info(
                "Đã xoay vòng audit log: %s -> %s (size=%d bytes)",
                self._log_path,
                archive_name,
                current_size,
            )
        except OSError as exc:
            _internal_logger.error("Không thể rename file log khi rotate: %s", exc)
            # Thử mở lại file cũ để không mất kết nối ghi log
            self._file = open(self._log_path, "a", encoding="utf-8", buffering=1)
            return

        # Bước 4: Mở file mới để tiếp tục ghi sau khi rotate
        self._file = open(self._log_path, "a", encoding="utf-8", buffering=1)

        # Đặt lại quyền truy cập cho file mới
        try:
            os.chmod(self._log_path, stat.S_IRUSR | stat.S_IWUSR)  # 0o600
        except OSError:
            _internal_logger.debug("Không thể set chmod sau rotate (bình thường trên Windows)")

        # Ghi entry đánh dấu file mới bắt đầu sau rotate (không dùng _write_raw_entry để tránh deadlock)
        rotate_marker = {
            "ts": datetime.now(tz=timezone.utc).isoformat(),
            "event": "log_rotated",
            "session_id": self._session_id,
            "archived_to": str(archive_name),
            "archived_size_bytes": current_size,
        }
        try:
            line = json.dumps(rotate_marker, ensure_ascii=False, separators=(",", ":"))
            self._file.write(line + "\n")
            self._file.flush()
        except OSError as exc:
            _internal_logger.error("Không thể ghi rotate marker: %s", exc)


# ------------------------------------------------------------------ #
#  Hàm tiện ích — Hash nhạy cảm trước khi ghi                        #
# ------------------------------------------------------------------ #

def hash_sensitive_value(value: str, algorithm: str = "sha256") -> str:
    """
    Hash một giá trị nhạy cảm để ghi vào audit log thay vì giá trị thực.

    Dùng khi cần truy vết "ai dùng query gì" mà không tiết lộ nội dung,
    ví dụ: hash subject, hash search query, hash to_list trong compose_draft.

    Args:
        value:     Giá trị cần hash (vd: search query, email subject).
        algorithm: Thuật toán hash (mặc định "sha256").

    Returns:
        Chuỗi dạng "sha256:<hex_digest>" — dễ nhận biết là hash trong log.

    Raises:
        ValueError: Nếu algorithm không được hỗ trợ.
    """
    if algorithm not in hashlib.algorithms_available:
        raise ValueError(f"Thuật toán hash không hỗ trợ: {algorithm}")

    digest = hashlib.new(algorithm, value.encode("utf-8")).hexdigest()
    return f"{algorithm}:{digest}"
