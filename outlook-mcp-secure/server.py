"""
server.py — Điểm vào chính của MCP server Claude-Outlook Secure

Chức năng tổng thể:
  - Khởi động MCP server theo giao thức stdio (Standard I/O) để Claude Code CLI kết nối.
  - Kiểm tra Outlook Desktop đang chạy trước khi chấp nhận kết nối.
  - Đăng ký toàn bộ 19 công cụ (tools): list_folders, list_emails, read_email,
    search_emails, compose_draft, reply_draft, forward_draft, mark_email_read,
    flag_email, move_email, list_all_folders, email_stats, get_email_thread,
    get_contact_stats, bulk_mark_read, get_flagged_emails, get_project_snapshot,
    list_calendar_events, create_calendar_event.
  - Đảm bảo audit log ghi đầy đủ mỗi lần tool được gọi.
  - Tắt server gọn gàng (graceful shutdown): giải phóng COM objects, flush audit log.
  - Hỗ trợ tham số --setup để chạy wizard cài đặt lần đầu.

Giải thuật tổng thể:
  - MCP server chạy trên asyncio event loop, nhận/gửi JSON qua stdin/stdout.
  - Mọi thao tác Outlook COM được đẩy vào 1 thread riêng (STA thread) để tránh lỗi
    threading của COM trên Windows.
  - Input validation → COM operation → audit log → trả kết quả JSON về Claude.
"""

from __future__ import annotations

import argparse
import asyncio
import gc
import json
import logging
import signal
import sys
import time
import traceback
import unicodedata
import uuid
from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager
from typing import Any

# Thư viện MCP — giao thức kết nối Claude với server
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    CallToolResult,
    TextContent,
    Tool,
    ToolAnnotations,
)

# Bảo đảm stdout/stderr dùng UTF-8 — cần thiết trên Windows để tiếng Việt không bị vỡ
# khi MCP server ghi log qua stdio transport
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

# --- Cấu hình logging nội bộ (không phải audit log) ---
# Log này chỉ dành cho developer debug, không chứa dữ liệu nhạy cảm
_internal_logger = logging.getLogger("outlook_mcp.server")

# Phiên bản server — dùng trong audit log và tool description
SERVER_VERSION = "2.1.0"
SERVER_NAME = "outlook-mcp-secure"

# Session ID duy nhất cho mỗi lần chạy server — dùng để nhóm audit logs theo phiên
SESSION_ID: str = str(uuid.uuid4())

# Thread pool 1 worker duy nhất — bắt buộc vì COM (Windows Component Object Model)
# yêu cầu tất cả các cuộc gọi đến một COM object phải từ cùng một thread (STA - Single
# Threaded Apartment). Nếu gọi từ nhiều thread khác nhau sẽ gây lỗi không xác định.
_com_executor: ThreadPoolExecutor | None = None

# Các biến toàn cục được khởi tạo trong hàm _initialize_components()
_config: Any = None        # Config object từ config.py
_audit: Any = None         # AuditLogger từ security/audit.py
_com_bridge: Any = None    # OutlookCOMBridge từ outlook_com.py


class _TokenBucketRateLimiter:
    """
    Token bucket rate limiter — giới hạn số lượng yêu cầu tối đa mỗi phút.

    Nguyên lý hoạt động (token bucket):
      - Thùng chứa tối đa max_per_minute token.
      - Mỗi giây, thùng được nạp thêm (max_per_minute / 60) token tự động.
      - Mỗi yêu cầu đến tiêu thụ 1 token. Nếu thùng rỗng, yêu cầu bị từ chối.
      - Cho phép burst ngắn hạn (nhiều yêu cầu liên tiếp khi thùng đầy),
        đồng thời duy trì tốc độ trung bình không vượt quá giới hạn.

    Thread-safe: dùng threading.Lock để bảo vệ trạng thái nội bộ khi nhiều
    coroutine (luồng bất đồng bộ) gọi đồng thời.
    """

    def __init__(self, max_per_minute: int) -> None:
        """
        Khởi tạo rate limiter với giới hạn yêu cầu mỗi phút.

        Tham số:
          max_per_minute: Số yêu cầu tối đa được phép trong 1 phút (tối thiểu 1).
        """
        # F-RL-01: max_per_minute=0 sẽ block mọi request vĩnh viễn — validate ngay
        if max_per_minute < 1:
            raise ValueError(f"max_per_minute phải >= 1, nhận được: {max_per_minute}")
        self._max = max_per_minute
        # Bắt đầu với 10% capacity thay vì full bucket để giảm burst khởi động
        # Đảm bảo tối thiểu 5 tokens để không block request đầu tiên
        self._tokens = min(float(max_per_minute), float(max_per_minute) * 0.1 + 5.0)
        self._last_refill = time.monotonic()
        # threading.Lock an toàn ở đây: acquire() chỉ được gọi từ asyncio event loop thread
        # (single-threaded), nên lock không bao giờ contended → không block event loop.
        self._lock = __import__('threading').Lock()

    def acquire(self) -> bool:
        """
        Thử lấy 1 token từ thùng.

        Trả về:
          True nếu còn token (yêu cầu được phép tiếp tục).
          False nếu thùng rỗng (yêu cầu bị từ chối vì vượt giới hạn).
        """
        with self._lock:
            now = time.monotonic()
            elapsed = now - self._last_refill
            # Nạp lại token theo tỷ lệ thời gian đã trôi qua kể từ lần nạp trước
            self._tokens = min(
                float(self._max),
                self._tokens + elapsed * (self._max / 60.0)
            )
            self._last_refill = now
            if self._tokens >= 1.0:
                # Còn token — tiêu thụ 1 token và cho phép yêu cầu
                self._tokens -= 1.0
                return True
            # Thùng rỗng — từ chối yêu cầu
            return False

    def update_limit(self, max_per_minute: int) -> None:
        """
        Cập nhật giới hạn tốc độ mới (dùng khi config thay đổi lúc runtime).

        Tham số:
          max_per_minute: Giới hạn mới cần áp dụng.
        """
        with self._lock:
            self._max = max_per_minute


# Biến toàn cục cho rate limiter — được khởi tạo trong _initialize_components()
_rate_limiter: _TokenBucketRateLimiter | None = None


# ===========================================================================
# PHẦN 1: KHỞI TẠO VÀ KIỂM TRA MÔI TRƯỜNG
# ===========================================================================

def _setup_internal_logging(debug: bool = False) -> None:
    """
    Cấu hình hệ thống log nội bộ dành cho developer.

    Log này ghi ra stderr, KHÔNG ghi ra stdout để không làm nhiễu MCP protocol
    (MCP dùng stdout để giao tiếp với Claude Code CLI).

    Tham số:
      debug: Nếu True, bật log mức DEBUG để xem chi tiết hơn.
    """
    level = logging.DEBUG if debug else logging.INFO
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(
        logging.Formatter(
            fmt="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
            datefmt="%Y-%m-%dT%H:%M:%S",
        )
    )
    logging.basicConfig(level=level, handlers=[handler])
    # Tắt log của thư viện bên thứ ba để không làm nhiễu
    logging.getLogger("mcp").setLevel(logging.WARNING)
    logging.getLogger("asyncio").setLevel(logging.WARNING)


def _check_outlook_running() -> bool:
    """
    Kiểm tra xem Outlook Desktop có đang chạy không bằng cách kết nối COM.

    Dùng GetActiveObject() (chỉ kết nối vào Outlook đang mở) thay vì
    Dispatch() (có thể tự động mở Outlook mới — không mong muốn).

    Trả về:
      True nếu Outlook đang chạy và có thể kết nối COM.
      False nếu Outlook chưa chạy hoặc COM không khả dụng.
    """
    try:
        import win32com.client  # type: ignore
        import pythoncom        # type: ignore

        # Bước 1: Khởi tạo COM trong thread hiện tại (bắt buộc trước khi gọi COM)
        pythoncom.CoInitialize()
        try:
            # Bước 2: Thử kết nối vào Outlook đang mở — không tạo mới
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            # Bước 3: Đọc một thuộc tính đơn giản để xác nhận kết nối thực sự hoạt động
            _version = outlook.Version
            # Bước 4: Giải phóng COM object ngay sau khi kiểm tra xong
            win32com.client.ReleaseComObject(outlook)
            del outlook
            gc.collect()
            return True
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        # Không log chi tiết lỗi COM vì có thể chứa thông tin hệ thống
        return False


def _initialize_components() -> None:
    """
    Khởi tạo tất cả các thành phần cốt lõi của server theo thứ tự phụ thuộc.

    Thứ tự bắt buộc:
      1. Config — phải load trước vì các thành phần khác phụ thuộc vào config
      2. AuditLogger — cần config để biết đường dẫn log
      3. OutlookCOMBridge — cần config để biết account và allowed folders
      4. ThreadPoolExecutor — COM thread pool, khởi tạo sau cùng

    Nếu bất kỳ bước nào thất bại, server không thể khởi động và sẽ raise exception.
    """
    global _config, _audit, _com_bridge, _com_executor, _rate_limiter

    # Bước 1: Load cấu hình từ config.toml
    _internal_logger.info("Đang tải cấu hình từ config.toml...")
    try:
        from config import load_config  # type: ignore
        _config = load_config()
        _internal_logger.info("Cấu hình tải thành công.")
    except FileNotFoundError:
        _internal_logger.error(
            "Không tìm thấy file config.toml. "
            "Chạy 'python server.py --setup' để tạo cấu hình."
        )
        raise
    except Exception as exc:
        _internal_logger.error("Lỗi tải cấu hình: %s", exc)
        raise

    # Bước 1b: Khởi tạo rate limiter ngay sau khi có config — giới hạn tốc độ yêu cầu
    _rate_limiter = _TokenBucketRateLimiter(max_per_minute=_config.MAX_CALLS_PER_MINUTE)
    _internal_logger.info("Rate limiter khởi tạo: %d req/phút", _config.MAX_CALLS_PER_MINUTE)

    # Bước 2: Khởi tạo AuditLogger — phải có trước khi ghi bất kỳ sự kiện nào
    _internal_logger.info("Đang khởi tạo Audit Logger...")
    try:
        from security.audit import AuditLogger  # type: ignore
        _audit = AuditLogger(
            config=_config,
            session_id=SESSION_ID,
            server_version=SERVER_VERSION,
        )
        _internal_logger.info("Audit Logger sẵn sàng.")
    except Exception as exc:
        _internal_logger.error("Lỗi khởi tạo Audit Logger: %s", exc)
        raise

    # Bước 3: Khởi tạo COM Bridge để giao tiếp với Outlook
    _internal_logger.info("Đang khởi tạo Outlook COM Bridge...")
    try:
        from outlook_com import OutlookCOMBridge  # type: ignore
        _com_bridge = OutlookCOMBridge(config=_config)
        _internal_logger.info("Outlook COM Bridge sẵn sàng.")
    except Exception as exc:
        _internal_logger.error("Lỗi khởi tạo COM Bridge: %s", exc)
        raise

    # Bước 4: Tạo thread pool COM — chỉ 1 worker để đảm bảo STA threading
    _com_executor = ThreadPoolExecutor(
        max_workers=1,
        thread_name_prefix="OutlookCOMThread",
    )
    _internal_logger.info("COM ThreadPoolExecutor sẵn sàng (1 worker).")


# ===========================================================================
# PHẦN 2: ĐIỀU PHỐI TOOL CALLS
# ===========================================================================

async def _run_in_com_thread(func: Any, *args: Any, **kwargs: Any) -> Any:
    """
    Chạy một hàm đồng bộ trong COM thread riêng và chờ kết quả bất đồng bộ.

    Lý do cần hàm này:
      - asyncio không thể trực tiếp gọi các hàm blocking (win32com) vì sẽ
        làm đóng băng toàn bộ event loop.
      - Tất cả COM operations phải chạy trong cùng một thread (STA constraint).
      - run_in_executor() cho phép asyncio chạy hàm blocking trong thread pool
        mà không blocking event loop.

    Tham số:
      func: Hàm đồng bộ cần chạy trong COM thread.
      *args, **kwargs: Tham số truyền vào func.

    Trả về:
      Kết quả của func, được chờ bất đồng bộ.
    """
    loop = asyncio.get_running_loop()
    try:
        return await asyncio.wait_for(
            loop.run_in_executor(
                _com_executor,
                lambda: func(*args, **kwargs),
            ),
            timeout=30.0,
        )
    except asyncio.TimeoutError:
        raise asyncio.TimeoutError(
            "COM operation timeout sau 30 giây — Outlook có thể đang bận"
        )


def _build_error_response(message: str) -> list[TextContent]:
    """
    Tạo response lỗi chuẩn để trả về Claude.

    Quy tắc bảo mật: Không bao giờ trả về stack trace hay thông tin hệ thống
    nội bộ về lỗi. Chỉ trả về thông điệp ngắn gọn, thân thiện với người dùng.

    Tham số:
      message: Thông điệp lỗi bằng tiếng Việt, không chứa chi tiết kỹ thuật.

    Trả về:
      Danh sách chứa một TextContent với JSON {"error": message}.
    """
    return [TextContent(type="text", text=json.dumps({"error": message}, ensure_ascii=False))]


def _build_success_response(data: dict) -> list[TextContent]:
    """
    Tạo response thành công chuẩn để trả về Claude.

    Tham số:
      data: Dictionary kết quả từ tool handler.

    Trả về:
      Danh sách chứa một TextContent với JSON của data.
    """
    return [TextContent(type="text", text=json.dumps(data, ensure_ascii=False, default=str))]


# ===========================================================================
# PHẦN 3: ĐỊNH NGHĨA TOOL SCHEMAS (JSON Schema cho từng tool)
# ===========================================================================

def _get_tool_definitions() -> list[Tool]:
    """
    Trả về danh sách định nghĩa đầy đủ cho 19 công cụ MCP.

    Mỗi Tool bao gồm:
      - name: Tên gọi từ Claude
      - description: Mô tả bằng tiếng Anh (Claude đọc để hiểu khi nào dùng tool này)
      - inputSchema: JSON Schema xác định các tham số hợp lệ

    Lưu ý: inputSchema dùng additionalProperties=false để từ chối tham số không
    được khai báo — ngăn chặn injection qua tham số không hợp lệ.
    """
    return [
        # Tool 1: Liệt kê thư mục email
        Tool(
            name="list_folders",
            description=(
                "List all accessible email folders in Outlook. "
                "Only returns folders configured in the allowed_folders list. "
                "Use this before list_emails to discover valid folder paths."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "include_subfolders": {
                        "type": "boolean",
                        "default": False,
                        "description": "Whether to include subfolders recursively.",
                    }
                },
                "required": [],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=True, idempotentHint=True),
        ),

        # Tool 2: Liệt kê email trong một thư mục
        Tool(
            name="list_emails",
            description=(
                "List emails in a specific folder. Returns metadata only "
                "(subject, sender, date, size) — not the full body. "
                "Use read_email to get the full content of a specific email."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "maxLength": 260,
                        "description": "Folder path from list_folders, e.g. 'Inbox' or 'Inbox/Projects'.",
                    },
                    "limit": {
                        "type": "integer",
                        "default": 20,
                        "maximum": 100,
                        "minimum": 1,
                        "description": "Maximum number of emails to return (capped at 100).",
                    },
                    "offset": {
                        "type": "integer",
                        "default": 0,
                        "minimum": 0,
                        "description": "Number of emails to skip for pagination.",
                    },
                    "unread_only": {
                        "type": "boolean",
                        "default": False,
                        "description": "If true, only return unread emails.",
                    },
                },
                "required": ["folder_path"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=True, idempotentHint=True),
        ),

        # Tool 3: Đọc nội dung đầy đủ của một email
        Tool(
            name="read_email",
            description=(
                "Read the full content of a specific email by its entry_id. "
                "Returns subject, sender, recipients, body text, and attachment info. "
                "HTML is stripped to plain text. "
                "Obtain entry_id from list_emails or search_emails."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "pattern": "^[0-9A-Fa-f]+$",
                        "maxLength": 256,
                        "description": "Hexadecimal entry ID of the email from list_emails.",
                    }
                },
                "required": ["entry_id"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=True, idempotentHint=True),
        ),

        # Tool 4: Tìm kiếm email
        Tool(
            name="search_emails",
            description=(
                "Search emails by keyword across allowed folders. "
                "Uses Outlook's built-in DASL filter for performance. "
                "Returns a snippet of matching emails, not full bodies."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "maxLength": 200,
                        "description": "Search keyword or phrase (max 200 characters).",
                    },
                    "folder_path": {
                        "type": "string",
                        "maxLength": 260,
                        "description": "Restrict search to this folder. Searches all allowed folders if omitted.",
                    },
                    "search_in": {
                        "type": "string",
                        "enum": ["subject", "body", "sender", "all"],
                        "default": "subject",
                        "description": "Which part of the email to search.",
                    },
                    "date_from": {
                        "type": "string",
                        "format": "date",
                        "description": "Filter emails on or after this date (YYYY-MM-DD).",
                    },
                    "date_to": {
                        "type": "string",
                        "format": "date",
                        "description": "Filter emails on or before this date (YYYY-MM-DD).",
                    },
                    "limit": {
                        "type": "integer",
                        "default": 20,
                        "maximum": 50,
                        "minimum": 1,
                        "description": "Maximum results to return (capped at 50).",
                    },
                },
                "required": ["query"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=True, idempotentHint=True),
        ),

        # Tool 5: Soạn email nháp mới
        Tool(
            name="compose_draft",
            description=(
                "Open a new email compose window in Outlook with pre-filled fields. "
                "The user MUST manually review and click Send — this tool NEVER sends automatically. "
                "Requires read_only_mode=false in config. "
                "Returns immediately after opening the compose window."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "maxItems": 50,
                        "description": "List of recipient email addresses.",
                    },
                    "cc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "maxItems": 50,
                        "default": [],
                        "description": "List of CC email addresses.",
                    },
                    "bcc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "maxItems": 50,
                        "default": [],
                        "description": "BCC (blind copy) email addresses — not visible to other recipients.",
                    },
                    "subject": {
                        "type": "string",
                        "maxLength": 500,
                        "description": "Email subject line.",
                    },
                    "body": {
                        "type": "string",
                        "maxLength": 50000,
                        "description": "Email body in plain text.",
                    },
                    "importance": {
                        "type": "string",
                        "enum": ["low", "normal", "high"],
                        "default": "normal",
                        "description": "Email importance level.",
                    },
                },
                "required": ["to", "subject", "body"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False),
        ),

        # Tool 6: Soạn email trả lời
        Tool(
            name="reply_draft",
            description=(
                "Open a reply window in Outlook for an existing email. "
                "The user MUST manually review and click Send — this tool NEVER sends automatically. "
                "Requires read_only_mode=false in config. "
                "Returns immediately after opening the reply window."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "pattern": "^[0-9A-Fa-f]+$",
                        "maxLength": 256,
                        "description": "Entry ID of the email to reply to.",
                    },
                    "body": {
                        "type": "string",
                        "maxLength": 50000,
                        "description": "Reply body in plain text.",
                    },
                    "reply_all": {
                        "type": "boolean",
                        "default": False,
                        "description": "If true, reply to all recipients.",
                    },
                    "additional_cc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "maxItems": 20,
                        "default": [],
                        "description": "Additional CC addresses to add to the reply.",
                    },
                },
                "required": ["entry_id", "body"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False),
        ),

        # Tool 7: Chuyển tiếp email
        Tool(
            name="forward_draft",
            description=(
                "Open a forward window in Outlook to forward an existing email to new recipients. "
                "The user MUST manually review and click Send — this tool NEVER sends automatically. "
                "Requires read_only_mode=false in config."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "pattern": "^[0-9A-Fa-f]+$",
                        "maxLength": 256,
                        "description": "Entry ID of the email to forward.",
                    },
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "maxItems": 50,
                        "description": "Recipient email addresses to forward to.",
                    },
                    "note": {
                        "type": "string",
                        "maxLength": 50000,
                        "default": "",
                        "description": "Optional note to prepend before the original email body.",
                    },
                },
                "required": ["entry_id", "to"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False),
        ),

        # Tool 8: Đánh dấu email đã đọc / chưa đọc
        Tool(
            name="mark_email_read",
            description=(
                "Mark an email as read or unread in Outlook. "
                "Requires read_only_mode=false in config."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "pattern": "^[0-9A-Fa-f]+$",
                        "maxLength": 256,
                        "description": "Entry ID of the email to mark.",
                    },
                    "read": {
                        "type": "boolean",
                        "default": True,
                        "description": "True to mark as read, false to mark as unread.",
                    },
                },
                "required": ["entry_id"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False),
        ),

        # Tool 9: Đặt flag theo dõi trên email
        Tool(
            name="flag_email",
            description=(
                "Set or clear a follow-up flag on an email in Outlook. "
                "Requires read_only_mode=false in config."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "pattern": "^[0-9A-Fa-f]+$",
                        "maxLength": 256,
                        "description": "Entry ID of the email.",
                    },
                    "flag_status": {
                        "type": "string",
                        "enum": ["flagged", "complete", "none"],
                        "default": "flagged",
                        "description": "'flagged'=mark for follow-up, 'complete'=done, 'none'=clear flag.",
                    },
                },
                "required": ["entry_id"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False),
        ),

        # Tool 10: Di chuyển email sang thư mục khác
        Tool(
            name="move_email",
            description=(
                "Move an email to another allowed folder in Outlook. "
                "Both source and destination folders must be in the allowed_folders list. "
                "Requires read_only_mode=false in config."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "pattern": "^[0-9A-Fa-f]+$",
                        "maxLength": 256,
                        "description": "Entry ID of the email to move.",
                    },
                    "destination_folder": {
                        "type": "string",
                        "maxLength": 260,
                        "description": "Name of the destination folder (must be in allowed_folders).",
                    },
                },
                "required": ["entry_id", "destination_folder"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False),
        ),

        # Tool 11: Liệt kê tất cả thư mục (không giới hạn allowlist)
        Tool(
            name="list_all_folders",
            description=(
                "List ALL folders in all Outlook stores (PST files, mailboxes) recursively. "
                "Unlike list_folders, this includes folders outside the allowed_folders list. "
                "Use this to discover the full folder structure and update config.toml."
            ),
            inputSchema={
                "type": "object",
                "properties": {},
                "required": [],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=True, idempotentHint=True),
        ),

        # Tool 12: Thống kê email theo thư mục
        Tool(
            name="email_stats",
            description=(
                "Get email statistics (total count and unread count) for all allowed folders. "
                "Returns a summary and per-folder breakdown."
            ),
            inputSchema={
                "type": "object",
                "properties": {},
                "required": [],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=True, idempotentHint=True),
        ),

        # Tool 13: Lấy toàn bộ email trong conversation thread
        Tool(
            name="get_email_thread",
            description=(
                "Get all emails in the same conversation thread as a given email. "
                "Uses Outlook ConversationID to find related emails across allowed folders. "
                "Returns emails sorted oldest-first."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "pattern": "^[0-9A-Fa-f]+$",
                        "maxLength": 512,
                        "description": "Entry ID of any email in the thread.",
                    },
                    "max_emails": {
                        "type": "integer",
                        "default": 20,
                        "minimum": 1,
                        "maximum": 50,
                        "description": "Maximum number of emails to return from the thread.",
                    },
                },
                "required": ["entry_id"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=True, idempotentHint=True),
        ),

        # Tool 14: Thống kê email theo contact
        Tool(
            name="get_contact_stats",
            description=(
                "Get email statistics for a specific contact: "
                "how many emails received from them across allowed folders."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "email": {
                        "type": "string",
                        "maxLength": 320,
                        "description": "Email address of the contact to look up.",
                    },
                },
                "required": ["email"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=True, idempotentHint=True),
        ),

        # Tool 15: Đánh dấu nhiều email đã đọc (bulk)
        Tool(
            name="bulk_mark_read",
            description=(
                "Mark all unread emails in a folder as read. "
                "Use dry_run=true first to preview what will be changed. "
                "Requires read_only_mode=false in config."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_name": {
                        "type": "string",
                        "maxLength": 260,
                        "description": "Name of the folder to process (must be in allowed_folders).",
                    },
                    "dry_run": {
                        "type": "boolean",
                        "default": True,
                        "description": "If true (default), preview only — no changes made.",
                    },
                    "max_emails": {
                        "type": "integer",
                        "default": 50,
                        "minimum": 1,
                        "maximum": 100,
                        "description": "Maximum number of emails to process.",
                    },
                },
                "required": ["folder_name"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=True),
        ),

        # Tool 16: Lấy danh sách email flagged trong một folder
        Tool(
            name="get_flagged_emails",
            description=(
                "Get all flagged (follow-up) emails in a specific folder. "
                "Returns emails where the user has set a follow-up flag in Outlook. "
                "Useful for PM daily review: what emails need my attention today?"
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_name": {
                        "type": "string",
                        "maxLength": 260,
                        "description": "Folder to search for flagged emails (must be in allowed_folders).",
                    },
                },
                "required": ["folder_name"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=True, idempotentHint=True),
        ),

        # Tool 17: Snapshot tong hop trang thai mot project folder
        Tool(
            name="get_project_snapshot",
            description=(
                "Get a comprehensive status snapshot of a project folder for PM review. "
                "Returns in ONE call: total emails received, unread count, flagged follow-up items, "
                "top senders, and a plain-language summary. "
                "Replaces 4-5 separate queries for morning briefing or pre-meeting prep."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_name": {
                        "type": "string",
                        "maxLength": 260,
                        "description": "Project folder name (must be in allowed_folders).",
                    },
                    "days_back": {
                        "type": "integer",
                        "default": 14,
                        "minimum": 1,
                        "maximum": 90,
                        "description": "How many days back to look (default: 14).",
                    },
                },
                "required": ["folder_name"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=True, idempotentHint=True),
        ),

        # Tool 18: Liệt kê sự kiện lịch Outlook sắp tới
        Tool(
            name="list_calendar_events",
            description=(
                "List upcoming calendar events from Outlook Calendar. "
                "Returns: title, start/end time, location, organizer, attendees. "
                "Use to view meeting schedule and find related emails before meetings."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "days_ahead": {
                        "type": "integer",
                        "default": 7,
                        "minimum": 0,
                        "maximum": 90,
                        "description": "Number of days ahead to look (default: 7, max: 90).",
                    },
                    "days_back": {
                        "type": "integer",
                        "default": 0,
                        "minimum": 0,
                        "maximum": 30,
                        "description": "Number of days back to include (default: 0 = only future events).",
                    },
                },
                "required": [],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=True, idempotentHint=True),
        ),

        # Tool 19: Tạo sự kiện lịch / gửi lời mời họp
        Tool(
            name="create_calendar_event",
            description=(
                "Create a new calendar event in Outlook and open it for review. "
                "If required_attendees is provided, opens as a Meeting Request. "
                "The user MUST manually click Send in Outlook to send invitations. "
                "This tool NEVER sends invitations automatically."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "subject": {
                        "type": "string",
                        "maxLength": 500,
                        "description": "Event title (required).",
                    },
                    "start": {
                        "type": "string",
                        "description": "Start time in 'YYYY-MM-DD HH:MM' format, e.g. '2026-07-15 14:00'.",
                    },
                    "end": {
                        "type": "string",
                        "description": "End time in 'YYYY-MM-DD HH:MM' format, e.g. '2026-07-15 15:00'.",
                    },
                    "location": {
                        "type": "string",
                        "maxLength": 500,
                        "default": "",
                        "description": "Meeting location or Teams/Zoom link (optional).",
                    },
                    "body": {
                        "type": "string",
                        "maxLength": 10000,
                        "default": "",
                        "description": "Meeting agenda or description (optional).",
                    },
                    "required_attendees": {
                        "type": "array",
                        "items": {"type": "string"},
                        "maxItems": 20,
                        "description": "Email addresses of required attendees (optional). Triggers Meeting Request mode.",
                    },
                },
                "required": ["subject", "start", "end"],
                "additionalProperties": False,
            },
            annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False),
        ),
    ]


# ===========================================================================
# PHẦN 4: TẠO VÀ CẤU HÌNH MCP SERVER
# ===========================================================================

def _create_mcp_server() -> Server:
    """
    Tạo và cấu hình MCP Server với toàn bộ tool handlers.

    Hàm này:
      1. Khởi tạo Server object
      2. Đăng ký handler cho list_tools (Claude hỏi "có những tool nào?")
      3. Đăng ký handler cho call_tool (Claude gọi một tool cụ thể)

    Trả về:
      Server object đã được cấu hình đầy đủ, sẵn sàng chạy.
    """
    server = Server(SERVER_NAME)

    # --- Đăng ký handler: Claude hỏi "có những tool nào?" ---
    @server.list_tools()
    async def handle_list_tools() -> list[Tool]:
        """Trả về danh sách định nghĩa tất cả các tool cho Claude."""
        return _get_tool_definitions()

    # --- Đăng ký handler: Claude gọi một tool cụ thể ---
    @server.call_tool()
    async def handle_call_tool(name: str, arguments: dict) -> list[TextContent]:
        """
        Xử lý mỗi tool call từ Claude theo pipeline bắt buộc:
          1. Ghi audit log (bắt đầu) — luôn luôn, ngay cả khi sau đó bị reject
          2. Validate input — từ chối ngay nếu không hợp lệ
          3. Dispatch đến tool handler tương ứng (trong COM thread)
          4. Ghi audit log (kết quả)
          5. Trả về kết quả JSON

        Lưu ý: Hàm này là async nhưng không trực tiếp gọi COM — tất cả COM
        operations được đẩy vào _com_executor qua _run_in_com_thread().

        Tham số:
          name: Tên tool, ví dụ "list_emails"
          arguments: Dict chứa tham số từ Claude (chưa được validate)

        Trả về:
          Danh sách TextContent với JSON kết quả hoặc lỗi.
        """
        # Đo thời gian thực thi để ghi vào audit log
        start_time = time.monotonic()

        # Bước 1: Ghi audit log — bắt đầu xử lý tool call
        # Audit log ghi TRƯỚC khi xử lý để đảm bảo không bị mất nếu server crash
        try:
            _audit.log_tool_start(tool_name=name, raw_arguments=arguments)
        except Exception as audit_exc:
            # Nguyên tắc fail-closed: nếu không ghi được audit log thì KHÔNG được tiếp tục xử lý.
            # Cho phép tool call tiếp tục khi audit thất bại sẽ tạo ra lỗ hổng bypass kiểm soát.
            _internal_logger.error(
                "Audit log thất bại cho tool '%s' — từ chối xử lý (fail-closed): %s",
                name,
                audit_exc,
            )
            return _build_error_response(
                "Không thể ghi nhật ký kiểm toán (audit log). Yêu cầu bị từ chối vì lý do an toàn."
            )

        # Bước 1b: Rate limiting — từ chối ngay nếu số yêu cầu vượt quá giới hạn cấu hình
        # Kiểm tra này chạy SAU khi đã ghi audit log để không bỏ sót yêu cầu bị chặn
        if _rate_limiter is None:
            # Cảnh báo khi rate limiter chưa được khởi tạo
            _internal_logger.warning(
                "Rate limiter chưa được khởi tạo — tool call '%s' không được kiểm soát tốc độ", name
            )
        elif not _rate_limiter.acquire():
            try:
                _audit.log_tool_blocked(
                    tool_name=name,
                    block_reason="rate_limit_exceeded",
                    duration_ms=0,
                )
            except Exception as audit_err:
                # Ghi vào internal logger khi audit thất bại — không để event biến mất
                _internal_logger.error(
                    "Không thể ghi tool_blocked event vào audit log: tool=%s, err=%s",
                    name, audit_err
                )
            return _build_error_response("Quá nhiều yêu cầu. Vui lòng thử lại sau vài giây.")

        # Bước 2: Kiểm tra tool name có hợp lệ không
        valid_tool_names = {
            "list_folders",
            "list_emails",
            "read_email",
            "search_emails",
            "compose_draft",
            "reply_draft",
            "forward_draft",
            "mark_email_read",
            "flag_email",
            "move_email",
            "list_all_folders",
            "email_stats",
            "get_email_thread",
            "get_contact_stats",
            "bulk_mark_read",
            "get_flagged_emails",
            "get_project_snapshot",
            "list_calendar_events",
            "create_calendar_event",
        }
        if name not in valid_tool_names:
            duration_ms = int((time.monotonic() - start_time) * 1000)
            try:
                _audit.log_tool_blocked(
                    tool_name=name,
                    block_reason="unknown_tool",
                    duration_ms=duration_ms,
                )
            except Exception:
                pass
            return _build_error_response(f"Công cụ '{name}' không tồn tại.")

        # Bước 3: Dispatch đến tool handler theo tên
        # Mỗi tool handler nằm trong module riêng trong tools/
        try:
            result = await _dispatch_tool(name, arguments)
            duration_ms = int((time.monotonic() - start_time) * 1000)

            # Bước 4: Ghi audit log thành công
            try:
                items_returned = None
                if isinstance(result, dict):
                    # Đếm số items để ghi log (không ghi nội dung)
                    for count_key in ("emails", "folders", "results"):
                        if count_key in result:
                            items_returned = len(result[count_key])
                            break
                _audit.log_tool_success(
                    tool_name=name,
                    duration_ms=duration_ms,
                    items_returned=items_returned,
                )
            except Exception:
                _internal_logger.warning("Không thể ghi audit log thành công cho tool: %s", name)

            return _build_success_response(result)

        # Bước 5: Xử lý các loại lỗi khác nhau theo phân loại từ PLAN.md
        except Exception as exc:
            duration_ms = int((time.monotonic() - start_time) * 1000)
            error_message, block_reason = _classify_exception(exc)

            try:
                if block_reason:
                    # Lỗi bảo mật / validation — ghi là "blocked"
                    _audit.log_tool_blocked(
                        tool_name=name,
                        block_reason=block_reason,
                        duration_ms=duration_ms,
                    )
                else:
                    # Lỗi kỹ thuật — ghi là "error" (không ghi message gốc)
                    _audit.log_tool_error(
                        tool_name=name,
                        error_code=type(exc).__name__,
                        duration_ms=duration_ms,
                    )
            except Exception:
                pass

            # Ghi stack trace vào internal log để developer debug
            # KHÔNG ghi vào audit log, KHÔNG trả về cho Claude
            _internal_logger.debug(
                "Tool '%s' thất bại (chi tiết):\n%s",
                name,
                traceback.format_exc(),
            )

            return _build_error_response(error_message)

    return server


def _classify_exception(exc: Exception) -> tuple[str, str | None]:
    """
    Phân loại exception thành thông điệp người dùng và block_reason cho audit.

    Quy tắc:
      - Lỗi có tên "Error" cụ thể từ codebase nội bộ → thông điệp cụ thể.
      - Mọi exception khác → thông điệp chung, không lộ chi tiết kỹ thuật.

    Tham số:
      exc: Exception đã bắt được.

    Trả về:
      Tuple (error_message_for_user, block_reason_for_audit).
      block_reason là None nếu đây là lỗi kỹ thuật (không phải bị chặn).
    """
    exc_type_name = type(exc).__name__

    # Lỗi validation input — bị chặn vì input không hợp lệ
    if exc_type_name in ("ValidationError", "ValueError") or "Validation" in exc_type_name:
        return f"Input không hợp lệ: {str(exc)}", "validation_error"

    # Lỗi thư mục không được phép truy cập
    if "FolderNotAllowed" in exc_type_name or "NotAllowed" in exc_type_name:
        return "Thư mục không được phép truy cập.", "not_in_allowlist"

    # Lỗi chế độ chỉ đọc — tool cần write nhưng read_only đang bật
    if "ReadOnly" in exc_type_name or "ReadOnlyMode" in exc_type_name:
        return (
            "Chế độ chỉ đọc đang bật. "
            "Cần tắt read_only_mode trong config.toml để dùng tính năng này.",
            "read_only_mode",
        )

    # Lỗi Outlook không đang chạy
    if "OutlookNotRunning" in exc_type_name or "OutlookNotFound" in exc_type_name:
        return (
            "Outlook không đang chạy. "
            "Vui lòng mở Outlook Desktop trước khi dùng công cụ này.",
            None,
        )

    # Lỗi thao tác Outlook (COM error đã được wrap)
    if "OutlookOperation" in exc_type_name or "COMError" in exc_type_name:
        return (
            "Không thể thực hiện thao tác trong Outlook. "
            "Đảm bảo Outlook đang chạy và không bị treo.",
            None,
        )

    # Lỗi rate limiting — quá nhiều request
    if "RateLimit" in exc_type_name:
        return "Quá nhiều yêu cầu. Vui lòng thử lại sau vài giây.", "rate_limit"

    # Mọi lỗi khác — thông điệp chung, không lộ chi tiết
    return "Lỗi nội bộ. Kiểm tra log để biết thêm chi tiết.", None


async def _dispatch_tool(name: str, arguments: dict) -> dict:
    """
    Điều phối tool call đến module handler tương ứng.

    Mỗi tool được import lazily (chỉ khi cần) để tránh lỗi import nếu
    một module nào đó chưa được cài đặt đầy đủ dependency.

    Tham số:
      name: Tên tool cần gọi.
      arguments: Dict tham số đã được validate trước (bởi tool handler).

    Trả về:
      Dict kết quả từ tool handler.
    """
    if name == "list_folders":
        from tools.list_folders import handle_list_folders  # type: ignore
        return await _run_in_com_thread(
            handle_list_folders,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "list_emails":
        from tools.read_email import handle_list_emails  # type: ignore
        return await _run_in_com_thread(
            handle_list_emails,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "read_email":
        from tools.read_email import handle_read_email  # type: ignore
        return await _run_in_com_thread(
            handle_read_email,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "search_emails":
        from tools.search import handle_search_emails  # type: ignore
        return await _run_in_com_thread(
            handle_search_emails,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "compose_draft":
        from tools.compose import handle_compose_draft  # type: ignore
        return await _run_in_com_thread(
            handle_compose_draft,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "reply_draft":
        from tools.compose import handle_reply_draft  # type: ignore
        return await _run_in_com_thread(
            handle_reply_draft,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "forward_draft":
        from tools.compose import handle_forward_draft  # type: ignore
        return await _run_in_com_thread(
            handle_forward_draft,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "mark_email_read":
        from tools.manage_email import handle_mark_read  # type: ignore
        return await _run_in_com_thread(
            handle_mark_read,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "flag_email":
        from tools.manage_email import handle_flag_email  # type: ignore
        return await _run_in_com_thread(
            handle_flag_email,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "move_email":
        from tools.manage_email import handle_move_email  # type: ignore
        return await _run_in_com_thread(
            handle_move_email,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "list_all_folders":
        from tools.list_folders import handle_list_all_folders  # type: ignore
        return await _run_in_com_thread(
            handle_list_all_folders,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "email_stats":
        from tools.list_folders import handle_email_stats  # type: ignore
        return await _run_in_com_thread(
            handle_email_stats,
            arguments,
            _config,
            _com_bridge,
        )

    elif name == "get_email_thread":
        from tools.manage_email import handle_get_email_thread  # type: ignore
        return await _run_in_com_thread(handle_get_email_thread, arguments, _config, _com_bridge)

    elif name == "get_contact_stats":
        from tools.list_folders import handle_get_contact_stats  # type: ignore
        return await _run_in_com_thread(handle_get_contact_stats, arguments, _config, _com_bridge)

    elif name == "bulk_mark_read":
        from tools.manage_email import handle_bulk_mark_read  # type: ignore
        return await _run_in_com_thread(handle_bulk_mark_read, arguments, _config, _com_bridge)

    elif name == "get_flagged_emails":
        from tools.manage_email import handle_get_flagged_emails  # type: ignore
        return await _run_in_com_thread(handle_get_flagged_emails, arguments, _config, _com_bridge)

    elif name == "get_project_snapshot":
        from tools.manage_email import handle_get_project_snapshot  # type: ignore
        return await _run_in_com_thread(handle_get_project_snapshot, arguments, _config, _com_bridge)

    elif name == "list_calendar_events":
        from tools.calendar import handle_list_calendar_events  # type: ignore
        return await _run_in_com_thread(handle_list_calendar_events, arguments, _config, _com_bridge)

    elif name == "create_calendar_event":
        from tools.calendar import handle_create_calendar_event  # type: ignore
        return await _run_in_com_thread(handle_create_calendar_event, arguments, _config, _com_bridge)

    # Không nên đến đây vì đã kiểm tra ở trên, nhưng để an toàn
    raise ValueError(f"Tool không xác định: {name}")


# ===========================================================================
# PHẦN 5: GRACEFUL SHUTDOWN
# ===========================================================================

def _teardown_components() -> None:
    """
    Dọn dẹp tài nguyên khi server tắt — đảm bảo không rò rỉ tài nguyên.

    Thứ tự dọn dẹp (ngược với thứ tự khởi tạo):
      1. Shutdown COM executor — chờ COM thread hoàn thành task hiện tại
      2. Đóng COM Bridge — giải phóng COM objects và gọi CoUninitialize
      3. Flush audit log — ghi entry shutdown và đóng file
      4. Thu gom rác (garbage collection) để đảm bảo COM objects được release

    Mỗi bước được wrap trong try/except riêng để lỗi ở một bước không
    ngăn các bước còn lại thực thi.
    """
    _internal_logger.info("Đang dọn dẹp tài nguyên trước khi tắt...")

    # Bước 1: Shutdown COM executor — chờ task đang chạy hoàn thành
    global _com_executor
    if _com_executor is not None:
        try:
            _com_executor.shutdown(wait=True, cancel_futures=False)
            _internal_logger.info("COM ThreadPoolExecutor đã shutdown.")
        except Exception as exc:
            _internal_logger.warning("Lỗi khi shutdown COM executor: %s", exc)
        finally:
            _com_executor = None

    # Bước 2: Đóng COM Bridge
    if _com_bridge is not None:
        try:
            _com_bridge.close()
            _internal_logger.info("Outlook COM Bridge đã đóng.")
        except Exception as exc:
            _internal_logger.warning("Lỗi khi đóng COM Bridge: %s", exc)

    # Bước 3: Flush và đóng audit log — ghi entry server_stop
    if _audit is not None:
        try:
            _audit.log_server_stop()
            _audit.flush()
            _internal_logger.info("Audit log đã flush và đóng.")
        except Exception as exc:
            _internal_logger.warning("Lỗi khi đóng Audit Logger: %s", exc)

    # Bước 4: Thu gom rác để đảm bảo COM objects được release
    gc.collect()
    _internal_logger.info("Dọn dẹp hoàn tất.")


# ===========================================================================
# PHẦN 6: MCP SERVER RUNNER
# ===========================================================================

async def _run_mcp_server() -> None:
    """
    Vòng lặp chính chạy MCP server cho đến khi bị tắt.

    Luồng thực thi:
      1. Kiểm tra Outlook đang chạy
      2. Khởi tạo các thành phần
      3. Tạo MCP server object
      4. Ghi audit log server_start
      5. Chạy stdio_server (nhận/gửi MCP messages qua stdin/stdout)
      6. Dọn dẹp khi kết thúc (trong finally block)

    stdio_server là context manager — khi Claude Code CLI ngắt kết nối,
    context manager thoát ra và finally block được thực thi.
    """
    # Bước 1: Kiểm tra Outlook — cảnh báo nếu chưa mở, KHÔNG thoát
    # Lý do: Claude Code CLI khởi động MCP server trước khi Outlook mở; nếu exit ở đây
    # thì Claude Code sẽ báo "Failed to connect" mãi mãi. Tốt hơn là server chờ,
    # và khi tool được gọi mà Outlook chưa mở thì trả về lỗi rõ ràng cho người dùng.
    _internal_logger.info("Kiểm tra Outlook Desktop đang chạy...")
    if not _check_outlook_running():
        _internal_logger.warning(
            "Outlook Desktop chưa mở. MCP server vẫn khởi động bình thường. "
            "Hãy mở Outlook trước khi dùng các tool email."
        )
    else:
        _internal_logger.info("Outlook Desktop đang chạy. Tiếp tục khởi động...")

    # Bước 2: Khởi tạo tất cả thành phần
    try:
        _initialize_components()
    except Exception as exc:
        _internal_logger.error("Không thể khởi tạo server: %s", exc)
        sys.exit(1)

    # Bước 3: Tạo MCP server với tất cả tool handlers
    server = _create_mcp_server()

    # Bước 4: Ghi audit log server đã khởi động thành công
    try:
        _audit.log_server_start(
            version=SERVER_VERSION,
            read_only=_config.READ_ONLY_MODE,
            allowlist_count=len(_config.ALLOWED_FOLDERS),
        )
        _internal_logger.info(
            "MCP server '%s' v%s đã khởi động. Session: %s",
            SERVER_NAME,
            SERVER_VERSION,
            SESSION_ID,
        )
    except Exception as exc:
        _internal_logger.warning("Không thể ghi audit log server_start: %s", exc)

    # Bước 5: Chạy MCP server qua stdio transport
    # stdio_server() là context manager quản lý stdin/stdout streams
    # app.run() là vòng lặp chính xử lý MCP messages cho đến khi ngắt kết nối
    try:
        async with stdio_server() as (read_stream, write_stream):
            await server.run(
                read_stream,
                write_stream,
                server.create_initialization_options(),
            )
    finally:
        # Bước 6: Luôn dọn dẹp, dù server thoát bình thường hay do lỗi
        _teardown_components()


# ===========================================================================
# PHẦN 7: SETUP WIZARD
# ===========================================================================

def _run_setup_wizard() -> None:
    """
    Chạy wizard cài đặt lần đầu để hướng dẫn người dùng cấu hình.

    Wizard thực hiện:
      1. Kiểm tra Python version (yêu cầu 3.9+)
      2. Kiểm tra các dependencies cần thiết đã được cài chưa
      3. Kiểm tra Outlook COM khả dụng
      4. Tạo config.toml từ config.toml.example nếu chưa có
      5. Hướng dẫn người dùng cấu hình Windows Credential Manager
      6. Kiểm tra keyring backend (phải là WinVaultKeyring)
      7. Chạy test thử kết nối Outlook

    Wizard in kết quả ra stdout để người dùng thấy trực tiếp.
    """
    print("=" * 60)
    print(" Claude-Outlook MCP Secure — Wizard Cài Đặt Lần Đầu")
    print("=" * 60)
    print()

    all_ok = True

    # --- Kiểm tra Python version ---
    print("[1/6] Kiểm tra Python version...")
    py_version = sys.version_info
    if py_version >= (3, 9):
        print(f"  OK  Python {py_version.major}.{py_version.minor}.{py_version.micro}")
    else:
        print(f"  LỖI  Python {py_version.major}.{py_version.minor} — cần Python 3.9 trở lên.")
        all_ok = False

    # --- Kiểm tra các thư viện bắt buộc ---
    print("\n[2/6] Kiểm tra thư viện đã cài...")
    required_packages = {
        "mcp": "mcp",
        "win32com": "pywin32",
        "pythoncom": "pywin32",
        "keyring": "keyring",
        "tomllib or tomli": "tomllib/tomli",
        "bs4": "beautifulsoup4",
        "pydantic": "pydantic",
    }
    for import_name, pkg_name in required_packages.items():
        module_to_try = import_name.split(" or ")[0]
        try:
            __import__(module_to_try)
            print(f"  OK  {pkg_name}")
        except ImportError:
            # Thử fallback cho tomllib (Python 3.11+) → tomli (3.9-3.10)
            if "tomllib" in import_name:
                try:
                    import tomli  # type: ignore  # noqa: F401
                    print(f"  OK  tomli (fallback)")
                    continue
                except ImportError:
                    pass
            print(f"  THIẾU  {pkg_name} — chạy: pip install {pkg_name}")
            all_ok = False

    # --- Kiểm tra Outlook COM ---
    print("\n[3/6] Kiểm tra Outlook Desktop...")
    if _check_outlook_running():
        print("  OK  Outlook Desktop đang chạy và COM khả dụng.")
    else:
        print(
            "  CẢNH BÁO  Outlook không chạy hoặc COM không khả dụng.\n"
            "           Hãy mở Outlook trước khi dùng MCP server."
        )
        # Đây là cảnh báo, không phải lỗi blocking — user có thể setup trước

    # --- Tạo config.toml nếu chưa có ---
    print("\n[4/6] Kiểm tra file cấu hình...")
    import os
    config_path = os.path.join(os.path.dirname(__file__), "config.toml")
    example_path = os.path.join(os.path.dirname(__file__), "config.toml.example")

    if os.path.exists(config_path):
        print("  OK  config.toml đã tồn tại.")
    elif os.path.exists(example_path):
        import shutil
        shutil.copy(example_path, config_path)
        print(
            "  TẠO MỚI  config.toml được tạo từ config.toml.example.\n"
            "           Hãy mở và chỉnh sửa config.toml theo hướng dẫn."
        )
    else:
        print(
            "  THIẾU  Không tìm thấy config.toml.example.\n"
            "         Hãy tạo config.toml thủ công theo mẫu trong docs/PLAN.md."
        )
        all_ok = False

    # --- Kiểm tra Windows Credential Manager ---
    print("\n[5/6] Kiểm tra Windows Credential Manager (keyring)...")
    try:
        import keyring
        # Kiểm tra backend có phải WinVaultKeyring không
        backend = keyring.get_keyring()
        backend_name = type(backend).__name__
        if "Win" in backend_name or "Windows" in backend_name:
            print(f"  OK  Keyring backend: {backend_name}")
        else:
            print(
                f"  CẢNH BÁO  Keyring backend: {backend_name}\n"
                "           Khuyến nghị dùng WinVaultKeyring (Windows Credential Manager).\n"
                "           Cài pip install keyring[windows] nếu chưa có."
            )
    except Exception as exc:
        print(f"  LỖI  Không thể kiểm tra keyring: {exc}")
        all_ok = False

    # --- Hướng dẫn lưu API key ---
    print("\n[6/6] Hướng dẫn lưu Anthropic API Key...")
    print(
        "  Lưu API key vào Windows Credential Manager bằng lệnh:\n"
        "  python -c \""
        "import keyring; "
        "keyring.set_password('outlook-mcp-secure', 'anthropic_api_key', 'sk-ant-...')"
        "\"\n"
        "  Thay 'sk-ant-...' bằng API key thực của bạn từ https://console.anthropic.com/"
    )

    # --- Tổng kết ---
    print()
    print("=" * 60)
    if all_ok:
        print(" Cài đặt hoàn tất! Chạy lệnh sau để khởi động MCP server:")
        print()
        print("   python server.py")
        print()
        print(" Hoặc thêm vào Claude Code CLI:")
        print()
        print("   claude mcp add outlook-secure -- python server.py")
    else:
        print(" Có một số vấn đề cần giải quyết trước khi chạy server.")
        print(" Xem hướng dẫn chi tiết tại docs/USER_GUIDE.md")
    print("=" * 60)


# ===========================================================================
# PHẦN 8: ENTRY POINT CHÍNH
# ===========================================================================

def main() -> None:
    """
    Điểm vào chính khi chạy 'python server.py'.

    Phân tích tham số dòng lệnh:
      --setup  : Chạy wizard cài đặt lần đầu
      --debug  : Bật log mức DEBUG cho developer
      (không có tham số): Khởi động MCP server bình thường

    MCP server chạy trên asyncio event loop.
    """
    # Phân tích tham số dòng lệnh
    parser = argparse.ArgumentParser(
        prog="server.py",
        description="Claude-Outlook MCP Secure Server",
    )
    parser.add_argument(
        "--setup",
        action="store_true",
        help="Chạy wizard cài đặt lần đầu.",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Bật chế độ debug (log chi tiết hơn).",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {SERVER_VERSION}",
    )
    args = parser.parse_args()

    # Cấu hình hệ thống log nội bộ (ghi ra stderr)
    _setup_internal_logging(debug=args.debug)

    if args.setup:
        # Chế độ setup wizard — không cần asyncio
        _run_setup_wizard()
        return

    # Chế độ bình thường — khởi động MCP server
    # Dùng asyncio.run() để quản lý event loop tự động
    _internal_logger.info(
        "Khởi động %s v%s (session=%s)...",
        SERVER_NAME,
        SERVER_VERSION,
        SESSION_ID,
    )

    try:
        asyncio.run(_run_mcp_server())
    except KeyboardInterrupt:
        # Ctrl+C từ người dùng — thoát gọn gàng
        _internal_logger.info("Server bị dừng bởi người dùng (Ctrl+C).")
    except SystemExit:
        # sys.exit() từ bên trong _run_mcp_server
        raise
    except Exception as exc:
        _internal_logger.critical(
            "Lỗi không xử lý được — server dừng khẩn cấp: %s\n%s",
            exc,
            traceback.format_exc(),
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
