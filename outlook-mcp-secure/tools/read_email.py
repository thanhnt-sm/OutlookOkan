"""
Module read_email — Các MCP tools đọc và liệt kê email trong Outlook.

Cung cấp 2 tools:
  - list_emails: Liệt kê tóm tắt email trong một thư mục (không đọc nội dung body)
  - read_email: Đọc chi tiết một email theo entry_id, body được truncate theo giới hạn config

Bảo mật:
  - folder_name phải qua allowlist validation trước khi gọi COM
  - entry_id phải là hex hợp lệ, max 256 ký tự, không chứa null bytes
  - Body HTML được strip thành plain text trước khi trả về
  - Audit log không ghi subject, body, sender — chỉ ghi entry_id[:8]
  - Body bị cắt bớt theo MAX_BODY_CHARS để tránh trả về email quá lớn
"""

from __future__ import annotations

import json
import time
from typing import TYPE_CHECKING, Any

from mcp.server import Server
from mcp.types import TextContent, Tool

if TYPE_CHECKING:
    pass


# ── Hằng số nội bộ ──────────────────────────────────────────────────────────

_TOOL_LIST_EMAILS = "list_emails"
_TOOL_READ_EMAIL = "read_email"

# Số ký tự tối đa của body trả về cho Claude (tránh flooding context window)
# Config có thể override giá trị này qua config.limits.email_body_max_length
MAX_BODY_CHARS_DEFAULT = 10_000

# Giá trị mặc định khi tham số không được truyền vào
DEFAULT_MAX_COUNT = 20
MAX_COUNT_HARD_CAP = 100  # Không bao giờ vượt quá giới hạn cứng này


# ── Hàm đăng ký tools ───────────────────────────────────────────────────────

def register_tools(server: Server, outlook_com: Any, audit: Any, config: Any) -> None:
    """
    Đăng ký 2 MCP tools đọc email vào server.

    Tham số:
        server      -- MCP Server instance
        outlook_com -- OutlookCOMBridge instance
        audit       -- AuditLogger instance
        config      -- Config object đã validate
    """

    @server.list_tools()
    async def list_tools_handler() -> list[Tool]:
        """Khai báo danh sách tools cho MCP protocol."""
        return [
            Tool(
                name=_TOOL_LIST_EMAILS,
                description=(
                    "Liệt kê danh sách email trong một thư mục Outlook. "
                    "Trả về tóm tắt: tiêu đề, người gửi, thời gian nhận, "
                    "có đính kèm hay không. Không trả về nội dung email."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "folder_name": {
                            "type": "string",
                            "maxLength": 260,
                            "description": "Tên thư mục cần liệt kê (phải trong allowlist)",
                        },
                        "max_count": {
                            "type": "integer",
                            "default": DEFAULT_MAX_COUNT,
                            "maximum": MAX_COUNT_HARD_CAP,
                            "description": f"Số email tối đa trả về (mặc định {DEFAULT_MAX_COUNT}, tối đa {MAX_COUNT_HARD_CAP})",
                        },
                        "since_date": {
                            "type": "string",
                            "format": "date",
                            "description": "Chỉ lấy email từ ngày này (định dạng YYYY-MM-DD)",
                        },
                    },
                    "required": ["folder_name"],
                    "additionalProperties": False,
                },
            ),
            Tool(
                name=_TOOL_READ_EMAIL,
                description=(
                    "Đọc nội dung chi tiết một email theo entry_id. "
                    "Trả về: tiêu đề, người gửi, người nhận, nội dung text (đã strip HTML), "
                    "danh sách file đính kèm."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "entry_id": {
                            "type": "string",
                            "pattern": "^[0-9A-Fa-f]+$",
                            "maxLength": 256,
                            "description": "Entry ID của email (chuỗi hex, lấy từ kết quả list_emails)",
                        }
                    },
                    "required": ["entry_id"],
                    "additionalProperties": False,
                },
            ),
        ]

    @server.call_tool()
    async def call_tool_handler(name: str, arguments: dict) -> list[TextContent]:
        """
        Điều phối lời gọi tool đến đúng hàm xử lý.

        Chỉ xử lý tool của module này; trả về danh sách rỗng cho các tool khác
        để server có thể chuyển tiếp.
        """
        if name == _TOOL_LIST_EMAILS:
            result = await _handle_list_emails(outlook_com, audit, config, arguments)
        elif name == _TOOL_READ_EMAIL:
            result = await _handle_read_email(outlook_com, audit, config, arguments)
        else:
            return []

        return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False))]


# ── Hàm xử lý nội bộ ────────────────────────────────────────────────────────

async def _handle_list_emails(
    outlook_com: Any,
    audit: Any,
    config: Any,
    arguments: dict,
) -> dict:
    """
    Xử lý tool list_emails.

    Bước 1: Validate folder_name — phải trong allowlist
    Bước 2: Validate và cap max_count không vượt giới hạn cứng
    Bước 3: Validate since_date nếu có
    Bước 4: Gọi COM bridge lấy danh sách email
    Bước 5: Ghi audit log (không ghi subject, sender)
    Bước 6: Trả về danh sách EmailSummary dạng dict

    Trả về list với mỗi phần tử: {entry_id, subject, sender_name, sender_email,
                                   received_time, has_attachment, is_read, size_kb}
    """
    start_ms = int(time.monotonic() * 1000)

    # Bước 1: Kiểm tra tham số bắt buộc
    folder_name: str = arguments.get("folder_name", "").strip()
    if not folder_name:
        return {"error": "Tham số folder_name không được để trống"}

    # Bước 2: Validate và giới hạn max_count
    max_count_raw = arguments.get("max_count", DEFAULT_MAX_COUNT)
    try:
        max_count = int(max_count_raw)
    except (TypeError, ValueError):
        max_count = DEFAULT_MAX_COUNT

    # Cap cứng theo config và giới hạn module
    config_max = getattr(getattr(config, "limits", None), "list_emails_max_limit", MAX_COUNT_HARD_CAP)
    max_count = min(max_count, config_max, MAX_COUNT_HARD_CAP)
    max_count = max(1, max_count)  # Phải lớn hơn 0

    # Bước 3: Xử lý tham số since_date (tùy chọn)
    since_date: str | None = arguments.get("since_date")
    if since_date:
        since_date = since_date.strip()
        # Validate định dạng YYYY-MM-DD đơn giản trước khi truyền vào COM
        if not _is_valid_date_format(since_date):
            return {"error": "Định dạng since_date không hợp lệ. Dùng định dạng YYYY-MM-DD"}

    try:
        # Bước 1b: Validate thư mục qua allowlist
        from security.validator import InputValidator
        validator = InputValidator(config)
        validated_folder = validator.validate_folder_name(folder_name)

        # Bước 4: Gọi COM bridge lấy danh sách email
        emails_raw = outlook_com.list_emails(
            folder_name=validated_folder,
            max_count=max_count,
            since_date=since_date,
        )

        # Chuyển đổi kết quả COM về danh sách dict an toàn
        emails: list[dict] = [_sanitize_email_summary(e) for e in (emails_raw or [])]

        elapsed_ms = int(time.monotonic() * 1000) - start_ms

        # Bước 5: Ghi audit log — KHÔNG ghi subject hay sender_email
        audit.log(
            tool=_TOOL_LIST_EMAILS,
            action="list_emails",
            params={
                "folder_name": validated_folder,
                "max_count": max_count,
                "has_since_date": since_date is not None,
            },
            result={
                "status": "ok",
                "items_returned": len(emails),
                "duration_ms": elapsed_ms,
            },
        )

        # Bước 6: Trả kết quả
        return {"emails": emails, "folder": validated_folder, "count": len(emails)}

    except ValueError as exc:
        # Thư mục không trong allowlist
        elapsed_ms = int(time.monotonic() * 1000) - start_ms
        audit.log(
            tool=_TOOL_LIST_EMAILS,
            action="list_emails",
            params={"folder_name": folder_name},
            result={
                "status": "blocked",
                "block_reason": "not_in_allowlist",
                "duration_ms": elapsed_ms,
            },
        )
        return {"error": f"Thư mục không được phép truy cập: {exc}"}

    except Exception:
        elapsed_ms = int(time.monotonic() * 1000) - start_ms
        audit.log(
            tool=_TOOL_LIST_EMAILS,
            action="list_emails",
            params={"folder_name": folder_name},
            result={"status": "error", "duration_ms": elapsed_ms},
        )
        return {"error": "Không thể lấy danh sách email. Đảm bảo Outlook đang chạy."}


async def _handle_read_email(
    outlook_com: Any,
    audit: Any,
    config: Any,
    arguments: dict,
) -> dict:
    """
    Xử lý tool read_email.

    Bước 1: Validate entry_id — chỉ chấp nhận hex, max 256 ký tự, không null bytes
    Bước 2: Gọi COM bridge lấy email theo entry_id
    Bước 3: Verify thư mục chứa email thuộc allowlist (chống TOCTOU)
    Bước 4: Strip HTML body thành plain text
    Bước 5: Truncate body theo MAX_BODY_CHARS
    Bước 6: Ghi audit log — chỉ ghi 8 ký tự đầu của entry_id
    Bước 7: Trả về EmailDetail dict

    Trả về: {subject, sender_name, sender_email, to_recipients, cc_recipients,
              received_time, body_text, attachments: [{name, size_kb, extension}]}
    """
    start_ms = int(time.monotonic() * 1000)

    # Bước 1: Validate entry_id
    entry_id_raw: str = arguments.get("entry_id", "")
    if not entry_id_raw:
        return {"error": "Tham số entry_id không được để trống"}

    try:
        from security.validator import InputValidator
        validator = InputValidator(config)
        entry_id = validator.validate_email_id(entry_id_raw)
    except ValueError as exc:
        audit.log(
            tool=_TOOL_READ_EMAIL,
            action="read_email",
            params={"entry_id_prefix": entry_id_raw[:8] if len(entry_id_raw) >= 8 else "???"},
            result={"status": "blocked", "block_reason": "invalid_entry_id"},
        )
        return {"error": f"entry_id không hợp lệ: {exc}"}

    # 8 ký tự đầu của entry_id dùng cho audit (không dùng toàn bộ để tránh expose)
    entry_id_prefix = entry_id[:8]

    try:
        # Bước 2: Lấy email qua COM bridge
        email_raw = outlook_com.read_email(entry_id)
        if email_raw is None:
            elapsed_ms = int(time.monotonic() * 1000) - start_ms
            audit.log(
                tool=_TOOL_READ_EMAIL,
                action="read_email",
                params={"entry_id_prefix": entry_id_prefix},
                result={"status": "not_found", "duration_ms": elapsed_ms},
            )
            return {"error": "Không tìm thấy email với entry_id đã cho"}

        # Bước 3: Verify thư mục chứa email phải trong allowlist
        # SECURITY: Áp dụng fail-closed — nếu không đọc được folder_name thì từ chối luôn,
        # không được bỏ qua kiểm tra (tránh lỗ hổng fail-open khi PST store hoặc COM lỗi)
        email_folder = email_raw.get("folder_name", "")
        if not email_folder:
            elapsed_ms = int(time.monotonic() * 1000) - start_ms
            audit.log(
                tool=_TOOL_READ_EMAIL,
                action="read_email",
                params={"entry_id_prefix": entry_id_prefix},
                result={
                    "status": "blocked",
                    "block_reason": "folder_name_missing_or_empty",
                    "duration_ms": elapsed_ms,
                },
            )
            return {"error": "Không thể xác minh thư mục chứa email. Từ chối truy cập."}
        from security.validator import InputValidator as _V
        try:
            _V(config).validate_folder_name(email_folder)
        except ValueError:
            elapsed_ms = int(time.monotonic() * 1000) - start_ms
            audit.log(
                tool=_TOOL_READ_EMAIL,
                action="read_email",
                params={"entry_id_prefix": entry_id_prefix},
                result={
                    "status": "blocked",
                    "block_reason": "folder_not_in_allowlist",
                    "duration_ms": elapsed_ms,
                },
            )
            return {"error": "Email nằm trong thư mục không được phép truy cập"}

        # Bước 4: Strip HTML body thành plain text
        body_raw: str = email_raw.get("body_html") or email_raw.get("body_text") or ""
        body_text = _strip_html_to_text(body_raw)

        # Bước 5: Truncate body theo giới hạn config
        max_body = getattr(
            getattr(config, "limits", None),
            "email_body_max_length",
            MAX_BODY_CHARS_DEFAULT,
        )
        was_truncated = len(body_text) > max_body
        if was_truncated:
            body_text = body_text[:max_body] + f"\n\n[... Nội dung bị cắt bớt — {len(body_text) - max_body} ký tự còn lại ...]"

        # Bước 6: Chuẩn bị kết quả trả về (không bao gồm folder_name nội bộ)
        result = {
            "subject": _sanitize_string(email_raw.get("subject", "")),
            "sender_name": _sanitize_string(email_raw.get("sender_name", "")),
            "sender_email": _sanitize_string(email_raw.get("sender_email", "")),
            "to_recipients": [
                _sanitize_string(r) for r in (email_raw.get("to_recipients") or [])
            ],
            "cc_recipients": [
                _sanitize_string(r) for r in (email_raw.get("cc_recipients") or [])
            ],
            "received_time": email_raw.get("received_time", ""),
            "body_text": body_text,
            "body_truncated": was_truncated,
            "attachments": _sanitize_attachments(email_raw.get("attachments") or []),
        }

        elapsed_ms = int(time.monotonic() * 1000) - start_ms

        # Bước 6 (audit): Chỉ ghi 8 ký tự đầu entry_id — TUYỆT ĐỐI không ghi body/subject/sender
        audit.log(
            tool=_TOOL_READ_EMAIL,
            action="read_email",
            params={"entry_id_prefix": entry_id_prefix},
            result={
                "status": "ok",
                "has_attachments": len(result["attachments"]) > 0,
                "body_truncated": was_truncated,
                "duration_ms": elapsed_ms,
            },
        )

        return result

    except Exception:
        elapsed_ms = int(time.monotonic() * 1000) - start_ms
        audit.log(
            tool=_TOOL_READ_EMAIL,
            action="read_email",
            params={"entry_id_prefix": entry_id_prefix},
            result={"status": "error", "duration_ms": elapsed_ms},
        )
        return {"error": "Không thể đọc email. Đảm bảo Outlook đang chạy."}


# ── Hàm tiện ích nội bộ ──────────────────────────────────────────────────────

def _sanitize_email_summary(raw: dict) -> dict:
    """
    Làm sạch và chuẩn hóa một EmailSummary từ COM bridge.

    Loại bỏ các trường nội bộ không cần thiết, sanitize string fields
    để tránh prompt injection qua subject hoặc sender.

    Tham số:
        raw -- dict thô từ COM bridge

    Trả về:
        dict đã được sanitize với các trường: entry_id, subject, sender_name,
        sender_email, received_time, has_attachment, is_read, size_kb
    """
    return {
        "entry_id": str(raw.get("entry_id", "")),
        "subject": _sanitize_string(raw.get("subject", "")),
        "sender_name": _sanitize_string(raw.get("sender_name", "")),
        "sender_email": _sanitize_string(raw.get("sender_email", "")),
        "received_time": str(raw.get("received_time", "")),
        "has_attachment": bool(raw.get("has_attachment", False)),
        "is_read": bool(raw.get("is_read", True)),
        "size_kb": int(raw.get("size_kb", 0)),
    }


def _sanitize_attachments(attachments: list) -> list[dict]:
    """
    Làm sạch danh sách file đính kèm.

    Escape tên file đính kèm vì tên file có thể chứa injection.
    Không trả về đường dẫn file — chỉ trả về tên hiển thị, kích thước, phần mở rộng.

    Tham số:
        attachments -- Danh sách thô từ COM bridge

    Trả về:
        Danh sách dict với {name, size_kb, extension}
    """
    result = []
    for att in attachments:
        if not isinstance(att, dict):
            continue
        name = _sanitize_string(str(att.get("name", "")))
        size_kb = int(att.get("size_kb", 0))
        # Lấy phần mở rộng từ tên file
        extension = ""
        if "." in name:
            extension = name.rsplit(".", 1)[-1].lower()[:10]  # Giới hạn 10 ký tự
        result.append({"name": name, "size_kb": size_kb, "extension": extension})
    return result


def _strip_html_to_text(html: str) -> str:
    """
    Strip HTML thành plain text để tránh:
      1. HTML comment injection (<!-- SYSTEM: admin mode -->)
      2. Invisible Unicode / hidden content
      3. Script injection trong body

    Dùng BeautifulSoup nếu có. Nếu không, dùng regex đơn giản làm fallback.

    Tham số:
        html -- Chuỗi HTML hoặc plain text từ Outlook

    Trả về:
        Chuỗi plain text đã được làm sạch
    """
    if not html:
        return ""

    try:
        from bs4 import BeautifulSoup
        # Dùng html.parser built-in, không cần lxml
        soup = BeautifulSoup(html, "html.parser")
        # Xóa toàn bộ thẻ script và style trước khi lấy text
        for tag in soup(["script", "style", "head"]):
            tag.decompose()
        text = soup.get_text(separator="\n")
    except ImportError:
        # Fallback: dùng regex đơn giản nếu BeautifulSoup chưa cài
        import re
        text = re.sub(r"<[^>]+>", "", html)
        text = re.sub(r"<!--.*?-->", "", text, flags=re.DOTALL)

    # Sanitize: loại bỏ invisible Unicode và control chars nguy hiểm
    return _sanitize_string(text)


def _sanitize_string(value: str) -> str:
    """
    Làm sạch chuỗi string để tránh prompt injection và các ký tự nguy hiểm.

    Loại bỏ:
      - Null bytes (\\x00) — có thể gây lỗi COM hoặc bypass validation
      - Invisible Unicode: U+200B (zero-width space), U+202E (right-to-left override)
      - Control characters \\x01-\\x1F ngoại trừ newline (\\x0A) và tab (\\x09)

    Tham số:
        value -- Chuỗi cần làm sạch

    Trả về:
        Chuỗi đã được loại bỏ ký tự nguy hiểm
    """
    if not isinstance(value, str):
        return str(value) if value is not None else ""

    # Loại bỏ null bytes
    value = value.replace("\x00", "")

    # Loại bỏ invisible Unicode phổ biến dùng để inject
    invisible_chars = [
        "​",  # Zero-width space
        "‌",  # Zero-width non-joiner
        "‍",  # Zero-width joiner
        "‮",  # Right-to-left override (nguy hiểm nhất)
        " ",  # Line separator
        " ",  # Paragraph separator
        "﻿",  # BOM
    ]
    for char in invisible_chars:
        value = value.replace(char, "")

    # Loại bỏ control chars (giữ lại newline \x0A=10 và tab \x09=9)
    result_chars = []
    for char in value:
        code = ord(char)
        if code < 0x20 and code not in (0x09, 0x0A, 0x0D):
            continue  # Bỏ qua control char
        result_chars.append(char)

    return "".join(result_chars)


def _is_valid_date_format(date_str: str) -> bool:
    """
    Kiểm tra chuỗi có đúng định dạng YYYY-MM-DD không.

    Tham số:
        date_str -- Chuỗi ngày tháng cần kiểm tra

    Trả về:
        True nếu hợp lệ, False nếu không
    """
    import re
    return bool(re.fullmatch(r"\d{4}-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12]\d|3[01])", date_str))


# ── Hàm dispatch đồng bộ cho server.py ──────────────────────────────────────

def handle_list_emails(arguments: dict, config, com_bridge) -> dict:
    """
    Wrapper đồng bộ liệt kê email cho server.py dispatch.

    Chạy trong STA thread executor của server.py.

    Tham số:
        arguments  -- dict tham số từ Claude
        config     -- Config object (có ALLOWED_FOLDERS, MAX_EMAILS_PER_REQUEST)
        com_bridge -- OutlookCOMBridge instance

    Trả về:
        dict {"emails": [...], "folder": str, "count": int}
        hoặc {"error": str}
    """
    from security.validator import InputValidator

    # Lấy và chuẩn hóa tham số
    folder_name = (arguments.get("folder_name") or "Inbox").strip() or "Inbox"
    max_count_raw = arguments.get("max_count", DEFAULT_MAX_COUNT)
    try:
        max_count = int(max_count_raw)
    except (TypeError, ValueError):
        max_count = DEFAULT_MAX_COUNT
    max_count = min(max(1, max_count), MAX_COUNT_HARD_CAP)

    since_date = arguments.get("since_date")
    if since_date:
        since_date = since_date.strip()
        if not _is_valid_date_format(since_date):
            return {"error": "Định dạng since_date không hợp lệ. Dùng định dạng YYYY-MM-DD"}

    unread_only = bool(arguments.get("unread_only", False))

    # Validate folder qua allowlist
    try:
        validated_folder = InputValidator(config).validate_folder_name(folder_name)
    except Exception as exc:
        return {"error": f"Thư mục không được phép truy cập: {exc}"}

    try:
        emails = com_bridge.list_emails(
            folder_name=validated_folder,
            max_count=max_count,
            allowed_folders=list(getattr(config, "ALLOWED_FOLDERS", []) or []),
            since_date=since_date,
            unread_only=unread_only,
        )
        return {"emails": emails, "folder": validated_folder, "count": len(emails)}
    except Exception as exc:
        return {"error": f"Không thể lấy danh sách email. Đảm bảo Outlook đang chạy. Chi tiết: {type(exc).__name__}"}


def handle_read_email(arguments: dict, config, com_bridge) -> dict:
    """
    Wrapper đồng bộ đọc email cho server.py dispatch.

    Bảo mật:
    - Validate entry_id trước khi gọi COM
    - Xác minh email nằm trong allowed_folders (fail-closed)
    - Truncate body theo MAX_BODY_CHARS

    Tham số:
        arguments  -- dict tham số từ Claude
        config     -- Config object
        com_bridge -- OutlookCOMBridge instance

    Trả về:
        dict thông tin email hoặc {"error": str}
    """
    from security.validator import InputValidator

    entry_id_raw = (arguments.get("entry_id") or "").strip()
    if not entry_id_raw:
        return {"error": "Tham số entry_id không được để trống"}

    # Validate entry_id — chỉ hex, max 256 ký tự, không null bytes
    try:
        entry_id = InputValidator(config).validate_email_id(entry_id_raw)
    except Exception as exc:
        return {"error": f"entry_id không hợp lệ: {exc}"}

    allowed = list(getattr(config, "ALLOWED_FOLDERS", []) or [])

    try:
        email_dict = com_bridge.read_email(entry_id=entry_id, allowed_folders=allowed)
    except Exception as exc:
        return {"error": f"Không thể đọc email. Đảm bảo Outlook đang chạy. Chi tiết: {type(exc).__name__}"}

    if email_dict is None:
        return {"error": "Không tìm thấy email với entry_id đã cho"}

    # Kiểm tra bảo mật folder — fail-closed
    email_folder = email_dict.get("folder_name", "")
    if not email_folder:
        return {"error": "Không thể xác minh thư mục chứa email. Từ chối truy cập."}
    try:
        InputValidator(config).validate_folder_name(email_folder)
    except Exception:
        return {"error": "Email nằm ngoài thư mục được phép truy cập."}

    # Dọn format body: chuẩn hóa line endings, bỏ blank lines thừa
    body_raw = email_dict.get("body_text", "") or ""
    if body_raw:
        # Chuẩn hóa \r\n → \n
        body_raw = body_raw.replace("\r\n", "\n").replace("\r", "\n")
        # Bỏ khoảng trắng cuối mỗi dòng
        lines = [line.rstrip() for line in body_raw.split("\n")]
        # Collapse nhiều blank lines liên tiếp thành tối đa 1 blank line
        cleaned: list[str] = []
        prev_blank = False
        for line in lines:
            is_blank = (line == "")
            if is_blank and prev_blank:
                continue  # Bỏ qua blank line liên tiếp thứ 2 trở đi
            cleaned.append(line)
            prev_blank = is_blank
        body_raw = "\n".join(cleaned).strip()

    # Truncate body nếu quá dài
    max_body = getattr(config, "MAX_BODY_CHARS", MAX_BODY_CHARS_DEFAULT)
    if body_raw and len(body_raw) > max_body:
        body_raw = body_raw[:max_body] + "\n...[nội dung bị cắt bớt theo giới hạn bảo mật]"

    if body_raw != (email_dict.get("body_text", "") or ""):
        email_dict = dict(email_dict)  # không mutate dict gốc
        email_dict["body_text"] = body_raw

    return email_dict
