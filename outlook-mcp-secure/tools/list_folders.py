"""
Module list_folders — Các MCP tools quản lý và hiển thị thư mục Outlook.

Cung cấp 2 tools:
  - list_allowed_folders: Liệt kê tất cả thư mục nằm trong allowlist
  - get_folder_stats: Thống kê số lượng email trong một thư mục cụ thể

Bảo mật:
  - Chỉ trả về thư mục đã cấu hình trong config.security.allowed_folders
  - Không expose đường dẫn PST file
  - Audit log mỗi lần gọi, chỉ ghi tên thư mục được phép (không ghi thư mục bị từ chối)
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

# Tên tool MCP theo đúng design spec
_TOOL_LIST_FOLDERS = "list_allowed_folders"
_TOOL_FOLDER_STATS = "get_folder_stats"


# ── Hàm đăng ký tools ───────────────────────────────────────────────────────

def register_tools(server: Server, outlook_com: Any, audit: Any, config: Any) -> None:
    """
    Đăng ký 2 MCP tools liên quan đến thư mục vào server.

    Tham số:
        server      -- MCP Server instance
        outlook_com -- OutlookCOMBridge instance
        audit       -- AuditLogger instance
        config      -- Config object đã validate
    """

    # ── Tool 1: list_allowed_folders ─────────────────────────────────────────

    @server.list_tools()
    async def list_tools_handler() -> list[Tool]:
        """Khai báo danh sách tools cho MCP protocol."""
        return [
            Tool(
                name=_TOOL_LIST_FOLDERS,
                description=(
                    "Liệt kê tất cả thư mục email Outlook được phép truy cập. "
                    "Chỉ trả về các thư mục đã cấu hình trong allowlist."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "include_subfolders": {
                            "type": "boolean",
                            "default": False,
                            "description": "Có bao gồm thư mục con hay không",
                        }
                    },
                    "required": [],
                    "additionalProperties": False,
                },
            ),
            Tool(
                name=_TOOL_FOLDER_STATS,
                description=(
                    "Thống kê số lượng email trong một thư mục cụ thể: "
                    "tổng số email và số email chưa đọc."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "folder_name": {
                            "type": "string",
                            "maxLength": 260,
                            "description": "Tên thư mục cần thống kê (phải nằm trong allowlist)",
                        }
                    },
                    "required": ["folder_name"],
                    "additionalProperties": False,
                },
            ),
        ]

    # ── Tool handler: call_tool ──────────────────────────────────────────────

    @server.call_tool()
    async def call_tool_handler(name: str, arguments: dict) -> list[TextContent]:
        """
        Điều phối lời gọi tool đến đúng hàm xử lý.

        Chỉ xử lý các tool thuộc module này; bỏ qua các tool khác.
        """
        if name == _TOOL_LIST_FOLDERS:
            result = await _handle_list_allowed_folders(outlook_com, audit, config, arguments)
        elif name == _TOOL_FOLDER_STATS:
            result = await _handle_get_folder_stats(outlook_com, audit, config, arguments)
        else:
            # Không phải tool của module này — để server xử lý tiếp
            return []

        return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False))]


# ── Hàm xử lý nội bộ ────────────────────────────────────────────────────────

async def _handle_list_allowed_folders(
    outlook_com: Any,
    audit: Any,
    config: Any,
    arguments: dict,
) -> dict:
    """
    Xử lý tool list_allowed_folders.

    Bước 1: Validate tham số đầu vào
    Bước 2: Lấy danh sách thư mục từ allowlist trong config
    Bước 3: Với mỗi thư mục trong allowlist, thử lấy thống kê qua COM
    Bước 4: Ghi audit log (chỉ ghi số lượng, không ghi tên thư mục ngoài allowlist)
    Bước 5: Trả về kết quả dạng dict

    Tham số:
        include_subfolders (bool) -- Có bao gồm thư mục con không (mặc định False)
    """
    start_ms = int(time.monotonic() * 1000)

    # Bước 1: Lấy tham số, dùng giá trị mặc định nếu không có
    include_subfolders: bool = bool(arguments.get("include_subfolders", False))

    folders_result: list[dict] = []

    try:
        # Bước 2: Lấy danh sách thư mục được phép từ config
        # Hỗ trợ cả ALLOWED_FOLDERS (Config dataclass) và config.security.allowed_folders (cũ)
        allowed_folders: list[str] = list(
            getattr(config, "ALLOWED_FOLDERS", None)
            or getattr(getattr(config, "security", None), "allowed_folders", None)
            or []
        )

        # Bước 3: Với mỗi thư mục trong allowlist, lấy thống kê qua COM bridge
        for folder_name in allowed_folders:
            folder_info = _get_folder_info_safe(
                outlook_com=outlook_com,
                folder_name=folder_name,
                include_subfolders=include_subfolders,
            )
            if folder_info is not None:
                folders_result.append(folder_info)

        # Bước 4: Ghi audit log — chỉ ghi số lượng folder trả về, không ghi tên
        elapsed_ms = int(time.monotonic() * 1000) - start_ms
        audit.log(
            tool=_TOOL_LIST_FOLDERS,
            action="list_allowed_folders",
            params={"include_subfolders": include_subfolders},
            result={
                "status": "ok",
                "folder_count": len(folders_result),
                "duration_ms": elapsed_ms,
            },
        )

        # Bước 5: Trả kết quả — thêm ghi chú bảo mật cho người dùng
        return {
            "folders": folders_result,
            "note": "Chỉ hiển thị các thư mục đã được cấu hình trong allowlist",
        }

    except Exception:
        # Ghi log lỗi nội bộ, không expose stack trace ra ngoài
        elapsed_ms = int(time.monotonic() * 1000) - start_ms
        audit.log(
            tool=_TOOL_LIST_FOLDERS,
            action="list_allowed_folders",
            params={"include_subfolders": include_subfolders},
            result={"status": "error", "duration_ms": elapsed_ms},
        )
        return {"error": "Không thể lấy danh sách thư mục. Đảm bảo Outlook đang chạy."}


async def _handle_get_folder_stats(
    outlook_com: Any,
    audit: Any,
    config: Any,
    arguments: dict,
) -> dict:
    """
    Xử lý tool get_folder_stats.

    Bước 1: Validate folder_name qua InputValidator (allowlist check)
    Bước 2: Gọi COM bridge lấy thống kê thư mục
    Bước 3: Ghi audit log
    Bước 4: Trả về {"folder", "total", "unread"}

    Tham số:
        folder_name (str) -- Tên thư mục cần thống kê
    """
    start_ms = int(time.monotonic() * 1000)

    # Bước 1: Kiểm tra tham số bắt buộc
    folder_name: str = arguments.get("folder_name", "").strip()
    if not folder_name:
        return {"error": "Tham số folder_name không được để trống"}

    try:
        # Bước 1b: Validate thư mục có nằm trong allowlist không
        from security.validator import InputValidator
        validator = InputValidator(config)
        validated_folder = validator.validate_folder_name(folder_name)

        # Bước 2: Lấy thống kê thư mục qua COM bridge
        stats = outlook_com.get_folder_stats(validated_folder)

        elapsed_ms = int(time.monotonic() * 1000) - start_ms

        # Bước 3: Ghi audit log — ghi tên thư mục vì đã qua allowlist validation
        audit.log(
            tool=_TOOL_FOLDER_STATS,
            action="get_folder_stats",
            params={"folder_name": validated_folder},
            result={
                "status": "ok",
                "total": stats.get("total", 0),
                "unread": stats.get("unread", 0),
                "duration_ms": elapsed_ms,
            },
        )

        # Bước 4: Trả kết quả
        return {
            "folder": validated_folder,
            "total": stats.get("total", 0),
            "unread": stats.get("unread", 0),
        }

    except ValueError as exc:
        # Lỗi validation — thư mục không trong allowlist
        elapsed_ms = int(time.monotonic() * 1000) - start_ms
        audit.log(
            tool=_TOOL_FOLDER_STATS,
            action="get_folder_stats",
            params={"folder_name": folder_name},
            result={
                "status": "blocked",
                "block_reason": "not_in_allowlist",
                "duration_ms": elapsed_ms,
            },
        )
        return {"error": f"Thư mục không được phép truy cập: {exc}"}

    except Exception:
        # Lỗi không xác định — không expose chi tiết ra ngoài
        elapsed_ms = int(time.monotonic() * 1000) - start_ms
        audit.log(
            tool=_TOOL_FOLDER_STATS,
            action="get_folder_stats",
            params={"folder_name": folder_name},
            result={"status": "error", "duration_ms": elapsed_ms},
        )
        return {"error": "Không thể lấy thống kê thư mục. Đảm bảo Outlook đang chạy."}


def _get_folder_info_safe(
    outlook_com: Any,
    folder_name: str,
    include_subfolders: bool,
) -> dict | None:
    """
    Lấy thông tin một thư mục qua COM bridge một cách an toàn.

    Nếu thư mục không tồn tại hoặc COM lỗi — trả về None để bỏ qua,
    không làm crash toàn bộ list.

    Tham số:
        outlook_com       -- COM bridge instance
        folder_name       -- Tên thư mục cần lấy thông tin
        include_subfolders -- Có đệ quy thư mục con không

    Trả về:
        dict với {name, path, unread_count, total_count} hoặc None nếu lỗi
    """
    try:
        stats = outlook_com.get_folder_stats(folder_name)
        result: dict = {
            "name": folder_name,
            # Không expose đường dẫn PST file — chỉ dùng tên hiển thị
            "path": folder_name,
            "unread_count": stats.get("unread", 0),
            "total_count": stats.get("total", 0),
        }

        # Nếu yêu cầu bao gồm thư mục con, lấy thêm danh sách subfolder
        if include_subfolders:
            subfolders = outlook_com.get_subfolders(folder_name)
            if subfolders:
                result["subfolders"] = subfolders

        return result

    except Exception:
        # Không raise — bỏ qua thư mục này, tiếp tục với thư mục khác
        return None


# ── Hàm dispatch đồng bộ cho server.py ──────────────────────────────────────

def handle_list_folders(arguments: dict, config, com_bridge) -> dict:
    """
    Wrapper đồng bộ cho server.py dispatch (thay thế register_tools() cũ).

    Hàm này chạy trong STA thread executor của server.py — đã được CoInitialize().
    KHÔNG dùng async/await ở đây.

    Tham số:
        arguments  -- dict tham số từ Claude (đã qua JSON Schema validation ở server.py)
        config     -- Config object (có ALLOWED_FOLDERS, ...)
        com_bridge -- OutlookCOMBridge instance

    Trả về:
        dict {"folders": [...], "note": str}
        hoặc {"error": str} nếu thất bại
    """
    include_subfolders = bool(arguments.get("include_subfolders", False))

    # Lấy danh sách thư mục được phép — hỗ trợ cả hai dạng config
    allowed_folders: list[str] = list(
        getattr(config, "ALLOWED_FOLDERS", None)
        or getattr(getattr(config, "security", None), "allowed_folders", None)
        or []
    )

    if not allowed_folders:
        return {"error": "Chưa cấu hình danh sách thư mục được phép (allowed_folders)."}

    # Bước bổ sung: Validate từng folder name trong config qua InputValidator
    # Mục đích: lọc bỏ các tên thư mục không hợp lệ trong config.toml (phòng trường hợp
    # config bị sửa tay và vô tình chèn tên thư mục chứa ký tự nguy hiểm như path traversal,
    # null byte, hay URL scheme — những thứ này sẽ bị reject bởi InputValidator.validate_folder).
    # Folder name không hợp lệ sẽ bị bỏ qua (ghi log warning), không làm crash toàn bộ list.
    from security.validator import InputValidator, ValidationError as _ValidationError
    _validator = InputValidator(config)
    sanitized_folders: list[str] = []
    for _name in allowed_folders:
        try:
            # validate_folder_name kiểm tra: null bytes, path traversal, URL scheme,
            # độ dài, và xác nhận tên nằm trong allowlist của chính config
            sanitized_folders.append(_validator.validate_folder_name(_name))
        except _ValidationError:
            # Ghi warning nhưng bỏ qua folder không hợp lệ — không crash list
            import logging as _logging
            _logging.getLogger(__name__).warning(
                "handle_list_folders: bỏ qua folder name không hợp lệ trong config "
                "(độ dài=%d). Kiểm tra lại config.toml.",
                len(_name) if isinstance(_name, str) else 0,
            )
    allowed_folders = sanitized_folders

    folders_result: list[dict] = []
    for folder_name in allowed_folders:
        folder_info = _get_folder_info_safe(
            outlook_com=com_bridge,
            folder_name=folder_name,
            include_subfolders=include_subfolders,
        )
        if folder_info is not None:
            folders_result.append(folder_info)

    return {
        "folders": folders_result,
        "note": "Chỉ hiển thị các thư mục đã được cấu hình trong allowlist",
    }


def handle_list_all_folders(arguments: dict, config, com_bridge) -> dict:
    """
    Liệt kê TẤT CẢ thư mục trong tất cả Store Outlook (không giới hạn allowlist).

    Dùng để khám phá cấu trúc thư mục thực tế trong Outlook.
    Kết quả trả về tất cả folder, kể cả folder nằm ngoài allowlist.

    Chạy trong STA thread executor của server.py — đã CoInitialize().

    Tham số arguments:
        (không có tham số)

    Trả về:
        dict {"folders": [...], "total_folders": int}
        Mỗi folder: {"name": str, "path": str, "store": str, "total": int, "unread": int}
    """
    try:
        folders = com_bridge.get_all_folders_recursive()
        return {
            "folders": folders,
            "total_folders": len(folders),
            "note": (
                "Đây là TOÀN BỘ thư mục trong Outlook, bao gồm cả các thư mục "
                "nằm ngoài danh sách allowed_folders. "
                "Dùng list_folders để xem danh sách thư mục có thể truy cập email."
            ),
        }
    except Exception as exc:
        import logging as _logging
        _logging.getLogger(__name__).warning("handle_list_all_folders thất bại: %s", exc)
        return {"error": "Không thể lấy danh sách thư mục. Đảm bảo Outlook đang chạy."}


def handle_email_stats(arguments: dict, config, com_bridge) -> dict:
    """
    Thống kê số lượng email (tổng số và chưa đọc) trên tất cả thư mục trong allowlist.

    Chạy trong STA thread executor của server.py — đã CoInitialize().

    Tham số arguments:
        (không có tham số)

    Trả về:
        dict {"summary": {...}, "folders": [...]}
        summary: tổng email, tổng chưa đọc, tổng thư mục
        folders: thống kê từng folder: name, total, unread
    """
    # Bước 1: Lấy danh sách folder được phép
    allowed_folders = list(
        getattr(config, "ALLOWED_FOLDERS", None)
        or getattr(getattr(config, "security", None), "allowed_folders", None)
        or []
    )
    if not allowed_folders:
        return {"error": "Chưa cấu hình danh sách thư mục được phép."}

    # Bước 2: Thu thập thống kê từng folder
    folder_stats = []
    total_all = 0
    unread_all = 0
    errors = 0

    for folder_name in allowed_folders:
        try:
            stats = com_bridge.get_folder_stats(folder_name)
            total = stats.get("total", 0)
            unread = stats.get("unread", 0)
            folder_stats.append({
                "name": folder_name,
                "total": total,
                "unread": unread,
            })
            total_all += total
            unread_all += unread
        except Exception:
            errors += 1
            folder_stats.append({
                "name": folder_name,
                "total": None,
                "unread": None,
                "error": "Không thể truy cập",
            })

    return {
        "summary": {
            "total_emails": total_all,
            "total_unread": unread_all,
            "total_folders": len(allowed_folders),
            "accessible_folders": len(allowed_folders) - errors,
        },
        "folders": folder_stats,
    }


def handle_get_contact_stats(arguments: dict, config, com_bridge) -> dict:
    """
    Thống kê lịch sử email với một contact cụ thể trên tất cả allowed folders.

    Tham số arguments:
        email (str): Địa chỉ email của contact cần thống kê

    Trả về:
        dict với tổng số email nhận từ và gửi đến contact, breakdown theo folder
    """
    import re as _re

    email_raw = str(arguments.get("email", "")).strip().lower()
    if not email_raw:
        return {"error": "Tham số email không được để trống."}

    # Validate email format đơn giản
    if not _re.match(r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$', email_raw):
        return {"error": "Địa chỉ email không hợp lệ."}

    allowed_folders = list(
        getattr(config, "ALLOWED_FOLDERS", None)
        or getattr(getattr(config, "security", None), "allowed_folders", None)
        or []
    )
    if not allowed_folders:
        return {"error": "Chưa cấu hình allowed_folders."}

    total_from = 0
    total_to = 0
    folder_breakdown = []

    safe_email = email_raw.replace("'", "''")

    for folder_name in allowed_folders:
        try:
            # Đếm email nhận từ contact
            from_emails = com_bridge.search_emails(
                query=safe_email,
                folder_name=folder_name,
                max_count=200,
                search_in="sender",
                allowed_folders=allowed_folders,
            )
            count_from = len(from_emails)
            total_from += count_from

            folder_breakdown.append({
                "folder": folder_name,
                "received_from_contact": count_from,
            })
        except Exception:
            folder_breakdown.append({
                "folder": folder_name,
                "received_from_contact": None,
                "error": "Không thể truy cập",
            })

    return {
        "contact": email_raw,
        "total_received_from": total_from,
        "folder_breakdown": folder_breakdown,
        "note": "Thống kê email nhận từ contact này. Gửi đến contact xem trong Sent Items.",
    }
