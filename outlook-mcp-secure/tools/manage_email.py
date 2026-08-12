"""
tools/manage_email.py — Các MCP tool quản lý trạng thái email trong Outlook.

Cung cấp 3 tools:
  - mark_email_read:  Đánh dấu email đã đọc hoặc chưa đọc
  - flag_email:       Đặt hoặc xóa flag theo dõi trên email
  - move_email:       Di chuyển email sang thư mục khác trong allowlist

Bảo mật:
  - Tất cả email phải nằm trong allowed_folders trước khi thay đổi (chống IDOR)
  - Thư mục đích của move_email phải trong allowlist
  - Không bao giờ xóa email — chỉ thay đổi metadata hoặc di chuyển
  - Audit log mỗi thao tác
"""

from __future__ import annotations

import logging
import time
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    pass

_logger = logging.getLogger(__name__)

# Hằng số FlagStatus tương ứng với olFlagStatus enum của Outlook
_FLAG_STATUS_NONE = 0       # Không đánh dấu
_FLAG_STATUS_COMPLETE = 1   # Đã hoàn thành
_FLAG_STATUS_FLAGGED = 2    # Đánh dấu để theo dõi

_FLAG_STATUS_LABELS = {
    _FLAG_STATUS_NONE:     "không flag",
    _FLAG_STATUS_COMPLETE: "đã hoàn thành",
    _FLAG_STATUS_FLAGGED:  "đã đánh dấu theo dõi",
}


def handle_mark_read(arguments: dict, config, com_bridge) -> dict:
    """
    Đánh dấu email đã đọc hoặc chưa đọc trong Outlook.

    Chạy trong STA thread executor của server.py — đã CoInitialize().

    Tham số arguments:
        entry_id (str):  ID hex của email cần thay đổi
        read (bool):     True = đánh dấu đã đọc, False = đánh dấu chưa đọc
                         (mặc định True)

    Trả về:
        dict {"status": "ok", "entry_id": str, "read": bool}
        hoặc {"status": "error", "error": str}
    """
    start_ts = time.monotonic()

    # Bước 1: Lấy và validate tham số
    entry_id_raw = str(arguments.get("entry_id", "")).strip()
    if not entry_id_raw:
        return {"status": "error", "error": "Tham số entry_id không được để trống."}

    read = bool(arguments.get("read", True))

    # Bước 2: Lấy allowed_folders từ config
    allowed_folders = list(
        getattr(config, "ALLOWED_FOLDERS", None)
        or getattr(getattr(config, "security", None), "allowed_folders", None)
        or []
    )
    if not allowed_folders:
        return {"status": "error", "error": "Chưa cấu hình danh sách thư mục được phép."}

    # Bước 3: Kiểm tra read_only_mode
    # F-TOOL-04: fallback là True (fail-safe) thay vì False — nếu config bị lỗi,
    # tốt hơn là block thao tác ghi hơn là vô tình cho phép (ngược với compose.py)
    read_only = bool(
        getattr(config, "READ_ONLY_MODE", None)
        if getattr(config, "READ_ONLY_MODE", None) is not None
        else (
            getattr(getattr(config, "security", None), "read_only_mode", None)
            if getattr(getattr(config, "security", None), "read_only_mode", None) is not None
            else True
        )
    )
    if read_only:
        return {
            "status": "blocked",
            "error": "Server đang ở chế độ chỉ đọc (read_only_mode=True). Không thể thay đổi trạng thái email.",
        }

    # Bước 4: Gọi COM bridge để đánh dấu
    try:
        com_bridge.mark_email_read(
            entry_id=entry_id_raw,
            read=read,
            allowed_folders=allowed_folders,
        )
        duration_ms = int((time.monotonic() - start_ts) * 1000)
        status_label = "đã đọc" if read else "chưa đọc"
        _logger.info(
            "mark_email_read: entry_id_prefix=%s..., read=%s, duration=%dms",
            entry_id_raw[:8], read, duration_ms
        )
        return {
            "status": "ok",
            "entry_id": entry_id_raw,
            "read": read,
            "message": f"Email đã được đánh dấu là {status_label}.",
        }

    except Exception as exc:
        duration_ms = int((time.monotonic() - start_ts) * 1000)
        exc_type = type(exc).__name__
        _logger.warning("mark_email_read thất bại: %s: %s", exc_type, exc)

        # Phân loại lỗi để trả về thông điệp phù hợp
        if "FolderNotAllowed" in exc_type or "NotAllowed" in exc_type:
            return {"status": "error", "error": "Email không nằm trong thư mục được phép."}
        if "InvalidEmailId" in exc_type:
            return {"status": "error", "error": "entry_id không hợp lệ."}
        if "OutlookNotRunning" in exc_type:
            return {"status": "error", "error": "Outlook chưa mở. Vui lòng mở Outlook Desktop trước."}
        return {"status": "error", "error": "Không thể đánh dấu trạng thái email. Kiểm tra Outlook và thử lại."}


def handle_flag_email(arguments: dict, config, com_bridge) -> dict:
    """
    Đặt hoặc xóa flag theo dõi trên email trong Outlook.

    Chạy trong STA thread executor của server.py — đã CoInitialize().

    Tham số arguments:
        entry_id (str):     ID hex của email cần thay đổi
        flag_status (str):  "flagged" | "complete" | "none"
                            (mặc định "flagged" — đánh dấu cần theo dõi)

    Trả về:
        dict {"status": "ok", "entry_id": str, "flag_status": str}
        hoặc {"status": "error", "error": str}
    """
    start_ts = time.monotonic()

    # Bước 1: Lấy và validate tham số
    entry_id_raw = str(arguments.get("entry_id", "")).strip()
    if not entry_id_raw:
        return {"status": "error", "error": "Tham số entry_id không được để trống."}

    # Chuyển đổi tên flag thân thiện sang FlagStatus integer
    flag_name = str(arguments.get("flag_status", "flagged")).lower().strip()
    flag_map = {
        "flagged":  _FLAG_STATUS_FLAGGED,
        "complete": _FLAG_STATUS_COMPLETE,
        "completed": _FLAG_STATUS_COMPLETE,
        "done":     _FLAG_STATUS_COMPLETE,
        "none":     _FLAG_STATUS_NONE,
        "clear":    _FLAG_STATUS_NONE,
        "remove":   _FLAG_STATUS_NONE,
        "unflag":   _FLAG_STATUS_NONE,
    }
    if flag_name not in flag_map:
        valid_names = ", ".join(sorted(set(flag_map.keys())))
        return {
            "status": "error",
            "error": f"flag_status không hợp lệ: '{flag_name}'. Giá trị hợp lệ: {valid_names}",
        }
    flag_status_int = flag_map[flag_name]

    # Bước 2: Lấy allowed_folders và kiểm tra read_only
    allowed_folders = list(
        getattr(config, "ALLOWED_FOLDERS", None)
        or getattr(getattr(config, "security", None), "allowed_folders", None)
        or []
    )
    if not allowed_folders:
        return {"status": "error", "error": "Chưa cấu hình danh sách thư mục được phép."}

    # F-TOOL-04: fallback là True (fail-safe) thay vì False — nếu config bị lỗi,
    # tốt hơn là block thao tác ghi hơn là vô tình cho phép (ngược với compose.py)
    read_only = bool(
        getattr(config, "READ_ONLY_MODE", None)
        if getattr(config, "READ_ONLY_MODE", None) is not None
        else (
            getattr(getattr(config, "security", None), "read_only_mode", None)
            if getattr(getattr(config, "security", None), "read_only_mode", None) is not None
            else True
        )
    )
    if read_only:
        return {
            "status": "blocked",
            "error": "Server đang ở chế độ chỉ đọc (read_only_mode=True). Không thể thay đổi flag email.",
        }

    # Bước 3: Gọi COM bridge để đặt flag
    try:
        com_bridge.flag_email(
            entry_id=entry_id_raw,
            flag_status=flag_status_int,
            allowed_folders=allowed_folders,
        )
        duration_ms = int((time.monotonic() - start_ts) * 1000)
        status_label = _FLAG_STATUS_LABELS.get(flag_status_int, str(flag_status_int))
        _logger.info(
            "flag_email: entry_id_prefix=%s..., flag=%s, duration=%dms",
            entry_id_raw[:8], flag_name, duration_ms
        )
        return {
            "status": "ok",
            "entry_id": entry_id_raw,
            "flag_status": flag_name,
            "message": f"Flag email đã được cập nhật: {status_label}.",
        }

    except Exception as exc:
        duration_ms = int((time.monotonic() - start_ts) * 1000)
        exc_type = type(exc).__name__
        _logger.warning("flag_email thất bại: %s: %s", exc_type, exc)

        if "FolderNotAllowed" in exc_type or "NotAllowed" in exc_type:
            return {"status": "error", "error": "Email không nằm trong thư mục được phép."}
        if "InvalidEmailId" in exc_type:
            return {"status": "error", "error": "entry_id không hợp lệ."}
        if "OutlookNotRunning" in exc_type:
            return {"status": "error", "error": "Outlook chưa mở. Vui lòng mở Outlook Desktop trước."}
        return {"status": "error", "error": "Không thể đặt flag email. Kiểm tra Outlook và thử lại."}


def handle_move_email(arguments: dict, config, com_bridge) -> dict:
    """
    Di chuyển email sang thư mục khác trong danh sách allowed_folders.

    Chạy trong STA thread executor của server.py — đã CoInitialize().

    Tham số arguments:
        entry_id (str):            ID hex của email cần di chuyển
        destination_folder (str):  Tên thư mục đích (phải trong allowlist)

    Trả về:
        dict {"status": "ok", "new_entry_id": str, "destination_folder": str}
        hoặc {"status": "error", "error": str}
    """
    start_ts = time.monotonic()

    # Bước 1: Lấy và validate tham số
    entry_id_raw = str(arguments.get("entry_id", "")).strip()
    if not entry_id_raw:
        return {"status": "error", "error": "Tham số entry_id không được để trống."}

    dest_folder = str(arguments.get("destination_folder", "")).strip()
    if not dest_folder:
        return {"status": "error", "error": "Tham số destination_folder không được để trống."}

    # Bước 2: Lấy allowed_folders và kiểm tra read_only
    allowed_folders = list(
        getattr(config, "ALLOWED_FOLDERS", None)
        or getattr(getattr(config, "security", None), "allowed_folders", None)
        or []
    )
    if not allowed_folders:
        return {"status": "error", "error": "Chưa cấu hình danh sách thư mục được phép."}

    # F-TOOL-04: fallback là True (fail-safe) thay vì False — nếu config bị lỗi,
    # tốt hơn là block thao tác ghi hơn là vô tình cho phép (ngược với compose.py)
    read_only = bool(
        getattr(config, "READ_ONLY_MODE", None)
        if getattr(config, "READ_ONLY_MODE", None) is not None
        else (
            getattr(getattr(config, "security", None), "read_only_mode", None)
            if getattr(getattr(config, "security", None), "read_only_mode", None) is not None
            else True
        )
    )
    if read_only:
        return {
            "status": "blocked",
            "error": "Server đang ở chế độ chỉ đọc (read_only_mode=True). Không thể di chuyển email.",
        }

    # Bước 3: Kiểm tra thư mục đích có trong allowlist
    import unicodedata as _ud
    dest_normalized = _ud.normalize("NFC", dest_folder.strip()).casefold()
    allowed_normalized = [_ud.normalize("NFC", a.strip()).casefold() for a in allowed_folders]
    if dest_normalized not in allowed_normalized:
        return {
            "status": "error",
            "error": f"Thư mục đích '{dest_folder}' không có trong danh sách được phép (allowed_folders).",
        }

    # Bước 4: Gọi COM bridge để di chuyển
    try:
        new_entry_id = com_bridge.move_email(
            entry_id=entry_id_raw,
            destination_folder=dest_folder,
            allowed_folders=allowed_folders,
        )
        duration_ms = int((time.monotonic() - start_ts) * 1000)
        _logger.info(
            "move_email: src_prefix=%s..., dest=%s, new_prefix=%s..., duration=%dms",
            entry_id_raw[:8], dest_folder, (new_entry_id or "")[:8], duration_ms
        )
        return {
            "status": "ok",
            "original_entry_id": entry_id_raw,
            "new_entry_id": new_entry_id or "",
            "destination_folder": dest_folder,
            "message": f"Email đã được di chuyển vào '{dest_folder}'.",
        }

    except Exception as exc:
        duration_ms = int((time.monotonic() - start_ts) * 1000)
        exc_type = type(exc).__name__
        _logger.warning("move_email thất bại: %s: %s", exc_type, exc)

        if "FolderNotAllowed" in exc_type or "NotAllowed" in exc_type:
            return {"status": "error", "error": "Email hoặc thư mục đích không được phép truy cập."}
        if "InvalidEmailId" in exc_type:
            return {"status": "error", "error": "entry_id không hợp lệ."}
        if "OutlookNotRunning" in exc_type:
            return {"status": "error", "error": "Outlook chưa mở. Vui lòng mở Outlook Desktop trước."}
        return {"status": "error", "error": "Không thể di chuyển email. Kiểm tra Outlook và thử lại."}


def handle_bulk_mark_read(arguments: dict, config, com_bridge) -> dict:
    """
    Đánh dấu tất cả email chưa đọc trong một folder là đã đọc.

    Chạy trong STA thread executor của server.py — đã CoInitialize().

    Tham số arguments:
        folder_name (str): Tên folder cần xử lý
        dry_run (bool):    Nếu True, chỉ preview — không thay đổi thực sự (mặc định True)
        max_emails (int):  Số email tối đa xử lý mỗi lần (mặc định 50, tối đa 100)

    Trả về:
        dry_run=True:  {"dry_run": True,  "folder": str, "would_mark": N, "preview": [...]}
        dry_run=False: {"dry_run": False, "folder": str, "marked_count": N, "errors": N}
    """
    start_ts = time.monotonic()

    # Bước 1: Validate tham số đầu vào
    folder_name = str(arguments.get("folder_name", "")).strip()
    if not folder_name:
        return {"status": "error", "error": "folder_name không được để trống."}

    dry_run = bool(arguments.get("dry_run", True))   # Mặc định True — an toàn
    max_emails = min(int(arguments.get("max_emails", 50)), 100)  # Giới hạn cứng 100

    # Bước 2: Kiểm tra read_only_mode — chặn khi dry_run=False
    # F-TOOL-04: fallback là True (fail-safe) thay vì False — nếu config bị lỗi,
    # tốt hơn là block thao tác ghi hơn là vô tình cho phép (ngược với compose.py)
    read_only = bool(
        getattr(config, "READ_ONLY_MODE", None)
        if getattr(config, "READ_ONLY_MODE", None) is not None
        else (
            getattr(getattr(config, "security", None), "read_only_mode", None)
            if getattr(getattr(config, "security", None), "read_only_mode", None) is not None
            else True
        )
    )
    if read_only and not dry_run:
        return {
            "status": "blocked",
            "error": "Server đang ở chế độ chỉ đọc (read_only_mode=True). Chỉ có thể dùng dry_run=True.",
        }

    # Bước 3: Kiểm tra folder_name có trong allowed_folders (chống IDOR)
    allowed_folders = list(
        getattr(config, "ALLOWED_FOLDERS", None)
        or getattr(getattr(config, "security", None), "allowed_folders", None)
        or []
    )
    if not allowed_folders:
        return {"status": "error", "error": "Chưa cấu hình danh sách thư mục được phép."}

    import unicodedata as _ud
    folder_norm = _ud.normalize("NFC", folder_name.strip()).casefold()
    allowed_norm = [_ud.normalize("NFC", a.strip()).casefold() for a in allowed_folders]
    if folder_norm not in allowed_norm:
        return {
            "status": "error",
            "error": f"Folder '{folder_name}' không có trong allowed_folders.",
        }

    # Bước 4: Lấy danh sách email chưa đọc trong folder chỉ định
    try:
        unread_emails = com_bridge.list_emails(
            folder_name=folder_name,
            max_count=max_emails,
            unread_only=True,
            allowed_folders=allowed_folders,
        )
    except Exception as exc:
        exc_type = type(exc).__name__
        _logger.warning("bulk_mark_read — list_emails thất bại: %s: %s", exc_type, exc)
        if "OutlookNotRunning" in exc_type:
            return {"status": "error", "error": "Outlook chưa mở. Vui lòng mở Outlook Desktop trước."}
        return {"status": "error", "error": f"Không thể lấy danh sách email: {exc_type}"}

    # Không có email chưa đọc — trả về sớm, không cần phân biệt dry_run
    if not unread_emails:
        return {
            "dry_run": dry_run,
            "folder": folder_name,
            "would_mark" if dry_run else "marked_count": 0,
            "message": "Không có email chưa đọc trong folder này.",
        }

    # Bước 5a: dry_run=True — chỉ preview, không thay đổi bất kỳ email nào
    if dry_run:
        duration_ms = int((time.monotonic() - start_ts) * 1000)
        _logger.info(
            "bulk_mark_read DRY-RUN: folder=%s, would_mark=%d, duration=%dms",
            folder_name, len(unread_emails), duration_ms,
        )
        return {
            "dry_run": True,
            "folder": folder_name,
            "would_mark": len(unread_emails),
            "preview": [
                {
                    "subject": e.get("subject", "")[:80],
                    "sender": e.get("sender_name", ""),
                }
                for e in unread_emails[:10]
            ],
            "note": (
                "dry_run=True: không có gì thay đổi. "
                "Đặt dry_run=False để thực hiện đánh dấu thật sự."
            ),
        }

    # Bước 5b: dry_run=False — thực hiện đánh dấu đã đọc từng email
    marked = 0
    errors = 0
    for email in unread_emails:
        try:
            com_bridge.mark_email_read(
                entry_id=email.get("entry_id", ""),
                read=True,
                allowed_folders=allowed_folders,
            )
            marked += 1
        except Exception as exc:
            _logger.warning(
                "bulk_mark_read — mark thất bại cho entry_id=%s: %s",
                str(email.get("entry_id", ""))[:8], exc,
            )
            errors += 1

    duration_ms = int((time.monotonic() - start_ts) * 1000)
    _logger.info(
        "bulk_mark_read DONE: folder=%s, marked=%d, errors=%d, duration=%dms",
        folder_name, marked, errors, duration_ms,
    )
    return {
        "dry_run": False,
        "folder": folder_name,
        "marked_count": marked,
        "errors": errors,
        "message": f"Đã đánh dấu {marked} email là đã đọc trong '{folder_name}'."
        + (f" ({errors} lỗi bỏ qua.)" if errors else ""),
    }


def handle_get_email_thread(arguments: dict, config, com_bridge) -> dict:
    """
    Lấy toàn bộ email trong cùng conversation thread (chuỗi hội thoại).

    Dùng ConversationID của email được chỉ định để tìm tất cả email liên quan
    trong các thư mục được phép (allowed_folders). Kết quả trả về được sắp xếp
    từ email cũ nhất đến mới nhất.

    Chạy trong STA thread executor của server.py — đã CoInitialize().

    Tham số arguments:
        entry_id (str):   ID hex của email bất kỳ trong thread cần xem
        max_emails (int): Số email tối đa muốn lấy (mặc định 20, tối đa 50)

    Trả về:
        dict {"thread": list, "count": int, "note": str}
        hoặc {"error": str} khi có lỗi
    """
    # Bước 1: Lấy và validate tham số entry_id
    entry_id_raw = str(arguments.get("entry_id", "")).strip()
    if not entry_id_raw:
        return {"error": "entry_id không được để trống."}

    # Bước 2: Lấy max_emails, giới hạn cứng tối đa 50
    max_emails = min(int(arguments.get("max_emails", 20)), 50)

    # Bước 3: Lấy allowed_folders từ config
    allowed_folders = list(
        getattr(config, "ALLOWED_FOLDERS", None)
        or getattr(getattr(config, "security", None), "allowed_folders", None)
        or []
    )
    if not allowed_folders:
        return {"error": "Chưa cấu hình allowed_folders."}

    # Bước 4: Gọi COM bridge để lấy thread
    try:
        thread = com_bridge.get_email_thread(
            entry_id=entry_id_raw,
            allowed_folders=allowed_folders,
            max_emails=max_emails,
        )
        _logger.info(
            "get_email_thread: entry_id_prefix=%s..., count=%d",
            entry_id_raw[:8], len(thread)
        )
        return {
            "thread": thread,
            "count": len(thread),
            "note": "Email được sắp xếp từ cũ đến mới trong cùng conversation thread.",
        }
    except Exception as exc:
        exc_type = type(exc).__name__
        _logger.warning("get_email_thread thất bại: %s: %s", exc_type, exc)
        if "FolderNotAllowed" in exc_type or "NotAllowed" in exc_type:
            return {"error": "Email không nằm trong thư mục được phép."}
        if "InvalidEmailId" in exc_type:
            return {"error": "entry_id không hợp lệ."}
        if "OutlookNotRunning" in exc_type:
            return {"error": "Outlook chưa mở."}
        return {"error": "Không thể lấy email thread."}


def handle_get_flagged_emails(arguments: dict, config, com_bridge) -> dict:
    """
    Lấy danh sách email đang được flag (follow-up) trong một folder cụ thể.

    Hữu ích cho PM để kiểm tra nhanh các email cần xử lý trong buổi sáng.
    Chạy trong STA thread executor của server.py — đã CoInitialize().

    Tham số arguments:
        folder_name (str): Tên folder cần tìm email flagged (phải trong allowed_folders)

    Trả về:
        dict {"flagged_emails": list, "count": int, "folder": str}
        hoặc {"error": str} khi có lỗi
    """
    import json as _json

    # Bước 1: Lấy và validate tham số folder_name
    folder_name = str(arguments.get("folder_name", "")).strip()
    if not folder_name:
        return {"error": "folder_name không được để trống."}

    # Bước 2: Lấy allowed_folders từ config
    allowed_folders = list(
        getattr(config, "ALLOWED_FOLDERS", None)
        or getattr(getattr(config, "security", None), "allowed_folders", None)
        or []
    )
    if not allowed_folders:
        return {"error": "Chưa cấu hình allowed_folders."}

    # Bước 3: Gọi COM bridge để lấy danh sách email flagged
    try:
        flagged = com_bridge.get_flagged_emails(
            folder_name=folder_name,
            allowed_folders=allowed_folders,
        )
        _logger.info(
            "get_flagged_emails: folder=%s, count=%d",
            folder_name, len(flagged)
        )
        return {
            "flagged_emails": flagged,
            "count": len(flagged),
            "folder": folder_name,
        }
    except Exception as exc:
        exc_type = type(exc).__name__
        _logger.warning("get_flagged_emails thất bại: %s: %s", exc_type, exc)
        if "FolderNotAllowed" in exc_type or "NotAllowed" in exc_type:
            return {"error": f"Folder '{folder_name}' không nằm trong danh sách được phép."}
        if "OutlookNotRunning" in exc_type:
            return {"error": "Outlook chưa mở."}
        return {"error": "Không thể lấy danh sách email flagged."}


def handle_get_project_snapshot(arguments: dict, config, com_bridge) -> dict:
    """
    Lấy snapshot tổng hợp trạng thái một project folder — thiết kế cho PM.

    Một lệnh duy nhất thay thế 4-5 queries riêng lẻ: tổng email, chưa đọc,
    flagged (cần theo dõi), top senders, và danh sách email gần đây.
    Chạy trong STA thread executor của server.py — đã CoInitialize().

    Tham số arguments:
        folder_name (str): Tên folder project (phải trong allowed_folders)
        days_back (int):   Số ngày nhìn lại (mặc định 14, tối đa 90)

    Trả về:
        dict tổng hợp với total_received, unread_count, flagged_count,
        flagged_emails, top_senders, recent_emails, summary
        hoặc {"error": str} khi có lỗi
    """
    # Bước 1: Lấy và validate tham số folder_name
    folder_name = str(arguments.get("folder_name", "")).strip()
    if not folder_name:
        return {"error": "folder_name không được để trống."}

    # Bước 2: Lấy days_back, giới hạn trong khoảng 1-90
    days_back = int(arguments.get("days_back", 14))
    days_back = max(1, min(days_back, 90))

    # Bước 3: Lấy allowed_folders từ config
    allowed_folders = list(
        getattr(config, "ALLOWED_FOLDERS", None)
        or getattr(getattr(config, "security", None), "allowed_folders", None)
        or []
    )
    if not allowed_folders:
        return {"error": "Chưa cấu hình allowed_folders."}

    # Bước 4: Gọi COM bridge để lấy snapshot tổng hợp
    try:
        snapshot = com_bridge.get_project_snapshot(
            folder_name=folder_name,
            days_back=days_back,
            allowed_folders=allowed_folders,
        )
        _logger.info(
            "get_project_snapshot: folder=%s, days_back=%d, total=%d",
            folder_name, days_back, snapshot.get("total_received", 0)
        )
        return snapshot
    except Exception as exc:
        exc_type = type(exc).__name__
        _logger.warning("get_project_snapshot thất bại: %s: %s", exc_type, exc)
        if "FolderNotAllowed" in exc_type or "NotAllowed" in exc_type:
            return {"error": f"Folder '{folder_name}' không nằm trong danh sách được phép."}
        if "OutlookNotRunning" in exc_type:
            return {"error": "Outlook chưa mở."}
        return {"error": "Không thể lấy project snapshot."}
