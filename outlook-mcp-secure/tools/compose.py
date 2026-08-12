"""
Module compose — MCP tools soạn thảo và trả lời email QUA Outlook Desktop.

NGUYÊN TẮC AN TOÀN TUYỆT ĐỐI:
  - Module này KHÔNG BAO GIỜ gọi .Send() hay bất kỳ phương thức gửi tự động nào.
  - Mọi thao tác chỉ gọi .Display() để mở Outlook window — người dùng phải
    tự nhấn Send trong Outlook Desktop.
  - Mọi địa chỉ email được validate theo RFC 5322 trước khi dùng.
  - Kiểm tra read_only_mode: nếu bật, từ chối ngay không xử lý tiếp.
  - COM objects được release sau mỗi thao tác qua context manager.
  - Mọi action được ghi vào audit log, không ghi nội dung email.

Tools được đăng ký:
  1. compose_new_email  — soạn email mới
  2. reply_to_email     — trả lời một email có sẵn
  3. forward_email      — chuyển tiếp một email có sẵn
"""

from __future__ import annotations

import html as _html_module
import re
import unicodedata
from typing import Any

import gc

# Thư viện COM của Windows
import win32com.client
import pythoncom
import pywintypes


# ============================================================
# Hằng số nội bộ
# ============================================================

# Độ dài tối đa cho subject
_MAX_SUBJECT_LENGTH: int = 500

# Độ dài tối đa cho body email (50 000 ký tự)
_MAX_BODY_LENGTH: int = 50_000

# Độ dài tối đa entry_id (hex string)
_MAX_ENTRY_ID_LENGTH: int = 256

# Số lượng người nhận tối đa mặc định (có thể override từ config)
_DEFAULT_MAX_RECIPIENTS: int = 20

# Regex kiểm tra entry_id hợp lệ: chỉ chứa ký tự hex
_ENTRY_ID_PATTERN: re.Pattern = re.compile(r'^[0-9A-Fa-f]+$')

# Regex kiểm tra email đơn giản theo RFC 5322 local part + domain
# (email-validator library được dùng thêm ở validator.py nếu cần)
_EMAIL_PATTERN: re.Pattern = re.compile(
    r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$'
)


# ============================================================
# Lớp ngoại lệ nội bộ
# ============================================================

class ReadOnlyModeError(Exception):
    """Server đang ở chế độ chỉ đọc — không cho phép soạn/gửi email."""


class ValidationError(Exception):
    """Dữ liệu đầu vào không hợp lệ."""


class OutlookNotRunningError(Exception):
    """Outlook Desktop chưa được mở."""


class OutlookOperationError(Exception):
    """Thao tác COM với Outlook thất bại."""


class FolderNotAllowedError(Exception):
    """Email nằm trong thư mục ngoài danh sách được phép (allowlist)."""


# ============================================================
# Hàm validate nội bộ
# ============================================================

def _validate_email_address(address: str) -> str:
    """
    Kiểm tra một địa chỉ email có hợp lệ không.

    Nhận vào: chuỗi địa chỉ email cần kiểm tra.
    Trả về:   địa chỉ đã được strip whitespace nếu hợp lệ.
    Raises:   ValidationError nếu không hợp lệ.
    """
    # Bước 1: Xóa khoảng trắng đầu/cuối
    address = address.strip()

    # Bước 2: Kiểm tra không rỗng
    if not address:
        raise ValidationError("Địa chỉ email không được để trống.")

    # Bước 3: Kiểm tra độ dài hợp lý (email không nên dài quá 320 ký tự theo RFC 5321)
    if len(address) > 320:
        raise ValidationError(f"Địa chỉ email quá dài: {len(address)} ký tự (tối đa 320).")

    # Bước 4: Kiểm tra null bytes và control characters
    if any(ord(c) < 0x20 for c in address):
        raise ValidationError("Địa chỉ email chứa ký tự điều khiển không hợp lệ.")

    # Bước 5: Kiểm tra định dạng cơ bản theo regex RFC 5322
    if not _EMAIL_PATTERN.match(address):
        raise ValidationError(f"Địa chỉ email không đúng định dạng: '{address}'.")

    return address


def _validate_email_list(raw_to: str | list[str], field_name: str = "to") -> list[str]:
    """
    Validate danh sách địa chỉ email từ tham số tool.

    Chấp nhận chuỗi một địa chỉ hoặc list các địa chỉ.
    Trả về list địa chỉ đã validate.
    Raises ValidationError nếu bất kỳ địa chỉ nào không hợp lệ.
    """
    # Bước 1: Chuẩn hóa về list
    if isinstance(raw_to, str):
        # Hỗ trợ nhập nhiều địa chỉ cách nhau bởi dấu phẩy hoặc chấm phẩy
        addresses = [a for a in re.split(r'[;,]', raw_to) if a.strip()]
    elif isinstance(raw_to, list):
        addresses = raw_to
    else:
        raise ValidationError(f"Tham số '{field_name}' phải là chuỗi hoặc danh sách địa chỉ email.")

    # Bước 2: Không được rỗng
    if not addresses:
        raise ValidationError(f"Tham số '{field_name}' không được để trống.")

    # Bước 3: Validate từng địa chỉ
    validated: list[str] = []
    for addr in addresses:
        validated.append(_validate_email_address(addr))

    return validated


def _validate_entry_id(entry_id: Any) -> str:
    """
    Kiểm tra entry_id (định danh email trong Outlook MAPI store) hợp lệ.

    entry_id phải là hex string thuần túy, tối đa 256 ký tự.
    Raises ValidationError nếu không hợp lệ.
    """
    if not isinstance(entry_id, str):
        raise ValidationError("entry_id phải là chuỗi ký tự hex.")

    entry_id = entry_id.strip()

    # Kiểm tra độ dài
    if not entry_id:
        raise ValidationError("entry_id không được để trống.")
    if len(entry_id) > _MAX_ENTRY_ID_LENGTH:
        raise ValidationError(f"entry_id quá dài: {len(entry_id)} ký tự (tối đa {_MAX_ENTRY_ID_LENGTH}).")

    # Kiểm tra null bytes
    if '\x00' in entry_id:
        raise ValidationError("entry_id chứa null byte không hợp lệ.")

    # Kiểm tra chỉ chứa ký tự hex
    if not _ENTRY_ID_PATTERN.match(entry_id):
        raise ValidationError("entry_id chỉ được chứa ký tự hex (0-9, A-F).")

    return entry_id


def _validate_subject(subject: Any) -> str:
    """
    Kiểm tra tiêu đề email hợp lệ.

    Raises ValidationError nếu không hợp lệ hoặc quá dài.
    """
    if not isinstance(subject, str):
        raise ValidationError("subject phải là chuỗi ký tự.")

    # Xóa khoảng trắng đầu/cuối và chuẩn hóa Unicode NFC
    subject = unicodedata.normalize('NFC', subject.strip())

    if not subject:
        raise ValidationError("subject (tiêu đề email) không được để trống.")

    if len(subject) > _MAX_SUBJECT_LENGTH:
        raise ValidationError(
            f"subject quá dài: {len(subject)} ký tự (tối đa {_MAX_SUBJECT_LENGTH})."
        )

    # Kiểm tra control characters (ngoại trừ tab \t)
    if any(ord(c) < 0x20 and c not in ('\t',) for c in subject):
        raise ValidationError("subject chứa ký tự điều khiển không hợp lệ.")

    return subject


def _validate_body(body: Any) -> str:
    """
    Kiểm tra nội dung email hợp lệ.

    Cho phép body rỗng (email chỉ có chủ đề).
    Raises ValidationError nếu quá dài.
    """
    if not isinstance(body, str):
        raise ValidationError("body phải là chuỗi ký tự.")

    # Chuẩn hóa Unicode NFC
    body = unicodedata.normalize('NFC', body)

    if len(body) > _MAX_BODY_LENGTH:
        raise ValidationError(
            f"body quá dài: {len(body)} ký tự (tối đa {_MAX_BODY_LENGTH})."
        )

    return body


# ============================================================
# Hàm kiểm tra bảo mật folder
# ============================================================

def _verify_folder_allowed(mail_item, allowed_folders: list[str]) -> None:
    """
    Xác minh rằng email nằm trong một thư mục thuộc danh sách cho phép (allowlist).

    Mục đích: Ngăn chặn tấn công IDOR (Insecure Direct Object Reference —
    truy cập trái phép đối tượng qua ID), trong đó kẻ tấn công cung cấp
    entry_id của email bí mật nằm ngoài allowlist.

    Nhận vào:
        mail_item      -- COM MailItem đã được resolve từ entry_id
        allowed_folders -- Danh sách tên thư mục được phép (từ config.security.allowed_folders)

    Raises:
        FolderNotAllowedError: Nếu thư mục chứa email không nằm trong allowlist.
        OutlookOperationError: Nếu không đọc được thông tin thư mục từ COM.
    """
    try:
        # Bước 1: Lấy đối tượng thư mục cha chứa email này
        parent_folder = mail_item.Parent
        folder_name: str = getattr(parent_folder, "Name", "") or ""
    except pywintypes.error as e:
        raise OutlookOperationError(
            "Không thể xác minh thư mục chứa email."
        ) from e

    # Bước 2: Từ chối ngay nếu không đọc được tên thư mục
    if not folder_name:
        raise FolderNotAllowedError(
            "Không thể xác minh thư mục chứa email. Từ chối truy cập."
        )

    # Bước 3: Chuẩn hóa để so sánh không phân biệt hoa/thường và khoảng trắng
    normalized_actual = folder_name.strip().casefold()
    normalized_allowed = [a.strip().casefold() for a in allowed_folders]

    # Bước 4: Kiểm tra folder có nằm trong allowlist không
    if normalized_actual not in normalized_allowed:
        raise FolderNotAllowedError(
            "Email không nằm trong thư mục được phép truy cập."
        )


# ============================================================
# Hàm COM nội bộ — chạy trong STA thread
# ============================================================

def _get_outlook_app():
    """
    Lấy Outlook Application đang chạy qua GetActiveObject.

    Dùng GetActiveObject thay vì Dispatch để KHÔNG tạo instance Outlook mới.
    Nếu Outlook chưa mở, raise OutlookNotRunningError.
    """
    try:
        return win32com.client.GetActiveObject('Outlook.Application')
    except pywintypes.error:
        raise OutlookNotRunningError(
            "Outlook chưa được mở. Vui lòng mở Outlook Desktop trước khi dùng tính năng này."
        )


def _text_to_html_utf8(text: str) -> str:
    """
    Bao plain text vào HTML với charset UTF-8 để Outlook hiển thị đúng tiếng Việt.
    Cần dùng HTMLBody thay Body vì Body dùng ANSI codepage (CP1252) làm mất dấu.
    """
    escaped = _html_module.escape(text, quote=False)
    escaped = escaped.replace("\r\n", "<br>\n").replace("\r", "<br>\n").replace("\n", "<br>\n")
    return (
        '<html><head>'
        '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
        '</head>'
        '<body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000;">'
        f'{escaped}'
        '</body></html>'
    )


def _prepend_to_html_body(new_text: str, existing_html: str) -> str:
    """
    Chèn plain text mới vào đầu HTMLBody có sẵn (dùng cho reply/forward).
    Giữ nguyên nội dung email gốc phía dưới.
    """
    escaped = _html_module.escape(new_text, quote=False)
    escaped = escaped.replace("\r\n", "<br>\n").replace("\r", "<br>\n").replace("\n", "<br>\n")
    new_block = (
        '<div style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000;">'
        f'{escaped}'
        '</div><br>'
    )
    import re as _re
    match = _re.search(r'(<body[^>]*>)', existing_html, _re.IGNORECASE)
    if match:
        pos = match.end()
        return existing_html[:pos] + new_block + existing_html[pos:]
    return new_block + existing_html


def _com_open_compose(
    to_addresses: list[str],
    subject: str,
    body: str,
    cc_addresses: list[str] | None = None,
    importance_value: int = 1,  # 1 = olNormal
    bcc_addresses: list[str] | None = None,
) -> str:
    """
    Tạo email mới trong Outlook và mở cửa sổ soạn thảo.

    Tuyệt đối KHÔNG gọi .Send(). Chỉ gọi .Display().
    Trả về entry_id của draft item đã tạo (dạng hex string).

    Hằng số importance của Outlook MAPI:
      olImportanceLow    = 0
      olImportanceNormal = 1
      olImportanceHigh   = 2
    """
    app = None
    mail_item = None
    try:
        # Bước 1: Lấy Outlook application đang chạy
        app = _get_outlook_app()

        # Bước 2: Tạo mail item mới (olMailItem = 0)
        mail_item = app.CreateItem(0)

        # Bước 3: Đặt người nhận chính (To)
        mail_item.To = "; ".join(to_addresses)

        # Bước 4: Đặt người nhận CC nếu có
        if cc_addresses:
            mail_item.CC = "; ".join(cc_addresses)

        # Bước 4b: Đặt BCC nếu có — người nhận không thấy danh sách BCC
        if bcc_addresses:
            mail_item.BCC = "; ".join(bcc_addresses)

        # Bước 5: Đặt tiêu đề
        mail_item.Subject = subject

        # Bước 6: Đặt nội dung qua HTMLBody để hỗ trợ tiếng Việt đầy đủ dấu
        mail_item.HTMLBody = _text_to_html_utf8(body)

        # Bước 7: Đặt mức độ ưu tiên
        mail_item.Importance = importance_value

        # Bước 8: Lưu draft để có entry_id trước khi hiển thị
        mail_item.Save()

        # Bước 9: Lấy entry_id của draft (dùng để báo cáo về)
        draft_entry_id: str = mail_item.EntryID or ""

        # Bước 10: Mở cửa sổ Outlook — người dùng tự nhấn Send
        # TUYỆT ĐỐI KHÔNG gọi .Send() ở đây
        mail_item.Display()

        return draft_entry_id

    except OutlookNotRunningError:
        raise
    except pywintypes.error as e:
        raise OutlookOperationError(
            f"Lỗi Windows COM khi tạo email mới (hresult=0x{e.winerror:08X}). "
            "Đảm bảo Outlook đang chạy."
        ) from e
    except pythoncom.error as e:
        raise OutlookOperationError(
            f"Mất kết nối với Outlook. Vui lòng khởi động lại Outlook và thử lại. ({e})"
        ) from e
    finally:
        # Bước 11: Release COM objects để tránh memory leak
        if mail_item is not None:
            try:
                win32com.client.ReleaseComObject(mail_item)
            except Exception:
                pass
        if app is not None:
            try:
                win32com.client.ReleaseComObject(app)
            except Exception:
                pass
        gc.collect()


def _com_open_reply(
    entry_id: str,
    body: str,
    allowed_folders: list[str],
    reply_all: bool = False,
    additional_cc: list[str] | None = None,
) -> str:
    """
    Mở cửa sổ reply cho một email trong Outlook.

    Tuyệt đối KHÔNG gọi .Send(). Chỉ gọi .Display().
    Trả về entry_id của reply item đã tạo.

    Tham số:
      allowed_folders -- Danh sách thư mục được phép (từ config.security.allowed_folders).
                         Bắt buộc phải truyền để kiểm tra bảo mật IDOR trước khi mở reply.

    Hằng số Outlook:
      olFolderInbox = 6 (không dùng ở đây nhưng giải thích context)
    """
    app = None
    namespace = None
    original_item = None
    reply_item = None
    try:
        # Bước 1: Lấy Outlook application đang chạy
        app = _get_outlook_app()

        # Bước 2: Lấy MAPI namespace để truy cập email theo entry_id
        namespace = app.GetNamespace("MAPI")

        # Bước 3: Tìm email gốc theo entry_id
        # GetItemFromID nhận entry_id dạng hex string
        original_item = namespace.GetItemFromID(entry_id)

        # Bước 3b: Kiểm tra bảo mật — xác minh email gốc nằm trong allowed folder
        # QUAN TRỌNG: Phải kiểm tra TRƯỚC khi tạo reply để ngăn tấn công IDOR.
        # Kẻ tấn công có thể cung cấp entry_id của email bí mật nằm ngoài allowlist,
        # server sẽ mở cửa sổ reply với nội dung email đó, khiến người dùng vô tình
        # tiết lộ nội dung email bí mật qua reply body được hiển thị sẵn.
        _verify_folder_allowed(original_item, allowed_folders)

        # Bước 4: Tạo reply item
        # Reply()    = trả lời người gửi
        # ReplyAll() = trả lời tất cả (To + CC)
        if reply_all:
            reply_item = original_item.ReplyAll()
        else:
            reply_item = original_item.Reply()

        # Bước 5: Đặt nội dung reply ở đầu, giữ lại phần email gốc phía dưới
        # Dùng HTMLBody để hỗ trợ tiếng Việt đầy đủ dấu (Body dùng ANSI làm mất dấu)
        existing_html: str = reply_item.HTMLBody or ""
        if existing_html:
            reply_item.HTMLBody = _prepend_to_html_body(body, existing_html)
        else:
            reply_item.HTMLBody = _text_to_html_utf8(body)

        # Bước 6: Thêm CC bổ sung nếu có
        if additional_cc:
            current_cc: str = reply_item.CC or ""
            extra_cc = "; ".join(additional_cc)
            if current_cc:
                reply_item.CC = current_cc + "; " + extra_cc
            else:
                reply_item.CC = extra_cc

        # Bước 7: Lưu reply draft để lấy entry_id
        reply_item.Save()
        reply_entry_id: str = reply_item.EntryID or ""

        # Bước 8: Mở cửa sổ Outlook cho người dùng xem lại
        # TUYỆT ĐỐI KHÔNG gọi .Send()
        reply_item.Display()

        return reply_entry_id

    except OutlookNotRunningError:
        raise
    except FolderNotAllowedError:
        # Re-raise để caller xử lý và trả về lỗi bảo mật phù hợp
        raise
    except pywintypes.error as e:
        # Mã lỗi 0x8004010F: không tìm thấy email (entry_id không tồn tại)
        if hasattr(e, 'winerror') and e.winerror == -2147221233:
            raise OutlookOperationError(
                "Không tìm thấy email gốc. entry_id có thể đã cũ hoặc email đã bị xóa."
            ) from e
        raise OutlookOperationError(
            f"Lỗi Windows COM khi mở reply (hresult=0x{e.winerror:08X}). "
            "Đảm bảo Outlook đang chạy."
        ) from e
    except pythoncom.error as e:
        raise OutlookOperationError(
            f"Mất kết nối với Outlook khi mở reply. Khởi động lại Outlook và thử lại. ({e})"
        ) from e
    finally:
        # Bước 9: Release tất cả COM objects theo thứ tự ngược
        for obj in [reply_item, original_item, namespace, app]:
            if obj is not None:
                try:
                    win32com.client.ReleaseComObject(obj)
                except Exception:
                    pass
        gc.collect()


def _com_open_forward(
    entry_id: str,
    to_addresses: list[str],
    allowed_folders: list[str],
    note: str = "",
) -> str:
    """
    Mở cửa sổ forward cho một email trong Outlook.

    Tuyệt đối KHÔNG gọi .Send(). Chỉ gọi .Display().
    Phần ghi chú (note) của người dùng được chèn trước nội dung email gốc.
    Trả về entry_id của forward item đã tạo.

    Tham số:
      allowed_folders -- Danh sách thư mục được phép (từ config.security.allowed_folders).
                         Bắt buộc phải truyền để kiểm tra bảo mật IDOR trước khi mở forward.
    """
    app = None
    namespace = None
    original_item = None
    forward_item = None
    try:
        # Bước 1: Lấy Outlook application đang chạy
        app = _get_outlook_app()

        # Bước 2: Lấy MAPI namespace
        namespace = app.GetNamespace("MAPI")

        # Bước 3: Tìm email gốc theo entry_id
        original_item = namespace.GetItemFromID(entry_id)

        # Bước 3b: Kiểm tra bảo mật — xác minh email gốc nằm trong allowed folder
        # QUAN TRỌNG: Phải kiểm tra TRƯỚC khi tạo forward để ngăn tấn công IDOR.
        # Kẻ tấn công có thể cung cấp entry_id của email bí mật nằm ngoài allowlist,
        # server sẽ mở cửa sổ forward với toàn bộ nội dung email đó cho người dùng thấy.
        _verify_folder_allowed(original_item, allowed_folders)

        # Bước 4: Tạo forward item từ email gốc
        forward_item = original_item.Forward()

        # Bước 5: Đặt người nhận forward
        forward_item.To = "; ".join(to_addresses)

        # Bước 6: Thêm ghi chú của người dùng trước nội dung email gốc
        # Dùng HTMLBody để hỗ trợ tiếng Việt đầy đủ dấu (Body dùng ANSI làm mất dấu)
        if note:
            existing_html: str = forward_item.HTMLBody or ""
            if existing_html:
                forward_item.HTMLBody = _prepend_to_html_body(note, existing_html)
            else:
                forward_item.HTMLBody = _text_to_html_utf8(note)

        # Bước 7: Lưu draft để lấy entry_id
        forward_item.Save()
        forward_entry_id: str = forward_item.EntryID or ""

        # Bước 8: Mở cửa sổ Outlook — người dùng tự xem lại và nhấn Send
        # TUYỆT ĐỐI KHÔNG gọi .Send()
        forward_item.Display()

        return forward_entry_id

    except OutlookNotRunningError:
        raise
    except FolderNotAllowedError:
        # Re-raise để caller xử lý và trả về lỗi bảo mật phù hợp
        raise
    except pywintypes.error as e:
        if hasattr(e, 'winerror') and e.winerror == -2147221233:
            raise OutlookOperationError(
                "Không tìm thấy email gốc để forward. entry_id có thể đã cũ hoặc email đã bị xóa."
            ) from e
        raise OutlookOperationError(
            f"Lỗi Windows COM khi mở forward (hresult=0x{e.winerror:08X}). "
            "Đảm bảo Outlook đang chạy."
        ) from e
    except pythoncom.error as e:
        raise OutlookOperationError(
            f"Mất kết nối với Outlook khi mở forward. Khởi động lại Outlook và thử lại. ({e})"
        ) from e
    finally:
        # Bước 9: Release tất cả COM objects
        for obj in [forward_item, original_item, namespace, app]:
            if obj is not None:
                try:
                    win32com.client.ReleaseComObject(obj)
                except Exception:
                    pass
        gc.collect()


# ============================================================
# Hàm chạy COM trong STA executor (được gọi từ server.py)
# ============================================================

def _run_in_sta(sta_executor, func, *args, **kwargs):
    """
    Chạy một hàm COM trong STA thread executor để đảm bảo thread-safety.

    sta_executor: concurrent.futures.Executor với max_workers=1 (từ server.py)
    Trả về kết quả hoặc raise exception từ STA thread.
    """
    future = sta_executor.submit(func, *args, **kwargs)
    return future.result(timeout=30)  # timeout 30 giây theo config


# ============================================================
# Hàm đăng ký tools vào MCP server
# ============================================================

def register_tools(server, outlook_com_bridge, audit, config) -> None:
    """
    Đăng ký 3 MCP tools soạn/trả lời/forward email vào MCP server.

    Hàm này được gọi từ tools/__init__.py -> register_all_tools().

    Tham số:
        server           -- MCP Server instance
        outlook_com_bridge -- OutlookCOMBridge instance (dùng sta_executor)
        audit            -- AuditLogger instance
        config           -- Config object (có .security.read_only_mode,
                            .security.max_recipients_per_draft)
    """

    # ----------------------------------------------------------
    # Tool 1: compose_new_email — Soạn email mới
    # ----------------------------------------------------------
    @server.tool(
        name="compose_new_email",
        description=(
            "Soạn một email mới trong Outlook Desktop. "
            "Outlook sẽ mở cửa sổ soạn thảo — bạn phải tự nhấn Send để gửi. "
            "Tool này KHÔNG tự động gửi email."
        ),
    )
    async def tool_compose_new_email(
        to: str | list[str],
        subject: str,
        body: str,
        cc: str | list[str] | None = None,
        bcc: str | list[str] | None = None,
        importance: str = "normal",
    ) -> dict[str, str]:
        """
        Mở cửa sổ soạn email mới trong Outlook.

        Tham số:
            to         -- địa chỉ email người nhận (chuỗi hoặc danh sách)
            subject    -- tiêu đề email (tối đa 500 ký tự)
            body       -- nội dung email (tối đa 50 000 ký tự)
            cc         -- CC gửi thêm để biết thông tin (tùy chọn)
            bcc        -- BCC gửi kín — người nhận chính không thấy (tùy chọn)
            importance -- mức độ ưu tiên: "low", "normal", "high" (mặc định "normal")

        Trả về dict với status và message hướng dẫn người dùng.
        """
        # Bước 1: Kiểm tra read_only_mode — nếu bật thì từ chối ngay
        if getattr(getattr(config, 'security', None), 'read_only_mode', True):
            audit.log_blocked(
                tool="compose_new_email",
                block_reason="read_only_mode",
                risk_level="low",
                duration_ms=0,
            )
            return {
                "status": "blocked",
                "error": (
                    "Chế độ chỉ đọc (read_only_mode) đang bật. "
                    "Để soạn email, hãy đặt read_only_mode = false trong config.toml."
                ),
            }

        # Bước 2: Validate địa chỉ người nhận chính
        try:
            to_list = _validate_email_list(to, field_name="to")
        except ValidationError as e:
            audit.log_blocked(
                tool="compose_new_email",
                block_reason=f"validation_error_to: {e}",
                risk_level="low",
                duration_ms=0,
            )
            return {"status": "error", "error": f"Địa chỉ người nhận không hợp lệ: {e}"}

        # Bước 3: Validate CC nếu có
        cc_list: list[str] | None = None
        if cc is not None:
            try:
                cc_list = _validate_email_list(cc, field_name="cc")
            except ValidationError as e:
                audit.log_blocked(
                    tool="compose_new_email",
                    block_reason=f"validation_error_cc: {e}",
                    risk_level="low",
                    duration_ms=0,
                )
                return {"status": "error", "error": f"Địa chỉ CC không hợp lệ: {e}"}

        # Bước 3b: Validate BCC nếu có
        bcc_list: list[str] | None = None
        if bcc is not None:
            try:
                bcc_list = _validate_email_list(bcc, field_name="bcc")
            except ValidationError as e:
                audit.log_blocked(
                    tool="compose_new_email",
                    block_reason=f"validation_error_bcc: {e}",
                    risk_level="low",
                    duration_ms=0,
                )
                return {"status": "error", "error": f"Địa chỉ BCC không hợp lệ: {e}"}

        # Bước 4: Kiểm tra tổng số người nhận không vượt giới hạn cấu hình
        max_recipients: int = getattr(
            getattr(config, 'security', None),
            'max_recipients_per_draft',
            _DEFAULT_MAX_RECIPIENTS,
        )
        total_recipients = (
            len(to_list)
            + (len(cc_list)  if cc_list  else 0)
            + (len(bcc_list) if bcc_list else 0)
        )
        if total_recipients > max_recipients:
            audit.log_blocked(
                tool="compose_new_email",
                block_reason=f"too_many_recipients: {total_recipients} > {max_recipients}",
                risk_level="medium",
                duration_ms=0,
            )
            return {
                "status": "error",
                "error": (
                    f"Quá nhiều người nhận ({total_recipients}). "
                    f"Tối đa cho phép là {max_recipients} địa chỉ."
                ),
            }

        # Bước 5: Validate subject
        try:
            subject_clean = _validate_subject(subject)
        except ValidationError as e:
            audit.log_blocked(
                tool="compose_new_email",
                block_reason=f"validation_error_subject: {e}",
                risk_level="low",
                duration_ms=0,
            )
            return {"status": "error", "error": f"Tiêu đề email không hợp lệ: {e}"}

        # Bước 6: Validate body
        try:
            body_clean = _validate_body(body)
        except ValidationError as e:
            audit.log_blocked(
                tool="compose_new_email",
                block_reason=f"validation_error_body: {e}",
                risk_level="low",
                duration_ms=0,
            )
            return {"status": "error", "error": f"Nội dung email không hợp lệ: {e}"}

        # Bước 7: Chuyển đổi importance thành hằng số Outlook MAPI
        importance_map: dict[str, int] = {"low": 0, "normal": 1, "high": 2}
        importance_value: int = importance_map.get(
            str(importance).lower().strip(), 1  # Mặc định normal
        )

        # Bước 8: Ghi audit log bắt đầu thao tác
        # Không ghi subject/body plaintext — chỉ ghi metadata
        import hashlib, time
        start_ts = time.monotonic()
        to_hash = hashlib.sha256(";".join(to_list).encode("utf-8")).hexdigest()
        subject_hash = hashlib.sha256(subject_clean.encode("utf-8")).hexdigest()
        audit.log_start(
            tool="compose_new_email",
            params={
                "to_hash": f"sha256:{to_hash[:16]}",
                "subject_hash": f"sha256:{subject_hash[:16]}",
                "recipient_count": total_recipients,
                "importance": importance,
            },
        )

        # Bước 9: Thực thi COM trong STA thread
        try:
            draft_entry_id = _run_in_sta(
                outlook_com_bridge.sta_executor,
                _com_open_compose,
                to_list,
                subject_clean,
                body_clean,
                cc_list,
                importance_value,
                bcc_list,
            )
        except OutlookNotRunningError as e:
            duration_ms = int((time.monotonic() - start_ts) * 1000)
            audit.log_error(
                tool="compose_new_email",
                error="outlook_not_running",
                duration_ms=duration_ms,
            )
            return {"status": "error", "error": str(e)}
        except OutlookOperationError as e:
            duration_ms = int((time.monotonic() - start_ts) * 1000)
            audit.log_error(
                tool="compose_new_email",
                error="com_operation_failed",
                duration_ms=duration_ms,
            )
            return {"status": "error", "error": str(e)}
        except Exception:
            duration_ms = int((time.monotonic() - start_ts) * 1000)
            audit.log_error(
                tool="compose_new_email",
                error="internal_error",
                duration_ms=duration_ms,
            )
            return {"status": "error", "error": "Lỗi nội bộ. Kiểm tra log để biết thêm chi tiết."}

        # Bước 10: Ghi audit log thành công
        duration_ms = int((time.monotonic() - start_ts) * 1000)
        audit.log_success(tool="compose_new_email", duration_ms=duration_ms)

        return {
            "status": "opened",
            "message": (
                "Đã mở cửa sổ soạn email trong Outlook. "
                "Vui lòng kiểm tra nội dung và nhấn Send để gửi."
            ),
            "draft_entry_id": draft_entry_id,
        }

    # ----------------------------------------------------------
    # Tool 2: reply_to_email — Trả lời một email
    # ----------------------------------------------------------
    @server.tool(
        name="reply_to_email",
        description=(
            "Trả lời một email trong Outlook Desktop. "
            "Outlook sẽ mở cửa sổ reply — bạn phải tự nhấn Send để gửi. "
            "Tool này KHÔNG tự động gửi email."
        ),
    )
    async def tool_reply_to_email(
        entry_id: str,
        body: str,
        reply_all: bool = False,
        additional_cc: list[str] | None = None,
    ) -> dict[str, str]:
        """
        Mở cửa sổ trả lời email trong Outlook.

        Tham số:
            entry_id       -- ID hex của email gốc cần trả lời
            body           -- nội dung trả lời (tối đa 50 000 ký tự)
            reply_all      -- true = Reply All (trả lời tất cả), false = Reply (chỉ người gửi)
            additional_cc  -- danh sách CC bổ sung (tùy chọn, tối đa 20 địa chỉ)

        Trả về dict với status và message hướng dẫn.
        """
        # Bước 1: Kiểm tra read_only_mode
        if getattr(getattr(config, 'security', None), 'read_only_mode', True):
            audit.log_blocked(
                tool="reply_to_email",
                block_reason="read_only_mode",
                risk_level="low",
                duration_ms=0,
            )
            return {
                "status": "blocked",
                "error": (
                    "Chế độ chỉ đọc (read_only_mode) đang bật. "
                    "Để trả lời email, hãy đặt read_only_mode = false trong config.toml."
                ),
            }

        # Bước 2: Validate entry_id
        try:
            entry_id_clean = _validate_entry_id(entry_id)
        except ValidationError as e:
            audit.log_blocked(
                tool="reply_to_email",
                block_reason=f"validation_error_entry_id: {e}",
                risk_level="medium",
                duration_ms=0,
            )
            return {"status": "error", "error": f"entry_id không hợp lệ: {e}"}

        # Bước 3: Validate body
        try:
            body_clean = _validate_body(body)
        except ValidationError as e:
            audit.log_blocked(
                tool="reply_to_email",
                block_reason=f"validation_error_body: {e}",
                risk_level="low",
                duration_ms=0,
            )
            return {"status": "error", "error": f"Nội dung reply không hợp lệ: {e}"}

        # Bước 4: Validate additional_cc nếu có
        cc_list: list[str] | None = None
        if additional_cc:
            # Giới hạn tối đa 20 CC bổ sung theo PLAN.md
            if len(additional_cc) > 20:
                audit.log_blocked(
                    tool="reply_to_email",
                    block_reason=f"too_many_additional_cc: {len(additional_cc)}",
                    risk_level="low",
                    duration_ms=0,
                )
                return {
                    "status": "error",
                    "error": f"Quá nhiều địa chỉ CC bổ sung ({len(additional_cc)}). Tối đa 20.",
                }
            try:
                cc_list = _validate_email_list(additional_cc, field_name="additional_cc")
            except ValidationError as e:
                audit.log_blocked(
                    tool="reply_to_email",
                    block_reason=f"validation_error_additional_cc: {e}",
                    risk_level="low",
                    duration_ms=0,
                )
                return {"status": "error", "error": f"Địa chỉ CC bổ sung không hợp lệ: {e}"}

        # Bước 5: Lấy danh sách thư mục được phép từ cấu hình
        allowed_folders: list[str] = list(
            getattr(getattr(config, 'security', None), 'allowed_folders', []) or []
        )
        if not allowed_folders:
            audit.log_blocked(
                tool="reply_to_email",
                block_reason="allowed_folders_empty",
                risk_level="high",
                duration_ms=0,
            )
            return {
                "status": "error",
                "error": (
                    "Cấu hình allowed_folders rỗng — không thể xác minh thư mục chứa email. "
                    "Vui lòng kiểm tra config.toml."
                ),
            }

        # Bước 6: Ghi audit log bắt đầu
        import hashlib, time
        start_ts = time.monotonic()
        entry_id_prefix = entry_id_clean[:8]  # Chỉ ghi 8 ký tự đầu
        action_type = "reply_all" if reply_all else "reply"
        audit.log_start(
            tool="reply_to_email",
            params={
                "entry_id_prefix": entry_id_prefix,
                "action_type": action_type,
                "has_additional_cc": cc_list is not None and len(cc_list) > 0,
            },
        )

        # Bước 7: Thực thi COM trong STA thread
        try:
            reply_entry_id = _run_in_sta(
                outlook_com_bridge.sta_executor,
                _com_open_reply,
                entry_id_clean,
                body_clean,
                allowed_folders,
                reply_all,
                cc_list,
            )
        except OutlookNotRunningError as e:
            duration_ms = int((time.monotonic() - start_ts) * 1000)
            audit.log_error(
                tool="reply_to_email",
                error="outlook_not_running",
                duration_ms=duration_ms,
            )
            return {"status": "error", "error": str(e)}
        except FolderNotAllowedError as e:
            duration_ms = int((time.monotonic() - start_ts) * 1000)
            audit.log_blocked(
                tool="reply_to_email",
                block_reason=f"folder_not_allowed: {e}",
                risk_level="high",
                duration_ms=duration_ms,
            )
            return {
                "status": "error",
                "error": (
                    "Email gốc không nằm trong thư mục được phép truy cập. "
                    "Chỉ có thể trả lời email trong các thư mục đã cấu hình."
                ),
            }
        except OutlookOperationError as e:
            duration_ms = int((time.monotonic() - start_ts) * 1000)
            audit.log_error(
                tool="reply_to_email",
                error="com_operation_failed",
                duration_ms=duration_ms,
            )
            return {"status": "error", "error": str(e)}
        except Exception:
            duration_ms = int((time.monotonic() - start_ts) * 1000)
            audit.log_error(
                tool="reply_to_email",
                error="internal_error",
                duration_ms=duration_ms,
            )
            return {"status": "error", "error": "Lỗi nội bộ. Kiểm tra log để biết thêm chi tiết."}

        # Bước 8: Ghi audit log thành công
        duration_ms = int((time.monotonic() - start_ts) * 1000)
        audit.log_success(tool="reply_to_email", duration_ms=duration_ms)

        return {
            "status": "opened",
            "message": (
                "Đã mở cửa sổ trả lời email trong Outlook. "
                "Vui lòng kiểm tra nội dung và nhấn Send để gửi."
            ),
            "reply_entry_id": reply_entry_id,
            "action_type": action_type,
        }

    # ----------------------------------------------------------
    # Tool 3: forward_email — Chuyển tiếp một email
    # ----------------------------------------------------------
    @server.tool(
        name="forward_email",
        description=(
            "Chuyển tiếp một email trong Outlook Desktop đến người nhận mới. "
            "Outlook sẽ mở cửa sổ forward — bạn phải tự nhấn Send để gửi. "
            "Tool này KHÔNG tự động gửi email."
        ),
    )
    async def tool_forward_email(
        entry_id: str,
        to: str | list[str],
        note: str = "",
    ) -> dict[str, str]:
        """
        Mở cửa sổ chuyển tiếp email trong Outlook.

        Tham số:
            entry_id -- ID hex của email cần forward
            to       -- địa chỉ người nhận forward (chuỗi hoặc danh sách)
            note     -- ghi chú thêm của người dùng (tùy chọn, tối đa 50 000 ký tự)

        Trả về dict với status và message hướng dẫn.
        """
        # Bước 1: Kiểm tra read_only_mode
        if getattr(getattr(config, 'security', None), 'read_only_mode', True):
            audit.log_blocked(
                tool="forward_email",
                block_reason="read_only_mode",
                risk_level="low",
                duration_ms=0,
            )
            return {
                "status": "blocked",
                "error": (
                    "Chế độ chỉ đọc (read_only_mode) đang bật. "
                    "Để forward email, hãy đặt read_only_mode = false trong config.toml."
                ),
            }

        # Bước 2: Validate entry_id
        try:
            entry_id_clean = _validate_entry_id(entry_id)
        except ValidationError as e:
            audit.log_blocked(
                tool="forward_email",
                block_reason=f"validation_error_entry_id: {e}",
                risk_level="medium",
                duration_ms=0,
            )
            return {"status": "error", "error": f"entry_id không hợp lệ: {e}"}

        # Bước 3: Validate địa chỉ người nhận forward
        try:
            to_list = _validate_email_list(to, field_name="to")
        except ValidationError as e:
            audit.log_blocked(
                tool="forward_email",
                block_reason=f"validation_error_to: {e}",
                risk_level="low",
                duration_ms=0,
            )
            return {"status": "error", "error": f"Địa chỉ người nhận không hợp lệ: {e}"}

        # Bước 4: Kiểm tra số lượng người nhận
        max_recipients: int = getattr(
            getattr(config, 'security', None),
            'max_recipients_per_draft',
            _DEFAULT_MAX_RECIPIENTS,
        )
        if len(to_list) > max_recipients:
            audit.log_blocked(
                tool="forward_email",
                block_reason=f"too_many_recipients: {len(to_list)} > {max_recipients}",
                risk_level="medium",
                duration_ms=0,
            )
            return {
                "status": "error",
                "error": (
                    f"Quá nhiều người nhận ({len(to_list)}). "
                    f"Tối đa cho phép là {max_recipients} địa chỉ."
                ),
            }

        # Bước 5: Validate note nếu có
        note_clean = ""
        if note:
            try:
                note_clean = _validate_body(note)
            except ValidationError as e:
                audit.log_blocked(
                    tool="forward_email",
                    block_reason=f"validation_error_note: {e}",
                    risk_level="low",
                    duration_ms=0,
                )
                return {"status": "error", "error": f"Ghi chú không hợp lệ: {e}"}

        # Bước 6: Lấy danh sách thư mục được phép từ cấu hình
        allowed_folders: list[str] = list(
            getattr(getattr(config, 'security', None), 'allowed_folders', []) or []
        )
        if not allowed_folders:
            audit.log_blocked(
                tool="forward_email",
                block_reason="allowed_folders_empty",
                risk_level="high",
                duration_ms=0,
            )
            return {
                "status": "error",
                "error": (
                    "Cấu hình allowed_folders rỗng — không thể xác minh thư mục chứa email. "
                    "Vui lòng kiểm tra config.toml."
                ),
            }

        # Bước 7: Ghi audit log bắt đầu
        import hashlib, time
        start_ts = time.monotonic()
        entry_id_prefix = entry_id_clean[:8]
        to_hash = hashlib.sha256(";".join(to_list).encode("utf-8")).hexdigest()
        audit.log_start(
            tool="forward_email",
            params={
                "entry_id_prefix": entry_id_prefix,
                "to_hash": f"sha256:{to_hash[:16]}",
                "recipient_count": len(to_list),
                "has_note": bool(note_clean),
            },
        )

        # Bước 8: Thực thi COM trong STA thread
        try:
            forward_entry_id = _run_in_sta(
                outlook_com_bridge.sta_executor,
                _com_open_forward,
                entry_id_clean,
                to_list,
                allowed_folders,
                note_clean,
            )
        except OutlookNotRunningError as e:
            duration_ms = int((time.monotonic() - start_ts) * 1000)
            audit.log_error(
                tool="forward_email",
                error="outlook_not_running",
                duration_ms=duration_ms,
            )
            return {"status": "error", "error": str(e)}
        except FolderNotAllowedError as e:
            duration_ms = int((time.monotonic() - start_ts) * 1000)
            audit.log_blocked(
                tool="forward_email",
                block_reason=f"folder_not_allowed: {e}",
                risk_level="high",
                duration_ms=duration_ms,
            )
            return {
                "status": "error",
                "error": (
                    "Email gốc không nằm trong thư mục được phép truy cập. "
                    "Chỉ có thể chuyển tiếp email trong các thư mục đã cấu hình."
                ),
            }
        except OutlookOperationError as e:
            duration_ms = int((time.monotonic() - start_ts) * 1000)
            audit.log_error(
                tool="forward_email",
                error="com_operation_failed",
                duration_ms=duration_ms,
            )
            return {"status": "error", "error": str(e)}
        except Exception:
            duration_ms = int((time.monotonic() - start_ts) * 1000)
            audit.log_error(
                tool="forward_email",
                error="internal_error",
                duration_ms=duration_ms,
            )
            return {"status": "error", "error": "Lỗi nội bộ. Kiểm tra log để biết thêm chi tiết."}

        # Bước 9: Ghi audit log thành công
        duration_ms = int((time.monotonic() - start_ts) * 1000)
        audit.log_success(tool="forward_email", duration_ms=duration_ms)

        return {
            "status": "opened",
            "message": (
                "Đã mở cửa sổ chuyển tiếp email trong Outlook. "
                "Vui lòng kiểm tra nội dung và nhấn Send để gửi."
            ),
            "forward_entry_id": forward_entry_id,
        }


# ============================================================
# Hàm dispatch đồng bộ cho server.py
# ============================================================

def _get_config_read_only(config) -> bool:
    """Đọc read_only_mode từ config, hỗ trợ cả hai dạng attribute."""
    # Dạng mới: Config dataclass với READ_ONLY_MODE uppercase
    v = getattr(config, "READ_ONLY_MODE", None)
    if v is not None:
        return bool(v)
    # Dạng cũ: config.security.read_only_mode
    sec = getattr(config, "security", None)
    if sec is not None:
        v2 = getattr(sec, "read_only_mode", None)
        if v2 is not None:
            return bool(v2)
    return True  # fail-safe: mặc định là read-only nếu không đọc được


def _get_config_max_recipients(config) -> int:
    """Đọc max_recipients từ config, hỗ trợ cả hai dạng attribute."""
    v = getattr(config, "MAX_RECIPIENTS_PER_DRAFT", None)
    if v is not None:
        return int(v)
    sec = getattr(config, "security", None)
    if sec is not None:
        v2 = getattr(sec, "max_recipients_per_draft", None)
        if v2 is not None:
            return int(v2)
    return _DEFAULT_MAX_RECIPIENTS


def handle_compose_draft(arguments: dict, config, com_bridge) -> dict:
    """
    Wrapper đồng bộ mở cửa sổ soạn email mới cho server.py dispatch.

    Chạy trong STA thread executor của server.py — đã CoInitialize().
    Gọi trực tiếp _com_open_compose() thay vì qua _run_in_sta().

    KHÔNG bao giờ gửi email — chỉ mở cửa sổ Outlook cho người dùng xem lại.

    Tham số:
        arguments  -- dict tham số từ Claude
        config     -- Config object
        com_bridge -- OutlookCOMBridge (không dùng trực tiếp — compose gọi COM trực tiếp)

    Trả về:
        dict {"status": "opened", "message": str, "draft_entry_id": str}
        hoặc {"status": "blocked", "error": str}
        hoặc {"status": "error", "error": str}

    LƯU Ý FORMAT — có chủ ý khác các tool read-only:
        Compose và reply tools luôn trả về key "status" cùng với key "error" khi thất bại.
        Điều này cho phép caller phân biệt 3 trạng thái rõ ràng:
          - "opened"  → thành công, Outlook đã mở cửa sổ soạn thảo
          - "blocked" → bị từ chối bởi policy (read_only_mode, vượt giới hạn người nhận...)
          - "error"   → lỗi kỹ thuật (Outlook không chạy, COM lỗi, dữ liệu không hợp lệ)
        Key "error" luôn tồn tại khi status không phải "opened", key "message" chỉ có khi "opened".
    """
    # Bước 1: Kiểm tra read_only_mode — từ chối ngay nếu bật
    if _get_config_read_only(config):
        return {
            "status": "blocked",
            "error": "Server đang ở chế độ chỉ đọc (read_only_mode=True). Không thể soạn email.",
        }

    # Bước 2: Validate địa chỉ người nhận (To)
    raw_to = arguments.get("to", "")
    try:
        to_list = _validate_email_list(raw_to, "to")
    except ValidationError as exc:
        return {"status": "error", "error": f"Địa chỉ người nhận không hợp lệ: {exc}"}

    # Bước 3: Kiểm tra giới hạn số người nhận
    max_recipients = _get_config_max_recipients(config)
    cc_raw = arguments.get("cc") or []
    try:
        cc_list = _validate_email_list(cc_raw, "cc") if cc_raw else []
    except ValidationError as exc:
        return {"status": "error", "error": f"Địa chỉ CC không hợp lệ: {exc}"}

    # Bước 3b: Validate BCC — người nhận chính không thấy danh sách BCC
    bcc_raw = arguments.get("bcc") or []
    try:
        bcc_list = _validate_email_list(bcc_raw, "bcc") if bcc_raw else []
    except ValidationError as exc:
        return {"status": "error", "error": f"Địa chỉ BCC không hợp lệ: {exc}"}

    if len(to_list) + len(cc_list) + len(bcc_list) > max_recipients:
        return {
            "status": "blocked",
            "error": f"Vượt quá giới hạn {max_recipients} người nhận cho mỗi email.",
        }

    # Bước 4: Validate subject và body
    try:
        subject = _validate_subject(arguments.get("subject", ""))
    except ValidationError as exc:
        return {"status": "error", "error": f"Tiêu đề email không hợp lệ: {exc}"}

    try:
        body = _validate_body(arguments.get("body", ""))
    except ValidationError as exc:
        return {"status": "error", "error": f"Nội dung email không hợp lệ: {exc}"}

    # Bước 5: Validate importance
    importance_str = str(arguments.get("importance", "normal")).lower().strip()
    importance_map = {"low": 0, "normal": 1, "high": 2}
    importance_value = importance_map.get(importance_str, 1)

    # Bước 6: Gọi COM function trong thread hiện tại (đã là STA thread của server.py)
    try:
        draft_entry_id = _com_open_compose(
            to_addresses=to_list,
            subject=subject,
            body=body,
            cc_addresses=cc_list or None,
            importance_value=importance_value,
            bcc_addresses=bcc_list or None,
        )
        return {
            "status": "opened",
            "message": "Đã mở cửa sổ soạn email trong Outlook. Vui lòng kiểm tra và nhấn Send để gửi.",
            "draft_entry_id": draft_entry_id,
        }
    except OutlookNotRunningError as exc:
        return {"status": "error", "error": str(exc)}
    except OutlookOperationError as exc:
        return {"status": "error", "error": str(exc)}
    except Exception as exc:
        return {"status": "error", "error": f"Lỗi nội bộ khi mở cửa sổ soạn email: {type(exc).__name__}"}


def handle_reply_draft(arguments: dict, config, com_bridge) -> dict:
    """
    Wrapper đồng bộ mở cửa sổ trả lời email cho server.py dispatch.

    Chạy trong STA thread executor của server.py — đã CoInitialize().
    Gọi trực tiếp _com_open_reply() thay vì qua _run_in_sta().

    Kiểm tra bảo mật IDOR (Insecure Direct Object Reference — truy cập qua ID
    trực tiếp mà không kiểm tra quyền): xác minh email gốc nằm trong allowed_folders
    TRƯỚC khi mở reply, ngăn lộ nội dung email bí mật qua reply body.

    KHÔNG bao giờ gửi email — chỉ mở cửa sổ Outlook cho người dùng xem lại.

    Tham số:
        arguments  -- dict tham số từ Claude
        config     -- Config object (có ALLOWED_FOLDERS, READ_ONLY_MODE)
        com_bridge -- OutlookCOMBridge (không dùng trực tiếp — reply gọi COM trực tiếp)

    Trả về:
        dict {"status": "opened", "message": str, "reply_entry_id": str}
        hoặc {"status": "blocked", "error": str}
        hoặc {"status": "error", "error": str}

    LƯU Ý FORMAT — có chủ ý khác các tool read-only:
        Reply tool luôn trả về key "status" cùng với key "error" khi thất bại.
        Điều này cho phép caller phân biệt 3 trạng thái rõ ràng:
          - "opened"  → thành công, Outlook đã mở cửa sổ trả lời
          - "blocked" → bị từ chối bởi policy (read_only_mode, folder nằm ngoài allowlist...)
          - "error"   → lỗi kỹ thuật (Outlook không chạy, COM lỗi, entry_id không hợp lệ)
        Key "error" luôn tồn tại khi status không phải "opened", key "message" chỉ có khi "opened".
    """
    # Bước 1: Kiểm tra read_only_mode
    if _get_config_read_only(config):
        return {
            "status": "blocked",
            "error": "Server đang ở chế độ chỉ đọc (read_only_mode=True). Không thể trả lời email.",
        }

    # Bước 2: Validate entry_id
    entry_id_raw = (arguments.get("entry_id") or "").strip()
    try:
        entry_id = _validate_entry_id(entry_id_raw)
    except ValidationError as exc:
        return {"status": "error", "error": f"entry_id không hợp lệ: {exc}"}

    # Bước 3: Validate body
    try:
        body = _validate_body(arguments.get("body", ""))
    except ValidationError as exc:
        return {"status": "error", "error": f"Nội dung email không hợp lệ: {exc}"}

    # Bước 4: reply_all và CC bổ sung
    reply_all = bool(arguments.get("reply_all", False))
    cc_raw = arguments.get("additional_cc") or []
    try:
        additional_cc = _validate_email_list(cc_raw, "additional_cc") if cc_raw else None
    except ValidationError as exc:
        return {"status": "error", "error": f"Địa chỉ CC không hợp lệ: {exc}"}

    # Lấy allowed_folders từ config
    allowed_folders = list(getattr(config, "ALLOWED_FOLDERS", []) or [])
    if not allowed_folders:
        # Fallback sang config.security.allowed_folders
        sec = getattr(config, "security", None)
        if sec:
            allowed_folders = list(getattr(sec, "allowed_folders", []) or [])

    # Bước 5: Gọi COM function (đã là STA thread của server.py)
    try:
        reply_entry_id = _com_open_reply(
            entry_id=entry_id,
            body=body,
            allowed_folders=allowed_folders,
            reply_all=reply_all,
            additional_cc=additional_cc,
        )
        action_type = "reply_all" if reply_all else "reply"
        return {
            "status": "opened",
            "message": "Đã mở cửa sổ trả lời email trong Outlook. Vui lòng kiểm tra và nhấn Send để gửi.",
            "reply_entry_id": reply_entry_id,
            "action_type": action_type,
        }
    except OutlookNotRunningError as exc:
        return {"status": "error", "error": str(exc)}
    except FolderNotAllowedError as exc:
        return {"status": "blocked", "error": str(exc)}
    except OutlookOperationError as exc:
        return {"status": "error", "error": str(exc)}
    except Exception as exc:
        return {"status": "error", "error": f"Lỗi nội bộ khi mở cửa sổ trả lời: {type(exc).__name__}"}


def handle_forward_draft(arguments: dict, config, com_bridge) -> dict:
    """
    Wrapper đồng bộ mở cửa sổ chuyển tiếp (forward) email cho server.py dispatch.

    Chạy trong STA thread executor của server.py — đã CoInitialize().

    KHÔNG bao giờ gửi email — chỉ mở cửa sổ Outlook cho người dùng xem lại.

    Tham số arguments:
        entry_id (str):   ID hex của email gốc cần forward
        to (list[str]):   Danh sách địa chỉ email người nhận forward
        note (str):       Ghi chú thêm của người dùng trước email gốc (tùy chọn)

    Trả về:
        dict {"status": "opened", "message": str, "forward_entry_id": str}
        hoặc {"status": "blocked", "error": str}
        hoặc {"status": "error", "error": str}
    """
    # Bước 1: Kiểm tra read_only_mode
    if _get_config_read_only(config):
        return {
            "status": "blocked",
            "error": "Server đang ở chế độ chỉ đọc (read_only_mode=True). Không thể forward email.",
        }

    # Bước 2: Validate entry_id
    entry_id_raw = str(arguments.get("entry_id", "")).strip()
    try:
        entry_id = _validate_entry_id(entry_id_raw)
    except ValidationError as exc:
        return {"status": "error", "error": f"entry_id không hợp lệ: {exc}"}

    # Bước 3: Validate danh sách người nhận forward (To)
    raw_to = arguments.get("to", "")
    try:
        to_list = _validate_email_list(raw_to, "to")
    except ValidationError as exc:
        return {"status": "error", "error": f"Địa chỉ người nhận không hợp lệ: {exc}"}

    max_recipients = _get_config_max_recipients(config)
    if len(to_list) > max_recipients:
        return {
            "status": "blocked",
            "error": f"Vượt quá giới hạn {max_recipients} người nhận.",
        }

    # Bước 4: Validate note (ghi chú tùy chọn)
    note = str(arguments.get("note", "")).strip()
    if len(note) > _MAX_BODY_LENGTH:
        note = note[:_MAX_BODY_LENGTH]

    # Bước 5: Lấy allowed_folders từ config
    allowed_folders: list[str] = list(
        getattr(config, "ALLOWED_FOLDERS", None) or []
    )
    if not allowed_folders:
        sec = getattr(config, "security", None)
        if sec:
            allowed_folders = list(getattr(sec, "allowed_folders", []) or [])

    # Bước 6: Gọi COM function để mở cửa sổ forward
    try:
        forward_entry_id = _com_open_forward(
            entry_id=entry_id,
            to_addresses=to_list,
            allowed_folders=allowed_folders,
            note=note,
        )
        return {
            "status": "opened",
            "message": "Đã mở cửa sổ chuyển tiếp email trong Outlook. Vui lòng kiểm tra và nhấn Send để gửi.",
            "forward_entry_id": forward_entry_id,
        }
    except OutlookNotRunningError as exc:
        return {"status": "error", "error": str(exc)}
    except FolderNotAllowedError as exc:
        return {"status": "blocked", "error": str(exc)}
    except OutlookOperationError as exc:
        return {"status": "error", "error": str(exc)}
    except Exception as exc:
        return {"status": "error", "error": f"Lỗi nội bộ khi mở cửa sổ forward: {type(exc).__name__}"}
