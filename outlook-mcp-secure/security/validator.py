# Kiểm tra và làm sạch dữ liệu đầu vào cho Claude-Outlook MCP Secure
# Mọi tham số từ Claude (MCP client) phải đi qua validator này trước khi xử lý
# Ngăn chặn: path traversal, injection, Unicode trick, giá trị vượt giới hạn

from __future__ import annotations

import logging
import re
import unicodedata
from datetime import date, datetime
from typing import List

# Logger nội bộ cho validator
_logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Hằng số giới hạn — khớp với config.toml trong PLAN.md
# ---------------------------------------------------------------------------

# Độ dài tối đa cho từng loại đầu vào
_MAX_FOLDER_NAME_LENGTH = 260
_MAX_ENTRY_ID_LENGTH = 256
_MAX_SEARCH_QUERY_LENGTH = 200
_MAX_SUBJECT_LENGTH = 500
_MAX_BODY_LENGTH = 50_000
_MAX_EMAIL_ADDRESS_LENGTH = 254  # RFC 5321

# Regex kiểm tra entry_id: chỉ cho phép ký tự hex (chữ số và chữ cái A-F)
_ENTRY_ID_REGEX = re.compile(r"^[0-9A-Fa-f]+$")

# Regex kiểm tra địa chỉ email cơ bản (RFC 5322 đơn giản hóa)
# Đủ dùng cho mục đích bảo mật — không cần chuẩn RFC hoàn toàn
_EMAIL_REGEX = re.compile(
    r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"
)

# Ký tự nguy hiểm trong search query: SQL/DASL injection, shell metacharacters
# Danh sách này dựa trên PLAN.md section "Prompt Injection — Các Vector Phải Chặn"
_DANGEROUS_PATTERNS = [
    "--",           # SQL comment injection
    ";",            # Command separator
    "<script",      # XSS injection
    "exec(",        # Code execution
    "eval(",        # Code execution
    "\x00",         # Null byte
    "../",          # Path traversal
    "..\\",         # Path traversal (Windows)
    "://",          # URL scheme injection
    ":\\",          # Windows absolute path
]

# Ký tự điều khiển Unicode ẩn — dùng để trick AI hoặc bypass filter
_INVISIBLE_UNICODE_CHARS = [
    "​",  # Zero-width space
    "‌",  # Zero-width non-joiner
    "‍",  # Zero-width joiner
    "‮",  # Right-to-left override
    " ",  # Line separator
    " ",  # Paragraph separator
    "﻿",  # Byte order mark
]


class ValidationError(ValueError):
    """
    Lỗi khi dữ liệu đầu vào không hợp lệ.
    Kế thừa ValueError để dễ bắt trong handler chung.
    Message luôn bằng tiếng Việt để hiển thị cho người dùng.
    """
    pass


def _normalize_string(text: str) -> str:
    """
    Chuẩn hóa chuỗi theo đúng thứ tự bắt buộc trong PLAN.md:
    1. strip() — loại bỏ whitespace đầu cuối
    2. NFC normalize — chuẩn hóa Unicode tổ hợp ký tự
    3. casefold() — lowercase unicode-aware

    Dùng nội bộ để so sánh tên folder với allowlist.
    """
    return unicodedata.normalize("NFC", text.strip()).casefold()


def _check_control_chars(text: str, field_name: str) -> None:
    """
    Kiểm tra chuỗi không chứa ký tự điều khiển (control characters) và null bytes.
    Theo PLAN.md: reject ngay nếu phát hiện \x00 hoặc \x01-\x1F.

    Tham số:
        text       — chuỗi cần kiểm tra
        field_name — tên trường để hiển thị trong thông báo lỗi
    """
    # Null byte là vector tấn công path traversal và string termination
    if "\x00" in text:
        raise ValidationError(f"'{field_name}' chứa ký tự null byte không hợp lệ.")

    # Control characters (0x01 - 0x1F) ngoại trừ tab và newline thông thường
    for char in text:
        code = ord(char)
        if 0x01 <= code <= 0x1F and char not in ("\t", "\n", "\r"):
            raise ValidationError(
                f"'{field_name}' chứa ký tự điều khiển (control character) "
                f"không hợp lệ (U+{code:04X})."
            )


def _check_invisible_unicode(text: str, field_name: str) -> None:
    """
    Kiểm tra chuỗi không chứa các ký tự Unicode vô hình.
    Đây là vector tấn công prompt injection tinh vi — ký tự hiển thị bình thường
    nhưng có thể thay đổi ngữ nghĩa khi AI xử lý.

    Tham số:
        text       — chuỗi cần kiểm tra
        field_name — tên trường để hiển thị trong thông báo lỗi
    """
    for invisible_char in _INVISIBLE_UNICODE_CHARS:
        if invisible_char in text:
            raise ValidationError(
                f"'{field_name}' chứa ký tự Unicode ẩn không hợp lệ "
                f"(U+{ord(invisible_char):04X})."
            )


class InputValidator:
    """
    Kiểm tra toàn bộ đầu vào từ Claude (MCP client) trước khi xử lý.

    Mỗi phương thức:
    - Nhận giá trị thô từ MCP request
    - Kiểm tra độ dài, ký tự nguy hiểm, định dạng
    - Trả về giá trị đã được làm sạch (sanitized), hoặc raise ValidationError

    Không bao giờ log giá trị đầu vào đầy đủ — chỉ log metadata (tên trường, độ dài).
    """

    def __init__(self, config=None) -> None:
        """
        Khởi tạo validator với config tùy chọn.

        Tham số:
            config -- Config object từ config.py (có ALLOWED_FOLDERS, ENTRY_ID_MAX_LENGTH...)
                      Có thể là None khi dùng độc lập (cần truyền allowed_list cho validate_folder).
        """
        self._config = config

    # ── Alias methods — tương thích với cách gọi trong tool modules ──────────

    def validate_folder_name(self, name: str) -> str:
        """
        Alias tiện lợi cho validate_folder() — tự lấy allowed_list từ config.

        Tham số:
            name -- Tên folder từ client

        Trả về:
            Tên folder đã strip whitespace nếu hợp lệ và trong allowlist.

        Raise:
            ValidationError -- Nếu folder không hợp lệ hoặc không trong allowlist
        """
        # Lấy danh sách folder được phép từ config — mặc định rỗng nếu không có
        allowed: List[str] = []
        if self._config is not None:
            # Hỗ trợ cả config.ALLOWED_FOLDERS (Config dataclass) lẫn config.security.allowed_folders
            allowed = (
                getattr(self._config, "ALLOWED_FOLDERS", None)
                or getattr(getattr(self._config, "security", None), "allowed_folders", None)
                or []
            )

        if not allowed:
            raise ValidationError(
                "Danh sách thư mục được phép (allowed_folders) chưa được cấu hình. "
                "Vui lòng kiểm tra config.toml."
            )

        return self.validate_folder(name, allowed)

    def validate_email_id(self, eid: str) -> str:
        """
        Alias tiện lợi cho validate_entry_id() — tên quen thuộc hơn trong tool modules.

        Tham số:
            eid -- Entry ID cần kiểm tra

        Trả về:
            entry_id đã strip whitespace, đảm bảo an toàn.

        Raise:
            ValidationError -- nếu format sai hoặc quá dài
        """
        return self.validate_entry_id(eid)

    # ── Core validation methods ───────────────────────────────────────────────

    def validate_folder(self, name: str, allowed_list: List[str]) -> str:
        """
        Kiểm tra tên folder có hợp lệ và thuộc danh sách được phép không.

        Thuật toán theo PLAN.md section 6 (Folder Allowlist Validation):
        1. strip() whitespace
        2. NFC normalize
        3. casefold() để so sánh không phân biệt hoa/thường
        4. Kiểm tra null bytes và control chars
        5. Kiểm tra path traversal
        6. So sánh với allowed_list (sau khi normalize cả hai phía)

        Tham số:
            name         — tên folder từ client (ví dụ: "Inbox", "Inbox/Projects")
            allowed_list — danh sách tên folder được phép trong config.toml

        Trả về:
            Tên folder đã strip whitespace (giữ nguyên case gốc cho COM lookup).

        Raise:
            ValidationError — nếu folder không hợp lệ hoặc không trong allowlist
        """
        # Bước 1: Kiểm tra đầu vào cơ bản
        if not isinstance(name, str):
            raise ValidationError("Tên folder phải là chuỗi ký tự.")
        if not name or not name.strip():
            raise ValidationError("Tên folder không được để trống.")

        # Bước 2: Strip whitespace (giữ lại cho việc trả về)
        stripped_name = name.strip()

        # Bước 3: Kiểm tra độ dài
        if len(stripped_name) > _MAX_FOLDER_NAME_LENGTH:
            raise ValidationError(
                f"Tên folder quá dài: {len(stripped_name)} ký tự "
                f"(tối đa {_MAX_FOLDER_NAME_LENGTH})."
            )

        # Bước 4: Kiểm tra null bytes và control characters
        _check_control_chars(stripped_name, "tên folder")

        # Bước 5: Kiểm tra path traversal và ký tự nguy hiểm
        lower_name = stripped_name.lower()
        if "../" in lower_name or "..\\" in lower_name:
            raise ValidationError("Tên folder chứa path traversal '../' không hợp lệ.")
        if "://" in lower_name:
            raise ValidationError("Tên folder chứa URL scheme '://' không hợp lệ.")
        if ":\\" in stripped_name or stripped_name.startswith("/"):
            raise ValidationError("Tên folder không được là đường dẫn tuyệt đối.")

        # Bước 6: Chuẩn hóa để so sánh với allowlist
        normalized_input = _normalize_string(stripped_name)

        # Bước 7: So sánh với từng entry trong allowlist (cả hai phía đều normalized)
        for allowed in allowed_list:
            if _normalize_string(allowed) == normalized_input:
                _logger.debug("Folder hợp lệ: tên='%s', độ_dài=%d.", stripped_name, len(stripped_name))
                return stripped_name  # Trả về tên gốc đã strip — COM cần đúng case

        # Không tìm thấy trong allowlist
        _logger.warning(
            "Folder bị từ chối: không thuộc allowlist. Độ dài tên=%d.", len(stripped_name)
        )
        raise ValidationError(
            "Thư mục không được phép truy cập. "
            "Chỉ các thư mục trong danh sách cấu hình mới được phép."
        )

    def validate_entry_id(self, eid: str) -> str:
        """
        Kiểm tra entry_id hợp lệ: chỉ chứa ký tự hex, tối đa 256 ký tự.

        Entry ID là định danh nội bộ của Outlook cho mỗi email/item trong PST.
        Format: chuỗi hex dài (128-256 ký tự) do Outlook tạo ra.

        Tham số:
            eid — entry ID từ client

        Trả về:
            entry_id đã strip whitespace, đảm bảo an toàn.

        Raise:
            ValidationError — nếu format sai hoặc quá dài
        """
        # Bước 1: Kiểm tra kiểu dữ liệu
        if not isinstance(eid, str):
            raise ValidationError("Entry ID phải là chuỗi ký tự.")
        if not eid or not eid.strip():
            raise ValidationError("Entry ID không được để trống.")

        # Bước 2: Strip whitespace
        stripped_eid = eid.strip()

        # Bước 3: Kiểm tra độ dài (theo PLAN.md: max 256 chars)
        if len(stripped_eid) > _MAX_ENTRY_ID_LENGTH:
            raise ValidationError(
                f"Entry ID quá dài: {len(stripped_eid)} ký tự "
                f"(tối đa {_MAX_ENTRY_ID_LENGTH})."
            )

        # Bước 4: Kiểm tra format hex — chỉ cho phép [0-9A-Fa-f]
        if not _ENTRY_ID_REGEX.match(stripped_eid):
            raise ValidationError(
                "Entry ID không hợp lệ: chỉ được chứa ký tự hex (0-9, A-F)."
            )

        _logger.debug("Entry ID hợp lệ: độ_dài=%d.", len(stripped_eid))
        return stripped_eid

    def validate_search_query(self, q: str) -> str:
        """
        Kiểm tra và làm sạch câu truy vấn tìm kiếm.

        Loại bỏ các ký tự có thể dùng để tấn công DASL injection,
        SQL injection, hoặc prompt injection.

        Tham số:
            q — câu truy vấn từ client

        Trả về:
            Câu truy vấn đã được làm sạch, an toàn để dùng với DASL filter.

        Raise:
            ValidationError — nếu query rỗng, quá dài, hoặc chứa pattern nguy hiểm
        """
        # Bước 1: Kiểm tra kiểu dữ liệu
        if not isinstance(q, str):
            raise ValidationError("Câu truy vấn phải là chuỗi ký tự.")
        if not q or not q.strip():
            raise ValidationError("Câu truy vấn tìm kiếm không được để trống.")

        # Bước 2: Strip whitespace
        stripped_q = q.strip()

        # Bước 3: Kiểm tra độ dài (theo PLAN.md: max 200 chars)
        if len(stripped_q) > _MAX_SEARCH_QUERY_LENGTH:
            raise ValidationError(
                f"Câu truy vấn quá dài: {len(stripped_q)} ký tự "
                f"(tối đa {_MAX_SEARCH_QUERY_LENGTH})."
            )

        # Bước 4: Kiểm tra null bytes và control characters
        _check_control_chars(stripped_q, "câu truy vấn")

        # Bước 5: Kiểm tra ký tự Unicode vô hình (prompt injection)
        _check_invisible_unicode(stripped_q, "câu truy vấn")

        # Bước 6: Kiểm tra các pattern nguy hiểm
        lower_q = stripped_q.lower()
        for dangerous in _DANGEROUS_PATTERNS:
            if dangerous in lower_q:
                _logger.warning(
                    "Query bị từ chối: phát hiện pattern nguy hiểm. Độ dài=%d.", len(stripped_q)
                )
                raise ValidationError(
                    f"Câu truy vấn chứa ký tự hoặc cú pháp không được phép: '{dangerous}'."
                )

        # Bước 7: Escape dấu nháy đơn để ngăn DASL injection
        # DASL filter string dùng dấu nháy đơn bao quanh giá trị, ví dụ:
        #   "@SQL=""urn:schemas:httpmail:subject"" LIKE '%từ_tìm%'"
        # Nếu query chứa dấu ' thì chuỗi filter bị cắt ngang, kẻ tấn công có thể
        # chèn mệnh đề DASL tùy ý (DASL injection). Cách khắc phục tiêu chuẩn là
        # nhân đôi dấu nháy đơn: ' → '' (tương tự SQL string escaping).
        # Dấu nháy kép không cần escape vì DASL dùng dấu nháy đơn làm delimiter.
        safe_q = stripped_q.replace("'", "''")

        _logger.debug("Search query hợp lệ: độ_dài=%d.", len(stripped_q))
        return safe_q

    def validate_email(self, email_address: str) -> str:
        """
        Kiểm tra địa chỉ email có đúng định dạng không.

        Dùng regex RFC 5322 đơn giản hóa — đủ bảo mật cho mục đích validate,
        không cần phức tạp như thư viện email-validator đầy đủ.

        Tham số:
            email_address — địa chỉ email cần kiểm tra

        Trả về:
            Địa chỉ email đã strip, lowercase.

        Raise:
            ValidationError — nếu định dạng không hợp lệ
        """
        # Bước 1: Kiểm tra kiểu dữ liệu
        if not isinstance(email_address, str):
            raise ValidationError("Địa chỉ email phải là chuỗi ký tự.")
        if not email_address or not email_address.strip():
            raise ValidationError("Địa chỉ email không được để trống.")

        # Bước 2: Strip và lowercase
        cleaned = email_address.strip().lower()

        # Bước 3: Kiểm tra độ dài (RFC 5321: max 254 ký tự)
        if len(cleaned) > _MAX_EMAIL_ADDRESS_LENGTH:
            raise ValidationError(
                f"Địa chỉ email quá dài: {len(cleaned)} ký tự "
                f"(tối đa {_MAX_EMAIL_ADDRESS_LENGTH})."
            )

        # Bước 4: Kiểm tra null bytes
        _check_control_chars(cleaned, "địa chỉ email")

        # Bước 5: Kiểm tra định dạng bằng regex
        if not _EMAIL_REGEX.match(cleaned):
            raise ValidationError(
                f"Địa chỉ email không đúng định dạng: '{cleaned}'."
            )

        _logger.debug("Email hợp lệ: độ_dài=%d.", len(cleaned))
        return cleaned

    def validate_date(self, d: str | None) -> datetime | None:
        """
        Phân tích chuỗi ngày tháng theo định dạng ISO 8601 (YYYY-MM-DD).

        Tham số:
            d — chuỗi ngày tháng, hoặc None nếu không bắt buộc

        Trả về:
            datetime object tương ứng (giờ = 00:00:00), hoặc None nếu d là None.

        Raise:
            ValidationError — nếu chuỗi không đúng định dạng ISO 8601
        """
        # Bước 1: Cho phép None (trường không bắt buộc)
        if d is None:
            return None

        # Bước 2: Kiểm tra kiểu dữ liệu
        if not isinstance(d, str):
            raise ValidationError("Ngày tháng phải là chuỗi định dạng YYYY-MM-DD.")

        stripped_d = d.strip()
        if not stripped_d:
            return None

        # Bước 3: Phân tích định dạng ISO 8601 (YYYY-MM-DD)
        try:
            parsed = datetime.strptime(stripped_d, "%Y-%m-%d")
            _logger.debug("Ngày hợp lệ: '%s'.", stripped_d)
            return parsed
        except ValueError:
            raise ValidationError(
                f"Ngày tháng không đúng định dạng: '{stripped_d}'. "
                "Yêu cầu định dạng YYYY-MM-DD (ví dụ: 2026-06-24)."
            )

    def validate_int(self, v: int | str, min_v: int, max_v: int) -> int:
        """
        Kiểm tra giá trị nguyên nằm trong khoảng [min_v, max_v].

        Tham số:
            v     — giá trị cần kiểm tra (có thể là int hoặc string từ JSON)
            min_v — giá trị tối thiểu (inclusive)
            max_v — giá trị tối đa (inclusive)

        Trả về:
            Giá trị nguyên đã được kiểm tra.

        Raise:
            ValidationError — nếu không phải số nguyên hoặc nằm ngoài khoảng cho phép
        """
        # Bước 1: Chuyển đổi nếu cần
        if isinstance(v, str):
            try:
                v = int(v.strip())
            except (ValueError, AttributeError):
                raise ValidationError(f"Giá trị '{v}' không phải số nguyên hợp lệ.")

        # Bước 2: Kiểm tra kiểu
        if not isinstance(v, int) or isinstance(v, bool):
            raise ValidationError(f"Giá trị phải là số nguyên, nhận được: {type(v).__name__}.")

        # Bước 3: Kiểm tra khoảng giá trị
        if v < min_v:
            raise ValidationError(
                f"Giá trị {v} nhỏ hơn giới hạn tối thiểu {min_v}."
            )
        if v > max_v:
            raise ValidationError(
                f"Giá trị {v} lớn hơn giới hạn tối đa {max_v}."
            )

        return v

    def validate_subject(self, subject: str) -> str:
        """
        Kiểm tra tiêu đề email: không rỗng, tối đa 500 ký tự, không chứa ký tự nguy hiểm.

        Tham số:
            subject — tiêu đề email từ client

        Trả về:
            Tiêu đề đã strip whitespace.

        Raise:
            ValidationError — nếu tiêu đề không hợp lệ
        """
        if not isinstance(subject, str):
            raise ValidationError("Tiêu đề email phải là chuỗi ký tự.")
        if not subject or not subject.strip():
            raise ValidationError("Tiêu đề email không được để trống.")

        stripped = subject.strip()

        if len(stripped) > _MAX_SUBJECT_LENGTH:
            raise ValidationError(
                f"Tiêu đề quá dài: {len(stripped)} ký tự (tối đa {_MAX_SUBJECT_LENGTH})."
            )

        # Kiểm tra null bytes (không strip control chars vì subject có thể có tab)
        if "\x00" in stripped:
            raise ValidationError("Tiêu đề chứa ký tự null byte không hợp lệ.")

        _check_invisible_unicode(stripped, "tiêu đề email")

        return stripped

    def validate_body(self, body: str) -> str:
        """
        Kiểm tra nội dung email: tối đa 50,000 ký tự, không chứa null byte.

        Tham số:
            body — nội dung email từ client

        Trả về:
            Nội dung đã được kiểm tra.

        Raise:
            ValidationError — nếu nội dung không hợp lệ
        """
        if not isinstance(body, str):
            raise ValidationError("Nội dung email phải là chuỗi ký tự.")

        if len(body) > _MAX_BODY_LENGTH:
            raise ValidationError(
                f"Nội dung email quá dài: {len(body)} ký tự (tối đa {_MAX_BODY_LENGTH})."
            )

        # Chỉ kiểm tra null byte — không reject newlines hay control chars khác trong body
        if "\x00" in body:
            raise ValidationError("Nội dung email chứa ký tự null byte không hợp lệ.")

        return body
