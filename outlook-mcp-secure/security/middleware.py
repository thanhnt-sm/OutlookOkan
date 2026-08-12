"""
security/middleware.py — Security middleware layer cho MCP tool handlers.

Tập trung hóa các cross-cutting security concerns:
  - Giới hạn độ sâu/kích thước của input params
  - Kiểm tra an toàn folder name
  - Sanitize response trước khi trả về Claude

Cách dùng trong handle_X functions:
    result = sanitize_tool_response(raw_result)
"""

from __future__ import annotations

import logging
import unicodedata
from typing import Any

_logger = logging.getLogger(__name__)

# Giới hạn độ sâu tối đa của dict/list trong response (tránh JSON bomb)
_MAX_RESPONSE_DEPTH = 6
# Giới hạn số lượng items trong list response
_MAX_RESPONSE_LIST_ITEMS = 600
# Giới hạn độ dài string trong response
_MAX_RESPONSE_STRING_LENGTH = 20_000


def sanitize_tool_response(data: Any, _depth: int = 0) -> Any:
    """
    Sanitize dữ liệu trả về từ tool handler trước khi gửi cho Claude.

    Bước 1: Kiểm tra độ sâu — truncate nếu quá sâu
    Bước 2: Với dict — sanitize từng value đệ quy
    Bước 3: Với list — giới hạn số items, sanitize từng item
    Bước 4: Với string — cắt nếu quá dài
    Bước 5: Với types khác — trả về nguyên
    """
    if _depth > _MAX_RESPONSE_DEPTH:
        return "[truncated: quá sâu]"

    if isinstance(data, dict):
        return {
            k: sanitize_tool_response(v, _depth + 1)
            for k, v in data.items()
        }
    elif isinstance(data, list):
        truncated = data[:_MAX_RESPONSE_LIST_ITEMS]
        result = [sanitize_tool_response(item, _depth + 1) for item in truncated]
        if len(data) > _MAX_RESPONSE_LIST_ITEMS:
            result.append(f"[truncated: {len(data) - _MAX_RESPONSE_LIST_ITEMS} items bị lược bỏ]")
        return result
    elif isinstance(data, str):
        if len(data) > _MAX_RESPONSE_STRING_LENGTH:
            return data[:_MAX_RESPONSE_STRING_LENGTH] + f"...[truncated {len(data) - _MAX_RESPONSE_STRING_LENGTH} chars]"
        return data
    else:
        return data


def check_folder_name_safe(name: str) -> bool:
    """
    Kiểm tra nhanh tên folder có an toàn không.
    Trả về True nếu an toàn, False nếu phát hiện pattern nguy hiểm.
    Không raise — dùng như guard clause trong tool handlers.
    """
    if not name or not isinstance(name, str):
        return False
    if len(name) > 260:
        return False
    # Null bytes, control characters
    if any(ord(c) < 0x20 for c in name):
        return False
    # Path traversal
    normalized = unicodedata.normalize("NFC", name)
    if ".." in normalized or "://" in normalized or ":\\" in normalized:
        return False
    return True


def validate_input_size(data: Any, max_depth: int = 4, max_list_len: int = 100) -> bool:
    """
    Kiểm tra input từ Claude có vượt quá giới hạn kích thước không.
    Ngăn chặn JSON bomb qua tool arguments.
    """
    def _check(obj: Any, depth: int) -> bool:
        if depth > max_depth:
            return False
        if isinstance(obj, dict):
            if len(obj) > 50:
                return False
            return all(_check(v, depth + 1) for v in obj.values())
        elif isinstance(obj, list):
            if len(obj) > max_list_len:
                return False
            return all(_check(item, depth + 1) for item in obj)
        elif isinstance(obj, str):
            return len(obj) <= 100_000
        return True

    return _check(data, 0)
