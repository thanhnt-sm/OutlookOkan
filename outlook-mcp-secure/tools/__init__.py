"""
Module tools — Đăng ký tất cả MCP tools cho Claude-Outlook MCP Secure server.

Mỗi sub-module (list_folders, read_email, search, compose) đăng ký một nhóm tools
liên quan. File này tổng hợp toàn bộ qua hàm register_all_tools().

Cách dùng từ server.py:
    from tools import register_all_tools
    register_all_tools(server, outlook_com_bridge, audit_logger, config)
"""

from .list_folders import register_tools as register_folder_tools
from .read_email import register_tools as register_read_tools
from .search import register_tools as register_search_tools


def register_all_tools(server, outlook_com, audit, config) -> None:
    """
    Đăng ký tất cả MCP tools vào server.

    Tham số:
        server     -- MCP Server instance (từ mcp.server.Server)
        outlook_com -- OutlookCOMBridge instance (từ outlook_com.py)
        audit      -- AuditLogger instance (từ security/audit.py)
        config     -- Config object đã validate (từ config.py)
    """
    # Đăng ký nhóm tools quản lý thư mục
    register_folder_tools(server, outlook_com, audit, config)

    # Đăng ký nhóm tools đọc email
    register_read_tools(server, outlook_com, audit, config)

    # Đăng ký nhóm tools tìm kiếm email
    register_search_tools(server, outlook_com, audit, config)

    # Đăng ký nhóm tools soạn thảo / trả lời email
    try:
        from .compose import register_tools as register_compose_tools
        register_compose_tools(server, outlook_com, audit, config)
    except ImportError:
        pass

    # Đăng ký nhóm tools calendar — list_calendar_events, create_calendar_event
    try:
        from .calendar import register_tools as register_calendar_tools
        register_calendar_tools(server, outlook_com, audit, config)
    except ImportError:
        pass


__all__ = ["register_all_tools"]
