# Quản lý thông tin xác thực (credentials) cho Claude-Outlook MCP Secure
# Lưu trữ API key và mật khẩu trong Windows Credential Manager (keyring)
# TUYỆT ĐỐI không lưu dưới dạng plain text ra file hoặc biến môi trường
# TUYỆT ĐỐI không ghi giá trị credential vào log

from __future__ import annotations

import logging
import sys

import keyring
import keyring.backend
from keyring.backends.Windows import WinVaultKeyring

# Logger nội bộ — chỉ ghi thông tin meta, không bao giờ ghi giá trị credential
_logger = logging.getLogger(__name__)


def _verify_win_vault_backend() -> None:
    """
    Kiểm tra xem keyring có đang dùng Windows Credential Manager (WinVaultKeyring) không.
    Nếu không đúng backend, raise RuntimeError ngay khi khởi động để tránh lưu credential
    vào nơi không an toàn (ví dụ: file plain text trên Linux).
    Theo yêu cầu SEC-09 trong PLAN.md.
    """
    current_backend = keyring.get_keyring()
    if not isinstance(current_backend, WinVaultKeyring):
        raise RuntimeError(
            f"Backend keyring không hợp lệ: {type(current_backend).__name__}. "
            "Yêu cầu WinVaultKeyring (Windows Credential Manager). "
            "Chạy server này trên Windows với Python keyring >= 23.x."
        )


class CredentialManager:
    """
    Quản lý thông tin xác thực sử dụng Windows Credential Manager.

    Tất cả credentials được lưu dưới service name 'ClaudeOutlookMCP'.
    Không bao giờ log giá trị, chỉ log thao tác (store/get/delete) và trạng thái.

    Sử dụng:
        mgr = CredentialManager()
        mgr.store("anthropic_api_key", "sk-ant-...")
        key = mgr.get("anthropic_api_key")
    """

    # Tên service trong Windows Credential Manager — dùng để nhóm tất cả credentials của ứng dụng này
    SERVICE = "ClaudeOutlookMCP"

    def __init__(self) -> None:
        # Bước 1: Xác minh backend ngay khi khởi tạo — fail fast nếu không an toàn
        _verify_win_vault_backend()
        _logger.debug("CredentialManager khởi tạo thành công với WinVaultKeyring.")

    def store(self, key: str, value: str) -> None:
        """
        Lưu credential vào Windows Credential Manager.

        Tham số:
            key   — tên định danh credential (ví dụ: "anthropic_api_key")
            value — giá trị cần lưu (KHÔNG bao giờ log giá trị này)

        Raise:
            ValueError  — nếu key hoặc value rỗng
            RuntimeError — nếu keyring báo lỗi khi lưu
        """
        # Bước 1: Kiểm tra đầu vào cơ bản
        if not key or not key.strip():
            raise ValueError("Tên credential (key) không được để trống.")
        if not value:
            raise ValueError("Giá trị credential không được để trống.")

        # Bước 2: Lưu vào Windows Credential Manager
        try:
            keyring.set_password(self.SERVICE, key, value)
            # Chỉ ghi tên key vào log, KHÔNG bao giờ ghi value
            _logger.info("Đã lưu credential: key='%s', service='%s'.", key, self.SERVICE)
        except keyring.errors.KeyringError as exc:
            _logger.error("Lỗi khi lưu credential key='%s': %s", key, exc)
            raise RuntimeError(
                f"Không thể lưu credential '{key}' vào Windows Credential Manager."
            ) from exc

    def get(self, key: str) -> str | None:
        """
        Lấy credential từ Windows Credential Manager.

        Tham số:
            key — tên định danh credential (ví dụ: "anthropic_api_key")

        Trả về:
            Giá trị credential dưới dạng string, hoặc None nếu không tìm thấy.

        Raise:
            ValueError  — nếu key rỗng
            RuntimeError — nếu keyring báo lỗi khi đọc
        """
        # Bước 1: Kiểm tra đầu vào
        if not key or not key.strip():
            raise ValueError("Tên credential (key) không được để trống.")

        # Bước 2: Lấy từ Windows Credential Manager
        try:
            value = keyring.get_password(self.SERVICE, key)
            # Chỉ log trạng thái tìm thấy/không tìm thấy — KHÔNG log giá trị
            if value is not None:
                _logger.debug("Lấy credential thành công: key='%s'.", key)
            else:
                _logger.debug("Không tìm thấy credential: key='%s'.", key)
            return value
        except keyring.errors.KeyringError as exc:
            _logger.error("Lỗi khi đọc credential key='%s': %s", key, exc)
            raise RuntimeError(
                f"Không thể đọc credential '{key}' từ Windows Credential Manager."
            ) from exc

    def delete(self, key: str) -> None:
        """
        Xóa credential khỏi Windows Credential Manager.

        Tham số:
            key — tên định danh credential cần xóa

        Raise:
            ValueError  — nếu key rỗng
            RuntimeError — nếu không tìm thấy credential hoặc keyring báo lỗi
        """
        # Bước 1: Kiểm tra đầu vào
        if not key or not key.strip():
            raise ValueError("Tên credential (key) không được để trống.")

        # Bước 2: Xóa khỏi Windows Credential Manager
        try:
            keyring.delete_password(self.SERVICE, key)
            _logger.info("Đã xóa credential: key='%s', service='%s'.", key, self.SERVICE)
        except keyring.errors.PasswordDeleteError as exc:
            _logger.warning("Không tìm thấy credential để xóa: key='%s'. Chi tiết: %s", key, exc)
            raise RuntimeError(
                f"Không tìm thấy credential '{key}' để xóa trong Windows Credential Manager."
            ) from exc
        except keyring.errors.KeyringError as exc:
            _logger.error("Lỗi khi xóa credential key='%s': %s", key, exc)
            raise RuntimeError(
                f"Không thể xóa credential '{key}' khỏi Windows Credential Manager."
            ) from exc

    def setup_wizard(self) -> None:
        """
        Hướng dẫn người dùng thiết lập credentials lần đầu qua giao diện dòng lệnh.

        Người dùng sẽ được hỏi từng credential cần thiết và giá trị sẽ được lưu
        ngay vào Windows Credential Manager. Không hiển thị giá trị sau khi nhập.

        Sử dụng khi chạy lần đầu hoặc khi cần cập nhật credentials.
        """
        print()
        print("=" * 60)
        print("  THIẾT LẬP CREDENTIALS — Claude-Outlook MCP Secure")
        print("=" * 60)
        print()
        print("Chương trình sẽ hỏi bạn các thông tin xác thực cần thiết.")
        print("Thông tin sẽ được lưu an toàn vào Windows Credential Manager.")
        print("TUYỆT ĐỐI KHÔNG lưu vào file hoặc biến môi trường.")
        print()

        # Danh sách credentials cần thiết — (tên_key, mô_tả_cho_người_dùng, có_bắt_buộc_không)
        credentials_to_setup: list[tuple[str, str, bool]] = [
            (
                "anthropic_api_key",
                "Anthropic API Key (bắt đầu bằng 'sk-ant-...')",
                True,
            ),
        ]

        for key, description, is_required in credentials_to_setup:
            print(f"--- {description} ---")

            # Kiểm tra xem credential đã tồn tại chưa
            existing = self.get(key)
            if existing is not None:
                confirm = input(
                    f"Credential '{key}' đã tồn tại. Bạn có muốn ghi đè không? (y/N): "
                ).strip().lower()
                if confirm != "y":
                    print(f"Bỏ qua '{key}'.")
                    print()
                    continue

            # Nhập giá trị — dùng getpass để ẩn khi gõ nếu có thể
            try:
                import getpass
                value = getpass.getpass(f"Nhập {description}: ").strip()
            except Exception:
                # Fallback nếu getpass không hoạt động (ví dụ: môi trường không tương tác)
                value = input(f"Nhập {description} (sẽ hiển thị): ").strip()

            # Bước kiểm tra: không cho lưu giá trị rỗng nếu là bắt buộc
            if not value:
                if is_required:
                    print(f"LỖI: '{key}' là bắt buộc, không thể để trống.")
                    sys.exit(1)
                else:
                    print(f"Bỏ qua '{key}' (không bắt buộc).")
                    print()
                    continue

            # Lưu vào Windows Credential Manager
            try:
                self.store(key, value)
                print(f"Đã lưu '{key}' thành công vào Windows Credential Manager.")
            except RuntimeError as exc:
                print(f"LỖI: Không thể lưu '{key}'. Chi tiết: {exc}")
                if is_required:
                    sys.exit(1)
            print()

        print("=" * 60)
        print("  Thiết lập credentials hoàn tất.")
        print("=" * 60)
        print()
