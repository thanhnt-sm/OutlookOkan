# Test toàn diện 6 MCP tools — chạy trực tiếp, không qua MCP protocol
# Mục đích: kiểm tra từng tool hoạt động đúng với Outlook Desktop thực
#
# Chạy: .\venv\Scripts\python.exe test_all_tools.py

import sys
import json
import traceback
import time

sys.path.insert(0, ".")

# ── Màu sắc terminal ──────────────────────────────────────────────────────────
GREEN  = "\033[92m"
RED    = "\033[91m"
YELLOW = "\033[93m"
CYAN   = "\033[96m"
RESET  = "\033[0m"
BOLD   = "\033[1m"

def ok(msg):  print(f"  {GREEN}✔{RESET} {msg}")
def fail(msg): print(f"  {RED}✘{RESET} {msg}")
def warn(msg): print(f"  {YELLOW}⚠{RESET} {msg}")
def info(msg): print(f"  {CYAN}→{RESET} {msg}")

results = []  # (tool_name, passed, message)

def record(tool, passed, msg):
    results.append((tool, passed, msg))
    if passed:
        ok(f"{BOLD}{tool}{RESET}: {msg}")
    else:
        fail(f"{BOLD}{tool}{RESET}: {msg}")

# ── Load config + bridge ──────────────────────────────────────────────────────
print(f"\n{BOLD}=== KHỞI TẠO ==={RESET}")
try:
    from config import load_config
    cfg = load_config()
    info(f"Config: {cfg.ACCOUNT_NAME}, {len(cfg.ALLOWED_FOLDERS)} thư mục được phép")
    ok("load_config")
except Exception as e:
    fail(f"load_config: {e}")
    sys.exit(1)

try:
    from outlook_com import OutlookCOMBridge
    bridge = OutlookCOMBridge(config=cfg)
    ok("OutlookCOMBridge")
except Exception as e:
    fail(f"OutlookCOMBridge: {e}")
    sys.exit(1)

try:
    from security.validator import InputValidator
    validator = InputValidator(cfg)
    ok("InputValidator")
except Exception as e:
    fail(f"InputValidator: {e}")
    sys.exit(1)

# ── TEST 1: list_folders ──────────────────────────────────────────────────────
print(f"\n{BOLD}=== TEST 1: list_folders ==={RESET}")
try:
    from tools.list_folders import handle_list_folders
    result = handle_list_folders({"include_subfolders": False}, cfg, bridge)
    if "error" in result:
        record("list_folders", False, f"lỗi: {result['error']}")
    else:
        folders = result.get("folders", [])
        info(f"Tìm thấy {len(folders)} thư mục:")
        for f in folders:
            info(f"  • {f['name']}: {f['total_count']} mail, {f['unread_count']} chưa đọc")
        record("list_folders", True, f"{len(folders)} thư mục trả về")
except Exception as e:
    record("list_folders", False, traceback.format_exc().splitlines()[-1])

# ── TEST 2: list_emails ───────────────────────────────────────────────────────
print(f"\n{BOLD}=== TEST 2: list_emails ==={RESET}")
first_entry_id = None
try:
    from tools.read_email import handle_list_emails
    result = handle_list_emails(
        {"folder_name": "Inbox", "max_count": 5, "unread_only": False},
        cfg, bridge
    )
    if "error" in result:
        # Thử thư mục tiếng Việt nếu Inbox không tìm thấy
        result = handle_list_emails(
            {"folder_name": "Hộp thư đến", "max_count": 5, "unread_only": False},
            cfg, bridge
        )
    if "error" in result:
        record("list_emails", False, f"lỗi: {result['error']}")
    else:
        emails = result.get("emails", [])
        info(f"Tìm thấy {len(emails)} email:")
        for e in emails[:3]:
            info(f"  • [{e.get('date','?')}] {e.get('subject','(không có tiêu đề)')[:60]}")
            if first_entry_id is None:
                first_entry_id = e.get("entry_id")
        record("list_emails", True, f"{len(emails)} email trả về")
except Exception as e:
    record("list_emails", False, traceback.format_exc().splitlines()[-1])

# ── TEST 3: read_email ────────────────────────────────────────────────────────
print(f"\n{BOLD}=== TEST 3: read_email ==={RESET}")
if first_entry_id:
    try:
        from tools.read_email import handle_read_email
        result = handle_read_email(
            {"entry_id": first_entry_id},
            cfg, bridge
        )
        if "error" in result:
            record("read_email", False, f"lỗi: {result['error']}")
        else:
            subj = result.get("subject", "?")[:60]
            body_len = len(result.get("body", ""))
            info(f"Subject: {subj}")
            info(f"Body: {body_len} ký tự")
            info(f"Folder: {result.get('folder_name', '?')}")
            record("read_email", True, f"đọc OK, body={body_len} chars")
    except Exception as e:
        record("read_email", False, traceback.format_exc().splitlines()[-1])
else:
    warn("read_email: bỏ qua — không có entry_id từ list_emails")
    results.append(("read_email", None, "bỏ qua (không có entry_id)"))

# ── TEST 4: search_emails ─────────────────────────────────────────────────────
print(f"\n{BOLD}=== TEST 4: search_emails ==={RESET}")
try:
    from tools.search import handle_search_emails
    result = handle_search_emails(
        {"query": "test", "folder_name": "Inbox", "max_count": 3},
        cfg, bridge
    )
    if "error" in result:
        # Thử thư mục khác
        result = handle_search_emails(
            {"query": "test", "folder_name": "Hộp thư đến", "max_count": 3},
            cfg, bridge
        )
    if "error" in result:
        record("search_emails", False, f"lỗi: {result['error']}")
    else:
        emails = result.get("emails", [])
        info(f"Kết quả tìm kiếm 'test': {len(emails)} email")
        record("search_emails", True, f"{len(emails)} kết quả trả về")
except Exception as e:
    record("search_emails", False, traceback.format_exc().splitlines()[-1])

# ── TEST 5: compose_draft ─────────────────────────────────────────────────────
print(f"\n{BOLD}=== TEST 5: compose_draft ==={RESET}")
print(f"  {YELLOW}[Outlook sẽ mở cửa sổ soạn thảo — đóng lại sau khi kiểm tra]{RESET}")
try:
    from tools.compose import handle_compose_draft
    result = handle_compose_draft(
        {
            "to": ["thanhnt.sm@gmail.com"],
            "subject": "[MCP TEST] Kiểm tra soạn email tự động",
            "body": "Đây là email kiểm tra từ Claude MCP Outlook.\n\nEmail này được soạn tự động — KHÔNG gửi đi, chỉ xem draft.",
        },
        cfg, bridge
    )
    status = result.get("status", "unknown")
    if status == "opened":
        record("compose_draft", True, "cửa sổ soạn thảo đã mở trong Outlook")
    elif status == "blocked":
        warn(f"compose_draft: bị chặn vì READ_ONLY_MODE=true")
        results.append(("compose_draft", None, "blocked by read_only_mode"))
    else:
        record("compose_draft", False, f"status={status}, {result.get('error','?')}")
except Exception as e:
    record("compose_draft", False, traceback.format_exc().splitlines()[-1])

# ── TEST 6: reply_draft ───────────────────────────────────────────────────────
print(f"\n{BOLD}=== TEST 6: reply_draft ==={RESET}")
if first_entry_id:
    print(f"  {YELLOW}[Outlook sẽ mở cửa sổ trả lời — đóng lại sau khi kiểm tra]{RESET}")
    try:
        from tools.compose import handle_reply_draft
        result = handle_reply_draft(
            {
                "entry_id": first_entry_id,
                "body": "[MCP TEST] Đây là nội dung trả lời kiểm tra. KHÔNG gửi đi.",
                "reply_all": False,
            },
            cfg, bridge
        )
        status = result.get("status", "unknown")
        if status == "opened":
            record("reply_draft", True, "cửa sổ trả lời đã mở trong Outlook")
        elif status == "blocked":
            warn("reply_draft: bị chặn vì READ_ONLY_MODE=true")
            results.append(("reply_draft", None, "blocked by read_only_mode"))
        else:
            record("reply_draft", False, f"status={status}, {result.get('error','?')}")
    except Exception as e:
        record("reply_draft", False, traceback.format_exc().splitlines()[-1])
else:
    warn("reply_draft: bỏ qua — không có entry_id từ list_emails")
    results.append(("reply_draft", None, "bỏ qua (không có entry_id)"))

# ── TỔNG KẾT ─────────────────────────────────────────────────────────────────
print(f"\n{BOLD}{'='*50}")
print(f"KẾT QUẢ KIỂM TRA")
print(f"{'='*50}{RESET}")

passed  = [r for r in results if r[1] is True]
failed  = [r for r in results if r[1] is False]
skipped = [r for r in results if r[1] is None]

print(f"{GREEN}Thành công: {len(passed)}/{len(results)}{RESET}")
if failed:
    print(f"{RED}Thất bại:   {len(failed)}/{len(results)}{RESET}")
    for name, _, msg in failed:
        print(f"  {RED}✘ {name}: {msg}{RESET}")
if skipped:
    print(f"{YELLOW}Bỏ qua:     {len(skipped)}/{len(results)}{RESET}")

if not failed:
    print(f"\n{GREEN}{BOLD}✔ Hệ thống sẵn sàng sử dụng!{RESET}")
else:
    print(f"\n{RED}{BOLD}✘ Có {len(failed)} tool chưa hoạt động — xem chi tiết ở trên.{RESET}")
