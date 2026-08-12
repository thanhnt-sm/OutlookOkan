# ============================================================
# setup.ps1 — Cài đặt môi trường Claude-Outlook MCP Secure
# ============================================================
# Script này KHÔNG cần quyền admin.
# Chạy từ thư mục chứa file này:
#   cd <thư mục dự án>
#   .\setup.ps1
# ============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Màu sắc hiển thị để dễ đọc
function Write-Step  { param($Msg) Write-Host "`n[BƯỚC] $Msg" -ForegroundColor Cyan }
function Write-Ok    { param($Msg) Write-Host "  [OK] $Msg"   -ForegroundColor Green }
function Write-Warn  { param($Msg) Write-Host "  [!]  $Msg"   -ForegroundColor Yellow }
function Write-Fail  { param($Msg) Write-Host "  [LỖI] $Msg"  -ForegroundColor Red; exit 1 }

# ============================================================
# Bước 1: Kiểm tra phiên bản Python — yêu cầu 3.11 trở lên
# ============================================================
Write-Step "Kiểm tra phiên bản Python..."

$PythonExe = $null
foreach ($candidate in @('python', 'python3', 'py')) {
    try {
        $ver = & $candidate --version 2>&1
        if ($ver -match 'Python (\d+)\.(\d+)') {
            $major = [int]$Matches[1]
            $minor = [int]$Matches[2]
            if ($major -gt 3 -or ($major -eq 3 -and $minor -ge 11)) {
                $PythonExe = $candidate
                Write-Ok "Tìm thấy $ver — đạt yêu cầu (>= 3.11)."
                break
            } else {
                Write-Warn "$candidate phiên bản $ver quá cũ (cần >= 3.11)."
            }
        }
    } catch {
        # Lệnh không tồn tại, thử lệnh kế tiếp
    }
}

if (-not $PythonExe) {
    Write-Fail "Không tìm thấy Python 3.11+ trong PATH. Tải tại https://python.org"
}

# ============================================================
# Bước 2: Tạo virtual environment tại .\venv\
# ============================================================
Write-Step "Tạo virtual environment tại .\venv\ ..."

$VenvDir = Join-Path $PSScriptRoot 'venv'
if (Test-Path $VenvDir) {
    Write-Warn "Thư mục .\venv\ đã tồn tại — bỏ qua bước tạo mới."
} else {
    & $PythonExe -m venv $VenvDir
    if ($LASTEXITCODE -ne 0) { Write-Fail "Không thể tạo venv." }
    Write-Ok "Đã tạo venv tại $VenvDir"
}

$VenvPython = Join-Path $VenvDir 'Scripts\python.exe'
$VenvPip    = Join-Path $VenvDir 'Scripts\pip.exe'

if (-not (Test-Path $VenvPython)) {
    Write-Fail "Không tìm thấy $VenvPython sau khi tạo venv."
}

# ============================================================
# Bước 3: Nâng cấp pip trong venv lên phiên bản mới nhất
# ============================================================
Write-Step "Nâng cấp pip..."
& $VenvPython -m pip install --upgrade pip --quiet
Write-Ok "pip đã được nâng cấp."

# ============================================================
# Bước 4: Cài đặt tất cả thư viện từ requirements.txt
# ============================================================
Write-Step "Cài đặt thư viện từ requirements.txt..."

$ReqFile = Join-Path $PSScriptRoot 'requirements.txt'
if (-not (Test-Path $ReqFile)) {
    Write-Fail "Không tìm thấy requirements.txt tại $ReqFile"
}

& $VenvPip install -r $ReqFile
if ($LASTEXITCODE -ne 0) { Write-Fail "pip install thất bại." }
Write-Ok "Tất cả thư viện đã được cài đặt."

# ============================================================
# Bước 5: Chạy pywin32 post-install (bắt buộc cho win32com)
# ============================================================
Write-Step "Chạy pywin32 post-install (cần để win32com hoạt động)..."

# Tìm file pywin32_postinstall.py trong Scripts của venv
$PostInstall = Join-Path $VenvDir 'Scripts\pywin32_postinstall.py'
if (Test-Path $PostInstall) {
    & $VenvPython $PostInstall -install
    if ($LASTEXITCODE -ne 0) {
        Write-Warn "pywin32 post-install trả về lỗi — có thể bỏ qua nếu đã cài trước đó."
    } else {
        Write-Ok "pywin32 post-install hoàn tất."
    }
} else {
    Write-Warn "Không tìm thấy pywin32_postinstall.py — bỏ qua bước này."
}

# ============================================================
# Bước 6: Hỏi người dùng có muốn cài đặt credentials ngay không
# ============================================================
Write-Step "Cài đặt thông tin xác thực (API key)..."
Write-Host ""
Write-Host "  Bước này sẽ lưu Anthropic API key vào Windows Credential Manager." -ForegroundColor White
Write-Host "  Key KHÔNG được lưu vào file — hoàn toàn bảo mật." -ForegroundColor White
Write-Host ""

$answer = Read-Host "  Bạn muốn cài đặt credentials ngay bây giờ không? [y/N]"
if ($answer -match '^[Yy]$') {
    $ServerPy = Join-Path $PSScriptRoot 'server.py'
    if (-not (Test-Path $ServerPy)) {
        Write-Warn "Không tìm thấy server.py — bỏ qua bước setup credentials."
    } else {
        Write-Host "  Đang mở trình hướng dẫn cài đặt credentials..." -ForegroundColor White
        & $VenvPython $ServerPy --setup
        if ($LASTEXITCODE -ne 0) {
            Write-Warn "Setup credentials kết thúc với mã lỗi $LASTEXITCODE."
        } else {
            Write-Ok "Credentials đã được lưu vào Windows Credential Manager."
        }
    }
} else {
    Write-Warn "Bỏ qua cài đặt credentials. Chạy lại sau: .\venv\Scripts\python.exe server.py --setup"
}

# ============================================================
# Bước 7: Hiển thị hướng dẫn bước tiếp theo
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  CÀI ĐẶT HOÀN TẤT — BƯỚC TIẾP THEO" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  1. Đăng ký MCP server với Claude Code (chọn 1 trong 2 cách):" -ForegroundColor White
Write-Host ""
Write-Host "     Cách A — Dùng lệnh Claude CLI (khuyến nghị):" -ForegroundColor Yellow
Write-Host '       claude mcp add outlook -- .\venv\Scripts\python.exe server.py' -ForegroundColor Gray
Write-Host ""
Write-Host "     Cách B — Chỉnh tay file config của Claude Code:" -ForegroundColor Yellow
Write-Host '       Mở file claude-mcp.json trong thư mục này,' -ForegroundColor Gray
Write-Host '       thay ABSOLUTE_PATH_TO_THIS_DIR bằng đường dẫn thực tế,' -ForegroundColor Gray
Write-Host "       rồi merge nội dung vào `~\.claude\config.json`" -ForegroundColor Gray
Write-Host ""
Write-Host "  2. Mở Outlook Desktop TRƯỚC KHI dùng Claude." -ForegroundColor White
Write-Host "     (Server kết nối Outlook đang chạy — không tự khởi động Outlook.)" -ForegroundColor Gray
Write-Host ""
Write-Host "  3. Kiểm tra cấu hình trong config.toml:" -ForegroundColor White
Write-Host "     - allowed_folders: danh sách thư mục Claude được phép đọc" -ForegroundColor Gray
Write-Host "     - read_only_mode: true (mặc định an toàn)" -ForegroundColor Gray
Write-Host ""
Write-Host "  4. Tài liệu đầy đủ: .\docs\USER_GUIDE.md" -ForegroundColor White
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
