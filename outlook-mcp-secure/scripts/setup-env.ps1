#Requires -Version 5.1
<#
.SYNOPSIS
    Thiết lập biến môi trường OUTLOOK_MCP_AUDIT_KEY cho hệ thống Outlook MCP Secure.

.DESCRIPTION
    Script này tự động sinh một khóa bảo mật ngẫu nhiên 32 ký tự (256-bit entropy)
    và lưu vào biến môi trường cấp User. Khóa này được dùng để tính HMAC-SHA256
    cho từng entry trong audit log, giúp phát hiện nếu ai đó chỉnh sửa log sau khi ghi.

    Quy trình:
      Bước 1 — Kiểm tra xem khóa đã tồn tại chưa để tránh ghi đè vô ý.
      Bước 2 — Sinh khóa ngẫu nhiên 32 byte bằng RNGCryptoServiceProvider (an toàn mật mã).
      Bước 3 — Mã hóa sang Base64 URL-safe (loại bỏ +/= để tránh lỗi shell).
      Bước 4 — Lưu vào User Environment (chỉ áp dụng cho user hiện tại, không cần quyền Admin).
      Bước 5 — Verify: đọc lại từ registry để xác nhận đã lưu thành công.

.NOTES
    Phiên bản : 1.0.0
    Tác giả   : OutlookOkan Team
    Tên biến  : OUTLOOK_MCP_AUDIT_KEY
    Yêu cầu   : PowerShell 5.1+, Windows
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---- Màu sắc hiển thị ----
$Green  = "Green"
$Yellow = "Yellow"
$Red    = "Red"
$Cyan   = "Cyan"

# ---- Tên biến môi trường ----
$EnvVarName = "OUTLOOK_MCP_AUDIT_KEY"

Write-Host ""
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "  OUTLOOK MCP SECURE — Setup Audit Key  " -ForegroundColor $Cyan
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host ""

# ------------------------------------------------------------------
# Bước 1: Kiểm tra xem khóa đã tồn tại chưa
# ------------------------------------------------------------------
Write-Host "[Bước 1] Kiểm tra biến môi trường hiện tại..." -ForegroundColor $Yellow

$ExistingKey = [System.Environment]::GetEnvironmentVariable($EnvVarName, "User")

if ($ExistingKey -and $ExistingKey.Length -gt 0) {
    Write-Host "  CANH BAO: Bien $EnvVarName da ton tai (do dai: $($ExistingKey.Length) ky tu)." -ForegroundColor $Yellow
    Write-Host ""
    $Confirm = Read-Host "  Ban co muon ghi de khoa cu khong? (go 'YES' de xac nhan, Enter de huy)"
    if ($Confirm -ne "YES") {
        Write-Host ""
        Write-Host "  Da huy. Khoa cu van duoc giu nguyen." -ForegroundColor $Green
        Write-Host ""
        exit 0
    }
    Write-Host "  Dang ghi de khoa cu..." -ForegroundColor $Yellow
} else {
    Write-Host "  OK — Bien chua ton tai, se tao moi." -ForegroundColor $Green
}

# ------------------------------------------------------------------
# Bước 2: Sinh khóa ngẫu nhiên 32 byte (256-bit entropy)
# Dùng RNGCryptoServiceProvider — an toàn về mặt mật mã học (CSPRNG)
# ------------------------------------------------------------------
Write-Host ""
Write-Host "[Bước 2] Sinh khóa ngẫu nhiên bảo mật..." -ForegroundColor $Yellow

try {
    # Tạo mảng 32 byte ngẫu nhiên bằng CSPRNG (Cryptographically Secure Pseudo-Random Number Generator)
    $RngProvider = [System.Security.Cryptography.RNGCryptoServiceProvider]::new()
    $RandomBytes  = New-Object byte[] 32
    $RngProvider.GetBytes($RandomBytes)
    $RngProvider.Dispose()
} catch {
    Write-Host "  LOI: Khong the sinh random bytes: $_" -ForegroundColor $Red
    exit 1
}

# ------------------------------------------------------------------
# Bước 3: Mã hóa sang Base64 URL-safe (tránh ký tự +/= gây lỗi shell)
# Thay + thành -, / thành _, bỏ dấu = cuối
# Độ dài kết quả: 43 ký tự (32 byte Base64 URL-safe)
# ------------------------------------------------------------------
Write-Host ""
Write-Host "[Bước 3] Mã hóa khóa sang dạng Base64 URL-safe..." -ForegroundColor $Yellow

$Base64Raw  = [System.Convert]::ToBase64String($RandomBytes)
$NewKey     = $Base64Raw.Replace("+", "-").Replace("/", "_").TrimEnd("=")

Write-Host "  Khóa mới (đầu 8 ký tự): $($NewKey.Substring(0,8))..." -ForegroundColor $Green
Write-Host "  Độ dài: $($NewKey.Length) ký tự" -ForegroundColor $Green

# ------------------------------------------------------------------
# Bước 4: Lưu vào User Environment Variable
# "User" scope → lưu vào HKCU:\Environment, không cần quyền Admin
# Session hiện tại cũng được cập nhật ngay lập tức
# ------------------------------------------------------------------
Write-Host ""
Write-Host "[Bước 4] Lưu vào User Environment Variable..." -ForegroundColor $Yellow

try {
    [System.Environment]::SetEnvironmentVariable($EnvVarName, $NewKey, "User")
    # Cập nhật luôn cho session PowerShell hiện tại
    [System.Environment]::SetEnvironmentVariable($EnvVarName, $NewKey, "Process")
    Write-Host "  OK — Da luu vao User Environment." -ForegroundColor $Green
} catch {
    Write-Host "  LOI: Khong the luu bien moi truong: $_" -ForegroundColor $Red
    exit 1
}

# ------------------------------------------------------------------
# Bước 5: Verify — Đọc lại từ registry để xác nhận lưu thành công
# Đọc từ "User" scope (registry HKCU:\Environment) — không phải Process
# ------------------------------------------------------------------
Write-Host ""
Write-Host "[Bước 5] Xác minh đã lưu thành công..." -ForegroundColor $Yellow

$VerifiedKey = [System.Environment]::GetEnvironmentVariable($EnvVarName, "User")

if ($null -eq $VerifiedKey -or $VerifiedKey.Length -eq 0) {
    Write-Host "  LOI: Bien moi truong khong doc lai duoc sau khi luu!" -ForegroundColor $Red
    exit 1
}

if ($VerifiedKey -ne $NewKey) {
    Write-Host "  LOI: Gia tri doc lai khong khop voi gia tri vua luu!" -ForegroundColor $Red
    Write-Host "    Da luu  : $($NewKey.Substring(0,8))..." -ForegroundColor $Red
    Write-Host "    Doc lai : $($VerifiedKey.Substring(0,8))..." -ForegroundColor $Red
    exit 1
}

Write-Host "  PASS — Gia tri xac nhan khop." -ForegroundColor $Green

# ------------------------------------------------------------------
# Tóm tắt kết quả
# ------------------------------------------------------------------
Write-Host ""
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "  KET QUA                                " -ForegroundColor $Cyan
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host ""
Write-Host "  Bien : $EnvVarName" -ForegroundColor $Green
Write-Host "  Do dai: $($NewKey.Length) ky tu (256-bit entropy)" -ForegroundColor $Green
Write-Host "  Scope: User (chi user hien tai, khong can Admin)" -ForegroundColor $Green
Write-Host ""
Write-Host "  LUU Y QUAN TRONG:" -ForegroundColor $Yellow
Write-Host "    - Phai khoi dong lai terminal/IDE de bien co hieu luc o session moi." -ForegroundColor $Yellow
Write-Host "    - Khoa nay duoc dung de ky HMAC cho audit log." -ForegroundColor $Yellow
Write-Host "    - Neu doi khoa, cac entry log cu se bao loi integrity khi verify." -ForegroundColor $Yellow
Write-Host "    - Backup khoa o noi an toan truoc khi doi." -ForegroundColor $Yellow
Write-Host ""
Write-Host "  Hoan tat!" -ForegroundColor $Green
Write-Host ""
