#Requires -Version 5.1
<#
.SYNOPSIS
    Backup cấu hình và logs của Outlook MCP Secure với timestamp.

.DESCRIPTION
    Script tạo file ZIP backup chứa:
      - config.toml           (cấu hình chính)
      - logs/audit.jsonl      (audit log hiện tại)
      - logs/audit-*.jsonl    (các file log đã rotate)
      - .env (nếu tồn tại)    (biến môi trường local)

    Tên file ZIP: backup-outlook-mcp-YYYYMMDD-HHMMSS.zip
    Thư mục chứa backup: <project>/backups/ (tự tạo nếu chưa có)

    Tùy chọn:
      -BackupDir  : Chỉ định thư mục backup khác
      -MaxBackups : Số backup tối đa giữ lại (mặc định 10, xóa backup cũ nhất)
      -NoLogs     : Chỉ backup config, không backup logs (nếu log quá lớn)

.PARAMETER BackupDir
    Đường dẫn thư mục lưu file backup (mặc định: <project>/backups/)

.PARAMETER MaxBackups
    Số lượng backup tối đa giữ lại trong thư mục (mặc định: 10)

.PARAMETER NoLogs
    Nếu set, chỉ backup config.toml, bỏ qua logs

.EXAMPLE
    .\backup-config.ps1
    .\backup-config.ps1 -BackupDir "D:\MyBackups"
    .\backup-config.ps1 -MaxBackups 5
    .\backup-config.ps1 -NoLogs

.NOTES
    Phiên bản : 1.0.0
    Tác giả   : OutlookOkan Team
    Yêu cầu   : PowerShell 5.1+ (Compress-Archive có sẵn từ PS5+)
#>

[CmdletBinding()]
param(
    [string]$BackupDir  = "",
    [int]$MaxBackups    = 10,
    [switch]$NoLogs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---- Màu sắc hiển thị ----
$Green  = "Green"
$Yellow = "Yellow"
$Red    = "Red"
$Cyan   = "Cyan"
$White  = "White"
$Gray   = "Gray"

# ---- Đường dẫn project ----
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir

# Đường dẫn các file/thư mục cần backup
$ConfigPath  = Join-Path $ProjectDir "config.toml"
$LogDir      = Join-Path $ProjectDir "logs"
$EnvFilePath = Join-Path $ProjectDir ".env"

# Thư mục lưu backup
if (-not $BackupDir) {
    $BackupDir = Join-Path $ProjectDir "backups"
}

# Timestamp cho tên file
$Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$ZipName   = "backup-outlook-mcp-$Timestamp.zip"
$ZipPath   = Join-Path $BackupDir $ZipName

# Thư mục tạm để tập hợp file trước khi nén
$TempDir   = Join-Path $env:TEMP "outlook-mcp-backup-$Timestamp"

# ---- Tiêu đề ----
Write-Host ""
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "  OUTLOOK MCP SECURE — Backup Config    " -ForegroundColor $Cyan
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "  Project  : $ProjectDir" -ForegroundColor $White
Write-Host "  BackupDir: $BackupDir" -ForegroundColor $White
Write-Host "  Output   : $ZipName" -ForegroundColor $White
Write-Host "  Thoi diem: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor $White
Write-Host ""

# ------------------------------------------------------------------
# Bước 1: Tạo thư mục backup nếu chưa tồn tại
# ------------------------------------------------------------------
Write-Host "[Bước 1] Chuẩn bị thư mục backup..." -ForegroundColor $Yellow

try {
    if (-not (Test-Path $BackupDir)) {
        New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
        Write-Host "  Tao thu muc backup: $BackupDir" -ForegroundColor $Green
    } else {
        Write-Host "  Thu muc backup da ton tai: $BackupDir" -ForegroundColor $Green
    }

    # Tạo thư mục tạm để tập hợp file
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
} catch {
    Write-Host "  LOI: Khong the tao thu muc: $_" -ForegroundColor $Red
    exit 1
}

# ------------------------------------------------------------------
# Bước 2: Copy các file cần backup vào thư mục tạm
# ------------------------------------------------------------------
Write-Host ""
Write-Host "[Bước 2] Thu thập file cần backup..." -ForegroundColor $Yellow

$CopiedFiles = [System.Collections.Generic.List[string]]::new()
$SkippedFiles = [System.Collections.Generic.List[string]]::new()

# Hàm copy file an toàn — ghi log kết quả
function Copy-SafeFile {
    param(
        [string]$Source,
        [string]$DestDir,
        [string]$Label
    )
    if (Test-Path $Source) {
        try {
            $DestPath = Join-Path $DestDir (Split-Path -Leaf $Source)
            Copy-Item -Path $Source -Destination $DestPath -Force
            $SizeKB = [Math]::Round((Get-Item $Source).Length / 1KB, 2)
            Write-Host "  [OK] $Label ($SizeKB KB)" -ForegroundColor $Green
            $script:CopiedFiles.Add("$Label ($SizeKB KB)")
        } catch {
            Write-Host "  [LOI] $Label : $_" -ForegroundColor $Red
        }
    } else {
        Write-Host "  [BQ] $Label : khong tim thay, bo qua" -ForegroundColor $Yellow
        $script:SkippedFiles.Add($Label)
    }
}

# Copy config.toml
Copy-SafeFile -Source $ConfigPath -DestDir $TempDir -Label "config.toml"

# Copy .env nếu tồn tại
Copy-SafeFile -Source $EnvFilePath -DestDir $TempDir -Label ".env"

# Copy logs (nếu không có -NoLogs)
if (-not $NoLogs) {
    if (Test-Path $LogDir) {
        # Tạo thư mục logs trong temp
        $TempLogDir = Join-Path $TempDir "logs"
        New-Item -ItemType Directory -Path $TempLogDir -Force | Out-Null

        # Lấy tất cả file .jsonl trong thư mục logs
        $LogFiles = Get-ChildItem -Path $LogDir -Filter "*.jsonl" -File

        if ($LogFiles.Count -eq 0) {
            Write-Host "  [BQ] Khong co file .jsonl nao trong $LogDir" -ForegroundColor $Yellow
        } else {
            foreach ($LogFile in $LogFiles) {
                $DestLogPath = Join-Path $TempLogDir $LogFile.Name
                try {
                    Copy-Item -Path $LogFile.FullName -Destination $DestLogPath -Force
                    $SizeKB = [Math]::Round($LogFile.Length / 1KB, 2)
                    Write-Host "  [OK] logs\$($LogFile.Name) ($SizeKB KB)" -ForegroundColor $Green
                    $CopiedFiles.Add("logs\$($LogFile.Name) ($SizeKB KB)")
                } catch {
                    Write-Host "  [LOI] logs\$($LogFile.Name) : $_" -ForegroundColor $Red
                }
            }
        }
    } else {
        Write-Host "  [BQ] Thu muc logs/ khong ton tai: $LogDir" -ForegroundColor $Yellow
    }
} else {
    Write-Host "  [BQ] Logs bi bo qua theo tuy chon -NoLogs" -ForegroundColor $Yellow
}

# Ghi file README.txt vào backup để ghi lại ngữ cảnh
$ReadmePath = Join-Path $TempDir "BACKUP_INFO.txt"
$ReadmeContent = @"
OUTLOOK MCP SECURE — Backup Info
=================================
Thoi diem backup: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
May tinh         : $env:COMPUTERNAME
User             : $env:USERNAME
Project dir      : $ProjectDir
Script           : backup-config.ps1

Files da backup:
$($CopiedFiles | ForEach-Object { "  - $_" } | Out-String)

Files bo qua:
$($SkippedFiles | ForEach-Object { "  - $_" } | Out-String)

LUU Y:
- File nay khong chua OUTLOOK_MCP_AUDIT_KEY (khoa HMAC)
- Backup rieng khoa HMAC o noi an toan khac
- De phuc hoi: copy config.toml ve thu muc project
"@
$ReadmeContent | Out-File -FilePath $ReadmePath -Encoding UTF8
Write-Host "  [OK] BACKUP_INFO.txt (thong tin backup)" -ForegroundColor $Green

# ------------------------------------------------------------------
# Bước 3: Nén thành file ZIP
# ------------------------------------------------------------------
Write-Host ""
Write-Host "[Bước 3] Nén thành file ZIP..." -ForegroundColor $Yellow

try {
    # Compress-Archive yêu cầu đường dẫn source là thư mục hoặc file
    # Dùng -Path với wildcard để nén toàn bộ nội dung TempDir
    $SourcePath = Join-Path $TempDir "*"
    Compress-Archive -Path $SourcePath -DestinationPath $ZipPath -CompressionLevel Optimal -Force

    $ZipSizeKB = [Math]::Round((Get-Item $ZipPath).Length / 1KB, 2)
    Write-Host "  OK — Tao file ZIP thanh cong: $ZipName ($ZipSizeKB KB)" -ForegroundColor $Green
} catch {
    Write-Host "  LOI: Khong the tao ZIP: $_" -ForegroundColor $Red
    # Dọn dẹp thư mục tạm dù có lỗi
    Remove-Item -Path $TempDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

# ------------------------------------------------------------------
# Bước 4: Dọn dẹp thư mục tạm
# ------------------------------------------------------------------
try {
    Remove-Item -Path $TempDir -Recurse -Force
} catch {
    Write-Host "  CANH BAO: Khong xoa duoc thu muc tam: $TempDir" -ForegroundColor $Yellow
}

# ------------------------------------------------------------------
# Bước 5: Xóa backup cũ nếu vượt quá MaxBackups
# ------------------------------------------------------------------
Write-Host ""
Write-Host "[Bước 5] Kiểm tra số lượng backup (tối đa: $MaxBackups)..." -ForegroundColor $Yellow

try {
    $AllBackups = Get-ChildItem -Path $BackupDir -Filter "backup-outlook-mcp-*.zip" -File |
                  Sort-Object LastWriteTime

    $CurrentCount = $AllBackups.Count
    Write-Host "  Hien co $CurrentCount file backup trong thu muc." -ForegroundColor $White

    if ($CurrentCount -gt $MaxBackups) {
        $ToDelete = $CurrentCount - $MaxBackups
        Write-Host "  Xoa $ToDelete backup cu nhat..." -ForegroundColor $Yellow

        $AllBackups | Select-Object -First $ToDelete | ForEach-Object {
            Write-Host "    Xoa: $($_.Name)" -ForegroundColor $Yellow
            Remove-Item -Path $_.FullName -Force
        }
        Write-Host "  Da xoa $ToDelete backup cu." -ForegroundColor $Green
    } else {
        Write-Host "  Khong can xoa (con lai $CurrentCount / $MaxBackups)." -ForegroundColor $Green
    }
} catch {
    Write-Host "  CANH BAO: Loi khi kiem tra backup cu: $_" -ForegroundColor $Yellow
}

# ------------------------------------------------------------------
# Tóm tắt kết quả
# ------------------------------------------------------------------
Write-Host ""
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "  KET QUA BACKUP" -ForegroundColor $Cyan
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host ""
Write-Host "  File backup : $ZipPath" -ForegroundColor $Green
Write-Host "  Kich thuoc  : $([Math]::Round((Get-Item $ZipPath).Length / 1KB, 2)) KB" -ForegroundColor $Green
Write-Host "  So files    : $($CopiedFiles.Count)" -ForegroundColor $Green
Write-Host ""
Write-Host "  Files da backup:" -ForegroundColor $White
foreach ($F in $CopiedFiles) {
    Write-Host "    - $F" -ForegroundColor $Green
}
if ($SkippedFiles.Count -gt 0) {
    Write-Host "  Files bo qua:" -ForegroundColor $Yellow
    foreach ($F in $SkippedFiles) {
        Write-Host "    - $F" -ForegroundColor $Yellow
    }
}

Write-Host ""
Write-Host "  LUU Y QUAN TRONG:" -ForegroundColor $Yellow
Write-Host "    - OUTLOOK_MCP_AUDIT_KEY (khoa HMAC) KHONG duoc backup trong file nay." -ForegroundColor $Yellow
Write-Host "    - Luu tru khoa HMAC o noi an toan rieng biet (KeePass, Azure Key Vault, v.v.)." -ForegroundColor $Yellow
Write-Host "    - De restore: giai nen ZIP va copy config.toml ve thu muc project." -ForegroundColor $Yellow
Write-Host ""
Write-Host "  Hoan tat!" -ForegroundColor $Green
Write-Host ""
