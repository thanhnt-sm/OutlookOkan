#Requires -Version 5.1
<#
.SYNOPSIS
    Kiểm tra tính toàn vẹn HMAC của từng entry trong audit log.

.DESCRIPTION
    Script đọc từng dòng trong audit.jsonl, tính lại HMAC-SHA256 theo cùng
    thuật toán của AuditLogger._sign_entry() trong Python, so sánh với
    giá trị "hmac" được lưu trong entry.

    Thuật toán HMAC (khớp với audit.py):
      - Chỉ ký 4 trường cốt lõi theo thứ tự cố định: ts, event, session_id, tool
      - Dùng HMAC-SHA256 với key từ env var OUTLOOK_MCP_AUDIT_KEY
      - Lấy 16 ký tự đầu của digest hex (64-bit)

    Kết quả:
      - Entries hợp lệ   : hiển thị xanh (.)
      - Entries bị tamper : hiển thị đỏ với chi tiết
      - Entries không có HMAC: bỏ qua (entries cũ trước khi tính năng HMAC ra đời)

.PARAMETER LogPath
    Đường dẫn đến file audit.jsonl (mặc định: tự tìm từ vị trí script)

.PARAMETER Verbose
    Hiển thị chi tiết từng entry thay vì chỉ hiển thị lỗi

.PARAMETER SampleSize
    Chỉ kiểm tra N entries ngẫu nhiên (0 = kiểm tra tất cả, mặc định: 0)

.EXAMPLE
    .\verify-log-integrity.ps1
    .\verify-log-integrity.ps1 -Verbose
    .\verify-log-integrity.ps1 -LogPath "D:\logs\audit-20260620.jsonl"
    .\verify-log-integrity.ps1 -SampleSize 100

.NOTES
    Phiên bản : 1.0.0
    Tác giả   : OutlookOkan Team
    Yêu cầu   : Env var OUTLOOK_MCP_AUDIT_KEY phải được đặt
#>

[CmdletBinding()]
param(
    [string]$LogPath    = "",
    [switch]$ShowAll,
    [int]$SampleSize    = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ---- Màu sắc ----
$Green  = "Green"
$Red    = "Red"
$Yellow = "Yellow"
$Cyan   = "Cyan"
$White  = "White"
$Gray   = "Gray"

# ---- Tên env var chứa HMAC key (khớp với audit.py) ----
$EnvVarName = "OUTLOOK_MCP_AUDIT_KEY"

# ---- Đường dẫn mặc định ----
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir
if (-not $LogPath) {
    $LogPath = Join-Path $ProjectDir "logs\audit.jsonl"
}

# ---- Tiêu đề ----
Write-Host ""
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "  OUTLOOK MCP — Verify Log Integrity    " -ForegroundColor $Cyan
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "  File : $LogPath" -ForegroundColor $White
Write-Host "  Thoi diem: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor $White
Write-Host ""

# ------------------------------------------------------------------
# Bước 1: Đọc HMAC key từ env var
# ------------------------------------------------------------------
Write-Host "[Bước 1] Đọc HMAC key từ env var $EnvVarName..." -ForegroundColor $Yellow

# Thử đọc từ Process scope trước (session hiện tại), sau đó User scope
$HmacKeyStr = $env:OUTLOOK_MCP_AUDIT_KEY
if (-not $HmacKeyStr) {
    $HmacKeyStr = [System.Environment]::GetEnvironmentVariable($EnvVarName, "User")
}

if (-not $HmacKeyStr) {
    Write-Host ""
    Write-Host "  CANH BAO: Khong tim thay $EnvVarName trong moi truong." -ForegroundColor $Yellow
    Write-Host "  Neu audit log duoc tao khi KHONG co env var, AuditLogger dung session_id lam key." -ForegroundColor $Yellow
    Write-Host "  Trong truong hop do, khong the verify HMAC tu ben ngoai vi khong biet session_id." -ForegroundColor $Yellow
    Write-Host ""
    Write-Host "  Chay setup-env.ps1 truoc de thiet lap key, hoac specify key thu cong:" -ForegroundColor $Yellow
    Write-Host "    `$env:OUTLOOK_MCP_AUDIT_KEY = 'your-key'; .\verify-log-integrity.ps1" -ForegroundColor $White
    Write-Host ""

    $Continue = Read-Host "  Tiep tuc ma khong co key? (chi kiem tra cau truc JSON) [y/N]"
    if ($Continue -ne "y" -and $Continue -ne "Y") {
        exit 1
    }
    $HmacKeyStr = ""
}

if ($HmacKeyStr) {
    Write-Host "  OK — Doc duoc key (dau 8 ky tu: $($HmacKeyStr.Substring(0, [Math]::Min(8, $HmacKeyStr.Length)))...)" -ForegroundColor $Green
} else {
    Write-Host "  Tiep tuc che do kiem tra cau truc JSON only (khong verify HMAC)." -ForegroundColor $Yellow
}

Write-Host ""

# ------------------------------------------------------------------
# Bước 2: Tải hàm tính HMAC-SHA256 trong PowerShell
# Tái tạo y hệt logic _sign_entry() trong Python audit.py:
#   fields_to_sign = { "ts": ..., "event": ..., "session_id": ..., "tool": ... }
#   canonical = json.dumps(fields_to_sign, sort_keys=True, separators=(",", ":"))
#   hmac = HMAC-SHA256(key, canonical)[:16]
# ------------------------------------------------------------------

function Compute-AuditHmac {
    param(
        [string]$Key,
        [string]$Ts,
        [string]$Event,
        [string]$SessionId,
        [string]$Tool
    )

    # Tạo canonical JSON theo đúng format Python json.dumps(sort_keys=True, separators=(",",":"))
    # Thứ tự key theo alphabet: event, session_id, tool, ts
    # (Python sort_keys=True sắp xếp theo thứ tự byte/unicode của tên key)
    $CanonicalJson = '{"event":' + (ConvertTo-JsonString $Event) + ',' +
                     '"session_id":' + (ConvertTo-JsonString $SessionId) + ',' +
                     '"tool":' + (ConvertTo-JsonString $Tool) + ',' +
                     '"ts":' + (ConvertTo-JsonString $Ts) + '}'

    # Tính HMAC-SHA256
    $KeyBytes  = [System.Text.Encoding]::UTF8.GetBytes($Key)
    $DataBytes = [System.Text.Encoding]::UTF8.GetBytes($CanonicalJson)

    $Hmac = [System.Security.Cryptography.HMACSHA256]::new($KeyBytes)
    $Hash = $Hmac.ComputeHash($DataBytes)
    $Hmac.Dispose()

    # Lấy 16 ký tự đầu của hex digest (khớp với audit.py: [:16])
    $HexFull = ($Hash | ForEach-Object { $_.ToString("x2") }) -join ""
    return $HexFull.Substring(0, 16)
}

function ConvertTo-JsonString {
    param([string]$Value)
    # Serialize string theo JSON — phải escape " và \ và control chars
    if ($null -eq $Value) { return 'null' }
    $Escaped = $Value.Replace('\', '\\').Replace('"', '\"').Replace("`n", '\n').Replace("`r", '\r').Replace("`t", '\t')
    return '"' + $Escaped + '"'
}

# ------------------------------------------------------------------
# Bước 3: Kiểm tra file tồn tại
# ------------------------------------------------------------------
Write-Host "[Bước 3] Kiểm tra file log..." -ForegroundColor $Yellow

if (-not (Test-Path $LogPath)) {
    Write-Host "  LOI: Khong tim thay file: $LogPath" -ForegroundColor $Red
    exit 1
}

$FileInfo = Get-Item $LogPath
Write-Host "  File: $($FileInfo.FullName)" -ForegroundColor $White
Write-Host "  Kich thuoc: $([Math]::Round($FileInfo.Length / 1KB, 2)) KB" -ForegroundColor $White
Write-Host "  Cap nhat lan cuoi: $($FileInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor $White
Write-Host ""

# ------------------------------------------------------------------
# Bước 4: Đọc và verify từng entry
# ------------------------------------------------------------------
Write-Host "[Bước 4] Đang kiểm tra từng entry..." -ForegroundColor $Yellow
Write-Host ""

$AllLines = Get-Content $LogPath -Encoding UTF8

# Nếu SampleSize > 0, lấy ngẫu nhiên N dòng
if ($SampleSize -gt 0 -and $SampleSize -lt $AllLines.Count) {
    $Indices  = 0..($AllLines.Count - 1) | Get-Random -Count $SampleSize
    $AllLines = $AllLines[$Indices]
    Write-Host "  Che do kiem tra ngau nhien: $SampleSize / $($AllLines.Count) entries" -ForegroundColor $Yellow
    Write-Host ""
}

# Bộ đếm kết quả
$CountValid      = 0   # Entries có HMAC và PASS
$CountTampered   = 0   # Entries có HMAC nhưng FAIL
$CountNoHmac     = 0   # Entries không có trường hmac (entries cũ)
$CountStructErr  = 0   # Dòng không parse được thành JSON
$CountTotal      = 0

$TamperedEntries = [System.Collections.Generic.List[PSCustomObject]]::new()

$LineNum = 0
foreach ($Line in $AllLines) {
    $LineNum++
    $Line = $Line.Trim()
    if (-not $Line) { continue }

    $CountTotal++

    # Parse JSON
    $Entry = $null
    try {
        $Entry = $Line | ConvertFrom-Json
    } catch {
        $CountStructErr++
        Write-Host "  [STRUCT ERR] Dong $LineNum : Khong parse duoc JSON" -ForegroundColor $Red
        if ($ShowAll) {
            Write-Host "    $($Line.Substring(0, [Math]::Min(80, $Line.Length)))..." -ForegroundColor $Gray
        }
        continue
    }

    # Nếu không có HMAC key → chỉ kiểm tra cấu trúc
    if (-not $HmacKeyStr) {
        $CountNoHmac++
        if ($ShowAll) {
            Write-Host "  [JSON OK]  Dong $LineNum : Cau truc hop le (khong verify HMAC)" -ForegroundColor $Gray
        }
        continue
    }

    # Kiểm tra trường hmac có tồn tại không
    if (-not $Entry.hmac) {
        $CountNoHmac++
        if ($ShowAll) {
            $EventLabel = if ($Entry.event) { $Entry.event } elseif ($Entry.action) { $Entry.action } else { "?" }
            Write-Host "  [NO HMAC]  Dong $LineNum : event='$EventLabel' ts='$($Entry.ts)'" -ForegroundColor $Gray
        }
        continue
    }

    # Lấy 4 trường để tính HMAC (khớp với _sign_entry() trong audit.py)
    $Ts        = if ($Entry.ts)         { [string]$Entry.ts }         else { "" }
    $Event     = if ($Entry.event)      { [string]$Entry.event }      else { "" }
    $SessionId = if ($Entry.session_id) { [string]$Entry.session_id } else { "" }
    $Tool      = if ($Entry.tool)       { [string]$Entry.tool }       else { "" }
    $StoredHmac = [string]$Entry.hmac

    # Tính lại HMAC
    $ComputedHmac = Compute-AuditHmac -Key $HmacKeyStr -Ts $Ts -Event $Event -SessionId $SessionId -Tool $Tool

    # So sánh
    if ($ComputedHmac -eq $StoredHmac) {
        $CountValid++
        if ($ShowAll) {
            Write-Host "  [VALID]    Dong $LineNum : $Ts | $Event | $Tool" -ForegroundColor $Green
        } else {
            # Hiển thị dấu chấm để báo tiến trình khi kiểm tra số lượng lớn
            if ($CountTotal % 50 -eq 0) {
                Write-Host "." -ForegroundColor $Green -NoNewline
            }
        }
    } else {
        $CountTampered++
        $TamperedEntries.Add([PSCustomObject]@{
            LineNum      = $LineNum
            Ts           = $Ts
            Event        = $Event
            SessionId    = $SessionId
            Tool         = $Tool
            StoredHmac   = $StoredHmac
            ComputedHmac = $ComputedHmac
        })
        Write-Host ""
        Write-Host "  [TAMPERED] Dong $LineNum :" -ForegroundColor $Red
        Write-Host "    Thoi gian  : $Ts" -ForegroundColor $Red
        Write-Host "    Event      : $Event" -ForegroundColor $Red
        Write-Host "    Tool       : $Tool" -ForegroundColor $Red
        Write-Host "    HMAC stored  : $StoredHmac" -ForegroundColor $Red
        Write-Host "    HMAC computed: $ComputedHmac" -ForegroundColor $Yellow
        Write-Host "    => Entry nay bi chinh sua sau khi ghi!" -ForegroundColor $Red
        Write-Host ""
    }
}

if (-not $ShowAll -and $HmacKeyStr) { Write-Host "" }

# ------------------------------------------------------------------
# Bước 5: Báo cáo kết quả
# ------------------------------------------------------------------
Write-Host ""
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "  BAO CAO KET QUA VERIFY                " -ForegroundColor $Cyan
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host ""
Write-Host "  Tong entries kiem tra : $CountTotal" -ForegroundColor $White
Write-Host "  HMAC VALID            : $CountValid" -ForegroundColor $Green
Write-Host "  HMAC TAMPERED         : $CountTampered" -ForegroundColor $(if ($CountTampered -gt 0) { $Red } else { $Green })
Write-Host "  Khong co truong HMAC  : $CountNoHmac (entries cu hoac session_id-key)" -ForegroundColor $Yellow
Write-Host "  Loi parse JSON        : $CountStructErr" -ForegroundColor $(if ($CountStructErr -gt 0) { $Red } else { $Green })
Write-Host ""

if ($CountTampered -gt 0) {
    Write-Host "  [!!] CANH BAO: Phat hien $CountTampered entries bi chay sua!" -ForegroundColor $Red
    Write-Host ""
    Write-Host "  Danh sach entries bi tamper:" -ForegroundColor $Red
    foreach ($T in $TamperedEntries) {
        Write-Host "    - Dong $($T.LineNum): $($T.Ts) | $($T.Event) | $($T.Tool)" -ForegroundColor $Red
    }
    Write-Host ""
    Write-Host "  Hanh dong khuyen nghi:" -ForegroundColor $Yellow
    Write-Host "    1. Khong tin vao noi dung audit log nay de dieu tra bao mat." -ForegroundColor $Yellow
    Write-Host "    2. Sao chep file log bi hong vao noi luu tru rieng biet." -ForegroundColor $Yellow
    Write-Host "    3. Bao cao bao mat cho team." -ForegroundColor $Yellow
} elseif ($CountValid -gt 0) {
    Write-Host "  [OK] Toan bo $CountValid entries co HMAC deu hop le." -ForegroundColor $Green
    Write-Host "  Audit log khong bi chinh sua." -ForegroundColor $Green
} else {
    Write-Host "  Khong co entry nao de verify HMAC." -ForegroundColor $Yellow
    Write-Host "  (Co the log duoc tao truoc khi thiet lap OUTLOOK_MCP_AUDIT_KEY)" -ForegroundColor $Yellow
}

Write-Host ""
