#Requires -Version 5.1
<#
.SYNOPSIS
    Đọc và hiển thị audit log của Outlook MCP Secure theo nhiều bộ lọc.

.DESCRIPTION
    Script đọc file audit.jsonl (định dạng JSON Lines — mỗi dòng là 1 JSON object)
    và hiển thị đẹp với màu sắc theo loại sự kiện.

    Các chế độ lọc:
      -Range     : today | week | month | all   (mặc định: today)
      -EventType : Lọc theo loại sự kiện cụ thể (ví dụ: security_event, session_start, error)
      -TopN      : Chỉ hiển thị N entries gần nhất (mặc định: 50)
      -CountOnly : Chỉ đếm theo loại sự kiện, không hiển thị chi tiết
      -ShowRateLimit : Chỉ hiển thị các sự kiện rate limit

.PARAMETER Range
    Khoảng thời gian: today, week, month, all (mặc định: today)

.PARAMETER EventType
    Lọc theo loại sự kiện (khớp một phần, không phân biệt hoa thường)

.PARAMETER TopN
    Số entries tối đa hiển thị (mặc định: 50)

.PARAMETER CountOnly
    Chỉ đếm và nhóm theo loại sự kiện

.PARAMETER ShowRateLimit
    Chỉ hiển thị sự kiện rate_limit_exceeded

.EXAMPLE
    .\view-audit-log.ps1
    .\view-audit-log.ps1 -Range week -TopN 100
    .\view-audit-log.ps1 -EventType security_event
    .\view-audit-log.ps1 -CountOnly
    .\view-audit-log.ps1 -ShowRateLimit

.NOTES
    Phiên bản : 1.0.0
    Tác giả   : OutlookOkan Team
    Log format : JSON Lines (.jsonl), encoding UTF-8
#>

[CmdletBinding()]
param(
    [ValidateSet("today", "week", "month", "all")]
    [string]$Range = "today",

    [string]$EventType = "",

    [int]$TopN = 50,

    [switch]$CountOnly,

    [switch]$ShowRateLimit
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ---- Màu sắc hiển thị theo loại sự kiện ----
$ColDefault  = "White"
$ColSecurity = "Red"
$ColSession  = "Cyan"
$ColSuccess  = "Green"
$ColError    = "Red"
$ColBlocked  = "Yellow"
$ColInfo     = "Gray"
$ColTitle    = "Cyan"

# ---- Đường dẫn project ----
$ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir   = Split-Path -Parent $ScriptDir
$AuditLogPath = Join-Path $ProjectDir "logs\audit.jsonl"

# ---- Tiêu đề ----
Write-Host ""
Write-Host "========================================" -ForegroundColor $ColTitle
Write-Host "  OUTLOOK MCP — Audit Log Viewer        " -ForegroundColor $ColTitle
Write-Host "========================================" -ForegroundColor $ColTitle
Write-Host "  File: $AuditLogPath" -ForegroundColor $ColDefault
Write-Host "  Bo loc: range=$Range | event='$EventType' | top=$TopN | countOnly=$CountOnly | rateLimit=$ShowRateLimit" -ForegroundColor $ColDefault
Write-Host ""

# ------------------------------------------------------------------
# Kiểm tra file tồn tại
# ------------------------------------------------------------------
if (-not (Test-Path $AuditLogPath)) {
    Write-Host "  LOI: Khong tim thay file audit log: $AuditLogPath" -ForegroundColor $ColError
    Write-Host "  Dam bao MCP server da chay it nhat mot lan de tao file log." -ForegroundColor $ColDefault
    exit 1
}

# ------------------------------------------------------------------
# Tính ngưỡng thời gian theo -Range
# ------------------------------------------------------------------
$Now       = Get-Date
$StartTime = switch ($Range) {
    "today" { $Now.Date }                          # Từ 00:00:00 hôm nay
    "week"  { $Now.Date.AddDays(-7) }              # 7 ngày trước
    "month" { $Now.Date.AddDays(-30) }             # 30 ngày trước
    "all"   { [DateTime]::MinValue }               # Toàn bộ
}

Write-Host "  Khoang thoi gian: tu $($StartTime.ToString('yyyy-MM-dd HH:mm')) den bay gio" -ForegroundColor $ColDefault
Write-Host ""

# ------------------------------------------------------------------
# Đọc và parse từng dòng JSON
# Mỗi dòng là một JSON object độc lập (JSON Lines format)
# ------------------------------------------------------------------
Write-Host "  Dang doc file log..." -ForegroundColor $ColInfo

$AllLines    = Get-Content $AuditLogPath -Encoding UTF8
$TotalLines  = $AllLines.Count
$ParseErrors = 0
$Entries     = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($Line in $AllLines) {
    $Line = $Line.Trim()
    if (-not $Line) { continue }

    try {
        $Entry = $Line | ConvertFrom-Json
    } catch {
        $ParseErrors++
        continue
    }

    # Lấy timestamp — field tên "ts"
    $EntryTime = $null
    if ($Entry.ts) {
        try {
            $EntryTime = [DateTimeOffset]::Parse($Entry.ts).LocalDateTime
        } catch {
            $EntryTime = $null
        }
    }

    # Lọc theo khoảng thời gian
    if ($EntryTime -and $EntryTime -lt $StartTime) { continue }

    # Lọc theo EventType (nếu có)
    if ($EventType) {
        $EntryEvent  = if ($Entry.event) { $Entry.event } else { $Entry.action }
        $EntryTool   = if ($Entry.tool) { $Entry.tool } else { "" }
        $EntryResult = if ($Entry.result) { $Entry.result } else { "" }

        $MatchEvent  = $EntryEvent -match [regex]::Escape($EventType)
        $MatchTool   = $EntryTool  -match [regex]::Escape($EventType)
        $MatchResult = $EntryResult -match [regex]::Escape($EventType)

        if (-not ($MatchEvent -or $MatchTool -or $MatchResult)) { continue }
    }

    # Lọc theo ShowRateLimit
    if ($ShowRateLimit) {
        $IsRateLimit = $false
        if ($Entry.result -match "rate_limit" -or
            $Entry.event  -match "rate_limit" -or
            ($Entry.params -and ($Entry.params | ConvertTo-Json) -match "rate_limit") -or
            ($Entry.details -and $Entry.details -match "rate_limit")) {
            $IsRateLimit = $true
        }
        if (-not $IsRateLimit) { continue }
    }

    $Entries.Add([PSCustomObject]@{
        Time    = $EntryTime
        Raw     = $Entry
        Line    = $Line
    })
}

Write-Host "  Tong dong trong file : $TotalLines" -ForegroundColor $ColDefault
Write-Host "  Entries hop le       : $($Entries.Count)" -ForegroundColor $ColDefault
if ($ParseErrors -gt 0) {
    Write-Host "  Dong loi parse       : $ParseErrors" -ForegroundColor $ColError
}
Write-Host ""

# ------------------------------------------------------------------
# Chế độ -CountOnly: chỉ thống kê, không hiển thị chi tiết
# ------------------------------------------------------------------
if ($CountOnly) {
    Write-Host "  THONG KE THEO LOAI SU KIEN:" -ForegroundColor $ColTitle
    Write-Host "  ----------------------------------------" -ForegroundColor $ColInfo

    # Nhóm theo event type (ưu tiên field 'event', fallback 'action', fallback 'result')
    $GroupMap = @{}
    foreach ($E in $Entries) {
        $R = $E.Raw
        $Key = if ($R.event -and $R.event -ne "") { $R.event }
               elseif ($R.action -and $R.action -ne "") { "$($R.tool)/$($R.action)" }
               else { "(unknown)" }
        if (-not $GroupMap.ContainsKey($Key)) { $GroupMap[$Key] = 0 }
        $GroupMap[$Key]++
    }

    # Sắp xếp theo số lượng giảm dần
    $GroupMap.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
        $Count  = $_.Value
        $Label  = $_.Key
        $Bar    = "#" * [Math]::Min($Count, 40)

        # Chọn màu theo loại sự kiện
        $Color = switch -Wildcard ($Label) {
            "*security*"      { $ColSecurity }
            "*error*"         { $ColError }
            "*blocked*"       { $ColBlocked }
            "*rate_limit*"    { $ColError }
            "*session*"       { $ColSession }
            "*success*"       { $ColSuccess }
            default           { $ColDefault }
        }

        Write-Host ("  {0,-35} {1,5}  {2}" -f $Label, $Count, $Bar) -ForegroundColor $Color
    }

    Write-Host ""
    Write-Host "  Tong entries hien thi: $($Entries.Count)" -ForegroundColor $ColDefault
    Write-Host ""
    exit 0
}

# ------------------------------------------------------------------
# Hiển thị chi tiết entries — giới hạn theo -TopN
# Lấy TopN entries cuối cùng (gần nhất)
# ------------------------------------------------------------------
$DisplayEntries = $Entries | Select-Object -Last $TopN

Write-Host "  Hien thi $($DisplayEntries.Count) entries (toi da: $TopN)" -ForegroundColor $ColDefault
Write-Host ""
Write-Host ("  {0,-24}  {1,-20}  {2,-10}  {3,-12}  {4}" -f "THOI GIAN", "TOOL", "ACTION", "RESULT", "CHI TIET") -ForegroundColor $ColTitle
Write-Host "  $("-" * 95)" -ForegroundColor $ColInfo

foreach ($E in $DisplayEntries) {
    $R       = $E.Raw
    $TimeStr = if ($E.Time) { $E.Time.ToString("MM-dd HH:mm:ss") } else { "N/A              " }

    # Lấy các trường chính để hiển thị
    $Tool    = if ($R.tool)    { "$($R.tool)" }    else { if ($R.event) { "[event]" } else { "" } }
    $Action  = if ($R.action)  { $R.action }       else { if ($R.event) { $R.event } else { "" } }
    $Result  = if ($R.result)  { $R.result }       else { "" }
    $Detail  = if ($R.details) { $R.details }      else { "" }
    $Items   = if ($R.items_returned -ne $null) { "(n=$($R.items_returned))" } else { "" }

    # Ghép thêm event name nếu có
    if ($R.event -and -not $R.tool) {
        $Tool   = "[SYSTEM]"
        $Action = $R.event
    }

    # Rút gọn nếu quá dài
    if ($Tool.Length   -gt 20) { $Tool   = $Tool.Substring(0,17) + "..." }
    if ($Action.Length -gt 10) { $Action = $Action.Substring(0,7) + "..." }
    if ($Result.Length -gt 12) { $Result = $Result.Substring(0,9) + "..." }
    if ($Detail.Length -gt 40) { $Detail = $Detail.Substring(0,37) + "..." }

    # Chọn màu theo kết quả
    $RowColor = switch -Wildcard ($Result) {
        "ok"             { $ColSuccess }
        "blocked"        { $ColBlocked }
        "error"          { $ColError }
        "SECURITY_EVENT" { $ColSecurity }
        "pending"        { $ColInfo }
        default {
            switch -Wildcard ($Action) {
                "session_start" { $ColSession }
                "session_end"   { $ColSession }
                "server_start"  { $ColSession }
                "server_stop"   { $ColSession }
                default         { $ColDefault }
            }
        }
    }

    $DetailDisplay = if ($Detail) { "$Detail $Items" } else { $Items }
    Write-Host ("  {0,-24}  {1,-20}  {2,-10}  {3,-12}  {4}" -f `
        $TimeStr, $Tool, $Action, $Result, $DetailDisplay) -ForegroundColor $RowColor
}

Write-Host ""

# ------------------------------------------------------------------
# Thống kê nhanh cuối trang
# ------------------------------------------------------------------
Write-Host "  ----------------------------------------" -ForegroundColor $ColInfo
$SecurityCount  = ($Entries | Where-Object { $_.Raw.result -eq "SECURITY_EVENT" }).Count
$ErrorCount     = ($Entries | Where-Object { $_.Raw.result -eq "error" }).Count
$BlockedCount   = ($Entries | Where-Object { $_.Raw.result -eq "blocked" }).Count
$SuccessCount   = ($Entries | Where-Object { $_.Raw.result -eq "ok" }).Count
$RateLimitCount = ($Entries | Where-Object {
    $r = $_.Raw
    ($r.params -and ($r.params | ConvertTo-Json) -match "rate_limit") -or
    ($r.details -and $r.details -match "rate_limit") -or
    ($r.event   -and $r.event   -match "rate_limit")
}).Count

Write-Host "  Tom tat: " -ForegroundColor $ColTitle -NoNewline
Write-Host "OK=$SuccessCount  " -ForegroundColor $ColSuccess -NoNewline
Write-Host "BLOCKED=$BlockedCount  " -ForegroundColor $ColBlocked -NoNewline
Write-Host "ERROR=$ErrorCount  " -ForegroundColor $ColError -NoNewline
Write-Host "SECURITY=$SecurityCount  " -ForegroundColor $ColSecurity -NoNewline
Write-Host "RATE_LIMIT=$RateLimitCount" -ForegroundColor $ColError
Write-Host ""
