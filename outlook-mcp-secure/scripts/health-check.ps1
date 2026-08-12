#Requires -Version 5.1
<#
.SYNOPSIS
    Kiểm tra toàn bộ sức khỏe của hệ thống Outlook MCP Secure.

.DESCRIPTION
    Script chạy 5 bài kiểm tra theo thứ tự:
      1. Outlook Desktop — kiểm tra tiến trình OUTLOOK.EXE đang chạy
      2. Claude MCP — kiểm tra đăng ký MCP trong cấu hình Claude Desktop
      3. Audit log — kiểm tra file log được ghi trong vòng 1 giờ qua
      4. Config syntax — kiểm tra cú pháp config.toml bằng Python ast
      5. Venv packages — kiểm tra các gói Python cần thiết đã cài đủ chưa

    Kết quả mỗi mục hiển thị PASS (xanh) hoặc FAIL (đỏ) kèm chi tiết.
    Cuối cùng tổng kết bao nhiêu PASS / bao nhiêu FAIL.

.NOTES
    Phiên bản : 1.0.0
    Tác giả   : OutlookOkan Team
    Yêu cầu   : PowerShell 5.1+, Python trong PATH hoặc venv
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"  # Tiếp tục dù có lỗi để chạy đủ 5 bài kiểm tra

# ---- Màu sắc hiển thị ----
$Green  = "Green"
$Yellow = "Yellow"
$Red    = "Red"
$Cyan   = "Cyan"
$White  = "White"

# ---- Đường dẫn project (tính tương đối từ vị trí script) ----
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir   # Thư mục gốc outlook-mcp-secure

# Đường dẫn các file cần kiểm tra
$ConfigPath   = Join-Path $ProjectDir "config.toml"
$LogDir       = Join-Path $ProjectDir "logs"
$AuditLogPath = Join-Path $LogDir "audit.jsonl"
$VenvDir      = Join-Path $ProjectDir ".venv"
$VenvPython   = Join-Path $VenvDir "Scripts\python.exe"

# Danh sách gói Python bắt buộc phải cài
$RequiredPackages = @("mcp", "pywin32", "toml")

# Đếm kết quả
$TotalPass = 0
$TotalFail = 0

# ---- Hàm hiển thị kết quả ----
function Write-Result {
    param(
        [string]$CheckName,
        [bool]$Passed,
        [string]$Detail = ""
    )
    if ($Passed) {
        Write-Host "  [PASS] " -ForegroundColor $Green -NoNewline
        $script:TotalPass++
    } else {
        Write-Host "  [FAIL] " -ForegroundColor $Red -NoNewline
        $script:TotalFail++
    }
    Write-Host "$CheckName" -ForegroundColor $White -NoNewline
    if ($Detail) {
        Write-Host " — $Detail" -ForegroundColor $Yellow
    } else {
        Write-Host ""
    }
}

# ---- Tiêu đề ----
Write-Host ""
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "  OUTLOOK MCP SECURE — Health Check     " -ForegroundColor $Cyan
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "  Project: $ProjectDir" -ForegroundColor $White
Write-Host "  Thời gian: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor $White
Write-Host ""

# ==================================================================
# Bài kiểm tra 1: Outlook Desktop process
# Kiểm tra tiến trình OUTLOOK.EXE có đang chạy không
# Claude MCP dùng win32com để kết nối Outlook — cần Outlook mở sẵn
# ==================================================================
Write-Host "[1/5] Outlook Desktop Process Check" -ForegroundColor $Cyan

try {
    # Get-Process trả về lỗi nếu không tìm thấy process → bắt bằng ErrorAction
    $OutlookProc = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue

    if ($OutlookProc) {
        $ProcCount = ($OutlookProc | Measure-Object).Count
        $MemMB     = [Math]::Round($OutlookProc[0].WorkingSet64 / 1MB, 1)
        Write-Result -CheckName "OUTLOOK.EXE" -Passed $true `
                     -Detail "$ProcCount tien trinh, bo nho: ${MemMB} MB"
    } else {
        Write-Result -CheckName "OUTLOOK.EXE" -Passed $false `
                     -Detail "Tien trinh khong chay — hay mo Outlook Desktop truoc"
    }
} catch {
    Write-Result -CheckName "OUTLOOK.EXE" -Passed $false -Detail "Loi kiem tra: $_"
}

Write-Host ""

# ==================================================================
# Bài kiểm tra 2: Claude MCP Registration
# Kiểm tra outlook-mcp-secure đã được đăng ký trong Claude Desktop config
# File cấu hình Claude Desktop: %APPDATA%\Claude\claude_desktop_config.json
# ==================================================================
Write-Host "[2/5] Claude MCP Registration Check" -ForegroundColor $Cyan

try {
    $ClaudeConfigPath = Join-Path $env:APPDATA "Claude\claude_desktop_config.json"

    if (-not (Test-Path $ClaudeConfigPath)) {
        Write-Result -CheckName "Claude Desktop config" -Passed $false `
                     -Detail "Khong tim thay file: $ClaudeConfigPath"
    } else {
        $ClaudeConfig = Get-Content $ClaudeConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

        # Tìm MCP server có liên quan đến outlook hoặc outlook-mcp
        $McpServers = $ClaudeConfig.mcpServers
        if ($null -eq $McpServers) {
            Write-Result -CheckName "MCP Servers" -Passed $false `
                         -Detail "Khong co mcpServers trong config Claude"
        } else {
            # Lấy danh sách tên server (properties của object JSON)
            $ServerNames = $McpServers.PSObject.Properties.Name
            $OutlookMcp  = $ServerNames | Where-Object { $_ -match "outlook" }

            if ($OutlookMcp) {
                Write-Result -CheckName "MCP Registration" -Passed $true `
                             -Detail "Tim thay server: $($OutlookMcp -join ', ')"

                # Kiểm tra thêm: command trỏ đúng đến project này không
                foreach ($SrvName in $OutlookMcp) {
                    $SrvConfig = $McpServers.$SrvName
                    $SrvCmd    = $SrvConfig.command
                    $SrvArgs   = $SrvConfig.args -join " "
                    Write-Host "    Server '$SrvName': command='$SrvCmd' args='$SrvArgs'" -ForegroundColor $White
                }
            } else {
                Write-Result -CheckName "MCP Registration" -Passed $false `
                             -Detail "Khong co server nao ten 'outlook' — cac server hien co: $($ServerNames -join ', ')"
            }
        }
    }
} catch {
    Write-Result -CheckName "Claude MCP Registration" -Passed $false -Detail "Loi kiem tra: $_"
}

Write-Host ""

# ==================================================================
# Bài kiểm tra 3: Audit Log File Activity
# Kiểm tra file audit.jsonl có tồn tại và được cập nhật trong 1 giờ qua không
# Nếu server đang chạy, log phải có hoạt động gần đây
# ==================================================================
Write-Host "[3/5] Audit Log Activity Check (trong 1 gio)" -ForegroundColor $Cyan

try {
    if (-not (Test-Path $AuditLogPath)) {
        Write-Result -CheckName "Audit Log File" -Passed $false `
                     -Detail "File khong ton tai: $AuditLogPath"
    } else {
        $LogFile      = Get-Item $AuditLogPath
        $FileSizeKB   = [Math]::Round($LogFile.Length / 1KB, 2)
        $LastModified = $LogFile.LastWriteTime
        $AgeMins      = [Math]::Round(((Get-Date) - $LastModified).TotalMinutes, 1)

        if ($AgeMins -le 60) {
            Write-Result -CheckName "Audit Log Activity" -Passed $true `
                         -Detail "Cap nhat ${AgeMins} phut truoc, kich thuoc: ${FileSizeKB} KB"
        } else {
            # File tồn tại nhưng cũ hơn 1 giờ — server có thể không đang chạy
            Write-Result -CheckName "Audit Log Activity" -Passed $false `
                         -Detail "Log cu ${AgeMins} phut (>60 phut) — server co the da tat, kich thuoc: ${FileSizeKB} KB"
        }

        # Đếm số entries trong log
        $LineCount = (Get-Content $AuditLogPath -Encoding UTF8 | Measure-Object -Line).Lines
        Write-Host "    Tong so entries: $LineCount dong" -ForegroundColor $White
    }
} catch {
    Write-Result -CheckName "Audit Log Activity" -Passed $false -Detail "Loi kiem tra: $_"
}

Write-Host ""

# ==================================================================
# Bài kiểm tra 4: Config Syntax Check
# Dùng Python để parse config.toml và kiểm tra cú pháp hợp lệ
# Nếu có venv, dùng python từ venv; ngược lại dùng python hệ thống
# ==================================================================
Write-Host "[4/5] Config Syntax Check (Python TOML parse)" -ForegroundColor $Cyan

try {
    if (-not (Test-Path $ConfigPath)) {
        Write-Result -CheckName "Config File" -Passed $false `
                     -Detail "Khong tim thay: $ConfigPath"
    } else {
        # Xác định python nào sẽ dùng: venv trước, hệ thống sau
        $PythonExe = $null
        if (Test-Path $VenvPython) {
            $PythonExe = $VenvPython
            Write-Host "    Dung Python tu venv: $VenvPython" -ForegroundColor $White
        } else {
            # Thử tìm python trong PATH
            $PythonExe = (Get-Command python -ErrorAction SilentlyContinue)?.Source
            if ($PythonExe) {
                Write-Host "    Dung Python he thong: $PythonExe" -ForegroundColor $White
            }
        }

        if (-not $PythonExe) {
            Write-Result -CheckName "Config Syntax" -Passed $false `
                         -Detail "Khong tim thay Python — cai Python hoac kich hoat venv"
        } else {
            # Script Python nhỏ để parse TOML và in ra các section chính
            $PythonScript = @"
import sys
try:
    import tomllib
except ImportError:
    try:
        import tomli as tomllib
    except ImportError:
        try:
            import toml
            # toml dung API khac
            with open(sys.argv[1], 'r', encoding='utf-8') as f:
                data = toml.load(f)
            sections = list(data.keys())
            print('OK:' + ','.join(sections))
            sys.exit(0)
        except ImportError:
            print('ERROR:Khong co thu vien TOML (tomllib/tomli/toml)')
            sys.exit(1)
with open(sys.argv[1], 'rb') as f:
    data = tomllib.load(f)
sections = list(data.keys())
print('OK:' + ','.join(sections))
"@
            # Lưu script tạm thời
            $TempPy = [System.IO.Path]::GetTempFileName() + ".py"
            $PythonScript | Out-File -FilePath $TempPy -Encoding UTF8

            try {
                $PyResult = & $PythonExe $TempPy $ConfigPath 2>&1
                $ExitCode = $LASTEXITCODE

                if ($ExitCode -eq 0 -and $PyResult -match "^OK:") {
                    $Sections = ($PyResult -replace "^OK:", "").Trim()
                    Write-Result -CheckName "Config Syntax (TOML)" -Passed $true `
                                 -Detail "Hop le — cac section: $Sections"
                } else {
                    Write-Result -CheckName "Config Syntax (TOML)" -Passed $false `
                                 -Detail "LOI cu phap: $PyResult"
                }
            } finally {
                Remove-Item $TempPy -ErrorAction SilentlyContinue
            }
        }
    }
} catch {
    Write-Result -CheckName "Config Syntax" -Passed $false -Detail "Loi kiem tra: $_"
}

Write-Host ""

# ==================================================================
# Bài kiểm tra 5: Venv Packages Check
# Kiểm tra các gói Python cần thiết đã cài đầy đủ trong venv
# ==================================================================
Write-Host "[5/5] Venv Packages Check" -ForegroundColor $Cyan

try {
    if (-not (Test-Path $VenvDir)) {
        Write-Result -CheckName "Virtual Environment" -Passed $false `
                     -Detail "Khong tim thay venv tai: $VenvDir"
        Write-Host "    Chay lenh: python -m venv .venv" -ForegroundColor $Yellow
    } elseif (-not (Test-Path $VenvPython)) {
        Write-Result -CheckName "Venv Python" -Passed $false `
                     -Detail "Khong tim thay python.exe trong venv: $VenvPython"
    } else {
        # Lấy danh sách gói đã cài trong venv
        $InstalledRaw = & $VenvPython -m pip list --format=columns 2>&1
        $InstalledPkgs = $InstalledRaw | Select-Object -Skip 2 |
                         ForEach-Object { ($_ -split "\s+")[0].ToLower() }

        $MissingPkgs = @()
        $FoundPkgs   = @()

        foreach ($Pkg in $RequiredPackages) {
            $PkgLower = $Pkg.ToLower()
            # pywin32 cài xong tên gói là "pywin32" nhưng module là win32com
            if ($PkgLower -eq "pywin32") {
                # Kiểm tra module win32com thay vì tên gói
                $CheckResult = & $VenvPython -c "import win32com; print('ok')" 2>&1
                if ($CheckResult -eq "ok") {
                    $FoundPkgs += "pywin32"
                } else {
                    $MissingPkgs += "pywin32"
                }
            } elseif ($InstalledPkgs -contains $PkgLower) {
                $FoundPkgs += $Pkg
            } else {
                $MissingPkgs += $Pkg
            }
        }

        if ($MissingPkgs.Count -eq 0) {
            Write-Result -CheckName "Venv Packages" -Passed $true `
                         -Detail "Du goi: $($FoundPkgs -join ', ')"
        } else {
            Write-Result -CheckName "Venv Packages" -Passed $false `
                         -Detail "Thieu goi: $($MissingPkgs -join ', ')"
            Write-Host "    Chay lenh: .venv\Scripts\pip install $($MissingPkgs -join ' ')" -ForegroundColor $Yellow
        }

        # Hiển thị phiên bản Python
        $PyVer = & $VenvPython --version 2>&1
        Write-Host "    Python: $PyVer" -ForegroundColor $White
    }
} catch {
    Write-Result -CheckName "Venv Packages" -Passed $false -Detail "Loi kiem tra: $_"
}

Write-Host ""

# ==================================================================
# Tổng kết kết quả
# ==================================================================
$TotalChecks = $TotalPass + $TotalFail

Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "  TON KET KIEM TRA" -ForegroundColor $Cyan
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host ""

Write-Host "  Tong so bai kiem tra : $TotalChecks" -ForegroundColor $White
Write-Host "  PASS                 : $TotalPass" -ForegroundColor $Green
Write-Host "  FAIL                 : $TotalFail" -ForegroundColor $(if ($TotalFail -gt 0) { $Red } else { $Green })
Write-Host ""

if ($TotalFail -eq 0) {
    Write-Host "  [OK] He thong hoat dong binh thuong!" -ForegroundColor $Green
} else {
    Write-Host "  [!!] Co $TotalFail van de can xu ly — xem chi tiet o tren." -ForegroundColor $Red
}

Write-Host ""
