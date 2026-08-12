<#
.SYNOPSIS
    Install OutlookOkan VSTO Add-in into Outlook.

.DESCRIPTION
    Builds the solution (if needed) and installs the add-in using VSTOInstaller.
    Outlook must be closed before installation.

.PARAMETER SkipBuild
    Skip building and install from existing Release output.

.EXAMPLE
    .\install.ps1              # Build + Install
    .\install.ps1 -SkipBuild   # Install only (use existing build)
#>

param(
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "        OutlookOkan Add-in Installer                    " -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

# --- Check if Outlook is running ---
$outlookProcess = Get-Process "OUTLOOK" -ErrorAction SilentlyContinue
if ($outlookProcess) {
    Write-Host "[!] Outlook is running. Please close Outlook first." -ForegroundColor Red
    Write-Host ""
    $answer = Read-Host "Close Outlook automatically? (y/n)"
    if ($answer -eq 'y') {
        Write-Host "  Closing Outlook..." -ForegroundColor Yellow
        $outlookProcess | Stop-Process -Force
        Start-Sleep -Seconds 3
        Write-Host "  OK: Outlook closed." -ForegroundColor Green
    } else {
        Write-Host "  Aborted. Please close Outlook and try again." -ForegroundColor Yellow
        exit 1
    }
}

# --- Build if needed ---
$vstoFile = Join-Path $PSScriptRoot "OutlookOkan\bin\Release\OutlookOkan.vsto"

if (-not $SkipBuild) {
    Write-Host "[1/3] Building Release..." -ForegroundColor Yellow
    $buildScript = Join-Path $PSScriptRoot "build.ps1"
    & powershell -ExecutionPolicy Bypass -File $buildScript
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  FAIL: Build failed. Cannot install." -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "[1/3] Build skipped." -ForegroundColor Gray
}

if (-not (Test-Path $vstoFile)) {
    Write-Host "  FAIL: $vstoFile not found. Run build first." -ForegroundColor Red
    exit 1
}

# --- Find VSTOInstaller ---
Write-Host "[2/3] Finding VSTOInstaller..." -ForegroundColor Yellow

$vstoPaths = @(
    "C:\Program Files\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe",
    "C:\Program Files (x86)\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe"
)

$vstoInstaller = $null
foreach ($path in $vstoPaths) {
    if (Test-Path $path) {
        $vstoInstaller = $path
        break
    }
}

if ($null -eq $vstoInstaller) {
    Write-Host "  FAIL: VSTOInstaller.exe not found." -ForegroundColor Red
    Write-Host "  Install 'Visual Studio 2010 Tools for Office Runtime':" -ForegroundColor Yellow
    Write-Host "  https://www.microsoft.com/en-us/download/details.aspx?id=56961" -ForegroundColor Gray
    exit 1
}

Write-Host "  OK: $vstoInstaller" -ForegroundColor Green

# --- Install ---
Write-Host "[3/3] Installing OutlookOkan..." -ForegroundColor Yellow
Write-Host "  VSTO: $vstoFile" -ForegroundColor Gray

& $vstoInstaller /install $vstoFile /silent
$installCode = $LASTEXITCODE

if ($installCode -eq 0) {
    # --- [OPTIMIZATION] NGEN: Pre-compile DLL thành native code, loại bỏ JIT overhead ---
    # NGEN (Native Image Generator) biên dịch trước add-in DLL thành native code.
    # Kết quả: Outlook không cần JIT compile khi load add-in → tiết kiệm 200-500ms startup.
    Write-Host ""
    Write-Host "[+] Running NGEN to pre-compile add-in (loại bỏ JIT overhead)..." -ForegroundColor Yellow

    $dllPath = Join-Path $PSScriptRoot "OutlookOkan\bin\Release\OutlookOkan.dll"
    if (Test-Path $dllPath) {
        # Tìm ngen.exe theo .NET Framework version (ưu tiên 4.x)
        $ngenPaths = @(
            "${env:SystemRoot}\Microsoft.NET\Framework\v4.0.30319\ngen.exe",
            "${env:SystemRoot}\Microsoft.NET\Framework64\v4.0.30319\ngen.exe"
        )
        $ngenExe = $ngenPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

        if ($ngenExe) {
            & $ngenExe install $dllPath /nologo 2>&1 | Out-Null
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  OK: NGEN completed — startup sẽ nhanh hơn đáng kể." -ForegroundColor Green
            } else {
                Write-Host "  WARN: NGEN failed (non-critical, add-in vẫn hoạt động)." -ForegroundColor Yellow
            }
        } else {
            Write-Host "  SKIP: ngen.exe không tìm thấy (non-critical)." -ForegroundColor Gray
        }
    } else {
        Write-Host "  SKIP: $dllPath không tồn tại." -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Green
    Write-Host "        INSTALL SUCCESSFUL                              " -ForegroundColor Green
    Write-Host "========================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Open Outlook to use OutlookOkan." -ForegroundColor Cyan
    Write-Host "  Check: File > Options > Add-ins" -ForegroundColor Gray
    Write-Host ""
    exit 0
} else {
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host "        INSTALL FAILED (Exit: $installCode)             " -ForegroundColor Red
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Try: Double-click OutlookOkan.vsto manually:" -ForegroundColor Yellow
    Write-Host "  $vstoFile" -ForegroundColor Gray
    Write-Host ""
    exit $installCode
}
