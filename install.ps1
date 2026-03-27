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
