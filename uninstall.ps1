<#
.SYNOPSIS
    Uninstall OutlookOkan VSTO Add-in from Outlook.

.DESCRIPTION
    Removes the OutlookOkan add-in using VSTOInstaller.
    Also cleans up registry entries if VSTOInstaller fails.

.EXAMPLE
    .\uninstall.ps1
#>

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "        OutlookOkan Add-in Uninstaller                  " -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

# --- Check if Outlook is running ---
$outlookProcess = Get-Process "OUTLOOK" -ErrorAction SilentlyContinue
if ($outlookProcess) {
    Write-Host "[!] Outlook is running. Please close Outlook first." -ForegroundColor Red
    $answer = Read-Host "Close Outlook automatically? (y/n)"
    if ($answer -eq 'y') {
        $outlookProcess | Stop-Process -Force
        Start-Sleep -Seconds 3
        Write-Host "  OK: Outlook closed." -ForegroundColor Green
    } else {
        Write-Host "  Aborted." -ForegroundColor Yellow
        exit 1
    }
}

# --- Find VSTOInstaller ---
Write-Host "[1/2] Finding VSTOInstaller..." -ForegroundColor Yellow

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

# --- Uninstall ---
Write-Host "[2/2] Uninstalling OutlookOkan..." -ForegroundColor Yellow

$vstoFile = Join-Path $PSScriptRoot "OutlookOkan\bin\Release\OutlookOkan.vsto"
$uninstalled = $false

if ($vstoInstaller) {
    Write-Host "  Using: $vstoInstaller" -ForegroundColor Gray
    & $vstoInstaller /uninstall $vstoFile /silent 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $uninstalled = $true
    }
}

# --- Fallback: Clean registry ---
if (-not $uninstalled) {
    Write-Host "  VSTOInstaller failed or not found. Cleaning registry..." -ForegroundColor Yellow
}

$regPaths = @(
    "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookOkan",
    "HKLM:\Software\Microsoft\Office\Outlook\Addins\OutlookOkan",
    "HKLM:\Software\WOW6432Node\Microsoft\Office\Outlook\Addins\OutlookOkan"
)

foreach ($reg in $regPaths) {
    if (Test-Path $reg) {
        Remove-Item $reg -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed: $reg" -ForegroundColor Gray
        $uninstalled = $true
    }
}

if ($uninstalled) {
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Green
    Write-Host "        UNINSTALL SUCCESSFUL                            " -ForegroundColor Green
    Write-Host "========================================================" -ForegroundColor Green
    Write-Host ""
    exit 0
} else {
    Write-Host ""
    Write-Host "  OutlookOkan was not installed or already removed." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}
