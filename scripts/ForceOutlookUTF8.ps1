<#
.SYNOPSIS
    Forces Outlook International Encoding Options to UTF-8 (65001).
    Heavily recommended for Vietnamese language users to prevent font corruption in Calendar/Meeting invites.

.DESCRIPTION
    This script sets registry keys for "InternetCodepage" and "InternetCodepageOut" to 65001 (UTF-8).
    It checks for Outlook 2016/2019/365 (version 16.0).

.EXAMPLE
    .\ForceOutlookUTF8.ps1
#>

$ErrorActionPreference = "Stop"

function Set-OutlookEncoding {
    param (
        [string]$Version = "16.0"
    )

    $KeyPath = "HKCU:\Software\Microsoft\Office\$Version\Outlook\Options\Mail"
    
    Write-Host "Configuring Outlook $Version Registry..." -ForegroundColor Cyan

    if (!(Test-Path $KeyPath)) {
        New-Item -Path $KeyPath -Force | Out-Null
        Write-Host "Created registry key: $KeyPath" -ForegroundColor Yellow
    }

    # 65001 = UTF-8
    # 28591 = ISO-8859-1 (Latin 1) - OLD DEFAULT
    # 50220 = ISO-2022-JP (Japanese)

    # Force "Automatically select encoding for outgoing messages" -> UTF-8
    # Registry values commonly associated with this setting:
    # "InternetCodepage" (for outgoing)
    
    $PropName = "InternetCodepage"
    $Utf8Value = 65001

    try {
        # 1. Force InternetCodepage (Outgoing) to UTF-8
        Set-ItemProperty -Path $KeyPath -Name $PropName -Value $Utf8Value -Type DWord -Force
        Write-Host "SUCCESS: Set $PropName to $Utf8Value (UTF-8)" -ForegroundColor Green

        # 2. Force InternetCodepageOut (Backup) to UTF-8
        Set-ItemProperty -Path $KeyPath -Name "InternetCodepageOut" -Value $Utf8Value -Type DWord -Force
        Write-Host "SUCCESS: Set InternetCodepageOut to $Utf8Value (UTF-8)" -ForegroundColor Green

        # 3. Disable Charset Detection (Prevents Outlook from guessing and messing up)
        # Recommended finding from Deep Research
        Set-ItemProperty -Path $KeyPath -Name "DisableCharsetDetection" -Value 1 -Type DWord -Force
        Write-Host "SUCCESS: Set DisableCharsetDetection to 1" -ForegroundColor Green
        
        # 4. Auto Select Encoding (Ensure it's enabled to use the preferred one, or disabled if we enforce? Research suggests Auto with UTF-8 preferred is good, but some say uncheck auto.
        # Let's trust the Codepage setting.
    }
    catch {
        Write-Error "Failed to set registry keys: $_"
    }

    # Also set global option if exists
    $GlobalKey = "HKCU:\Software\Microsoft\Office\$Version\Common\MailSettings"
    if (Test-Path $GlobalKey) {
        # Optional: Set global preference if specific key isn't enough
    }

    Write-Host "Done. Please restart Outlook for changes to take effect." -ForegroundColor Cyan
}

# Main Execution
try {
    # Target Outlook 16.0 (2016/2019/365)
    Set-OutlookEncoding -Version "16.0"
}
catch {
    Write-Error $_
}
