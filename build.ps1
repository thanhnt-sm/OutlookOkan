<#
.SYNOPSIS
    Build script for OutlookOkan VSTO Add-in.
    
.DESCRIPTION
    Automatically locates MSBuild and rebuilds the solution.
    Supports Debug and Release configurations.
    
.PARAMETER Configuration
    Build configuration: Debug or Release (default: Release)
    
.PARAMETER Verbose
    Enable verbose output (default: quiet)
    
.EXAMPLE
    .\build.ps1                    # Build Release configuration
    .\build.ps1 -Configuration Debug   # Build Debug configuration
    .\build.ps1 -Verbose          # Build Release with verbose output

.NOTES
    Author: OutlookOkan Build Team
    Version: 1.0
#>

param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    
    [switch]$Verbose
)

# Error handling
$ErrorActionPreference = "Stop"

Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║          OutlookOkan Build Script v1.0            ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Find MSBuild path - check multiple possible locations
$msbuildPaths = @(
    "C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
)

$msbuildPath = $null
Write-Host "🔍 Searching for MSBuild..." -ForegroundColor Yellow

foreach ($path in $msbuildPaths) {
    if (Test-Path -Path $path) {
        $msbuildPath = $path
        Write-Host "✅ Found MSBuild at: $path" -ForegroundColor Green
        break
    }
}

if ($null -eq $msbuildPath) {
    Write-Host ""
    Write-Host "❌ ERROR: MSBuild not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Checked locations:" -ForegroundColor Yellow
    $msbuildPaths | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }
    Write-Host ""
    Write-Host "💡 Solution:" -ForegroundColor Cyan
    Write-Host "  1. Install Visual Studio 2019 or later" -ForegroundColor Gray
    Write-Host "  2. Make sure MSBuild tools are included" -ForegroundColor Gray
    Write-Host ""
    exit 1
}

$solutionFile = Join-Path $PSScriptRoot "OutlookOkan.sln"

if (-not (Test-Path -Path $solutionFile)) {
    Write-Host ""
    Write-Host "❌ ERROR: Solution file not found!" -ForegroundColor Red
    Write-Host "Expected: $solutionFile" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Write-Host ""
Write-Host "📦 Build Configuration:" -ForegroundColor Cyan
Write-Host "  Solution: $solutionFile" -ForegroundColor Gray
Write-Host "  Configuration: $Configuration" -ForegroundColor Gray
Write-Host "  Verbose: $(if ($Verbose) { 'Yes' } else { 'No' })" -ForegroundColor Gray
Write-Host ""

# Set verbosity level
$verbosityLevel = if ($Verbose) { "normal" } else { "quiet" }

Write-Host "🔨 Starting Build..." -ForegroundColor Cyan
Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
Write-Host ""

# Execute MSBuild
try {
    & $msbuildPath $solutionFile `
        /t:Rebuild `
        /p:Configuration=$Configuration `
        /v:$verbosityLevel `
        /nologo
    
    $buildExitCode = $LASTEXITCODE
}
catch {
    Write-Host ""
    Write-Host "❌ ERROR: Failed to execute MSBuild" -ForegroundColor Red
    Write-Host "Message: $_" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Write-Host ""
Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray

# Check build result
if ($buildExitCode -eq 0) {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                                                    ║" -ForegroundColor Green
    Write-Host "║        ✅ BUILD SUCCESSFUL ($Configuration)        ║" -ForegroundColor Green
    Write-Host "║                                                    ║" -ForegroundColor Green
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "📍 Output Location:" -ForegroundColor Cyan
    Write-Host "  - Main DLL: OutlookOkan\bin\$Configuration\OutlookOkan.dll" -ForegroundColor Gray
    Write-Host ""
    exit 0
}
else {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║                                                    ║" -ForegroundColor Red
    Write-Host "║        ❌ BUILD FAILED (Exit Code: $buildExitCode)   ║" -ForegroundColor Red
    Write-Host "║                                                    ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "💡 Troubleshooting:" -ForegroundColor Cyan
    Write-Host "  1. Check if all NuGet packages are restored" -ForegroundColor Gray
    Write-Host "  2. Verify Visual Studio installation" -ForegroundColor Gray
    Write-Host "  3. Run: nuget.exe restore OutlookOkan.sln" -ForegroundColor Gray
    Write-Host "  4. Run this script again with -Verbose for more details" -ForegroundColor Gray
    Write-Host ""
    exit $buildExitCode
}
