<#
.SYNOPSIS
    Build script for OutlookOkan VSTO Add-in.
    
.DESCRIPTION
    Automatically locates MSBuild, restores NuGet packages,
    and rebuilds the solution with proper .NET Framework targeting.
    
.PARAMETER Configuration
    Build configuration: Debug or Release (default: Release)
    
.PARAMETER Verbose
    Enable verbose output (default: quiet)
    
.EXAMPLE
    .\build.ps1                        # Build Release configuration
    .\build.ps1 -Configuration Debug   # Build Debug configuration
    .\build.ps1 -Verbose               # Build Release with verbose output

.NOTES
    Author: thanhnt-sm
    Version: 2.0
#>

param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    
    [switch]$Verbose
)

# Error handling
$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "          OutlookOkan Build Script v2.0                 " -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

# --- Step 1: Find MSBuild ---
$msbuildPaths = @(
    "C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
)

# Also check msbuild_path.txt if it exists
$msbuildPathFile = Join-Path $PSScriptRoot "msbuild_path.txt"
if (Test-Path $msbuildPathFile) {
    $customPath = (Get-Content $msbuildPathFile -First 1).Trim()
    if ($customPath) {
        $msbuildPaths = @($customPath) + $msbuildPaths
    }
}

$msbuildPath = $null
Write-Host "[1/4] Searching for MSBuild..." -ForegroundColor Yellow

foreach ($path in $msbuildPaths) {
    if (Test-Path -Path $path) {
        $msbuildPath = $path
        Write-Host "  OK: Found MSBuild at: $path" -ForegroundColor Green
        break
    }
}

if ($null -eq $msbuildPath) {
    Write-Host ""
    Write-Host "  FAIL: MSBuild not found!" -ForegroundColor Red
    Write-Host "  Checked locations:" -ForegroundColor Yellow
    $msbuildPaths | ForEach-Object { Write-Host "    - $_" -ForegroundColor Gray }
    Write-Host ""
    Write-Host "  Install Visual Studio with MSBuild tools." -ForegroundColor Cyan
    exit 1
}

# --- Step 2: Find .NET Framework Reference Assemblies ---
Write-Host "[2/4] Detecting .NET Framework 4.6.2..." -ForegroundColor Yellow

$frameworkBasePath = "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework"
$frameworkPathOverride = $null

# Prefer v4.6.2 (project target), fallback to other 4.6.x versions
$frameworkVersions = @("v4.6.2", "v4.6.1", "v4.6")

foreach ($ver in $frameworkVersions) {
    $testPath = Join-Path $frameworkBasePath $ver
    if (Test-Path $testPath) {
        $frameworkPathOverride = $testPath
        Write-Host "  OK: Found .NET Framework $ver at: $testPath" -ForegroundColor Green
        break
    }
}

if ($null -eq $frameworkPathOverride) {
    Write-Host "  WARN: .NET Framework reference assemblies not found." -ForegroundColor Yellow
    Write-Host "  Build may fail. Install .NET Framework 4.6.2 Developer Pack." -ForegroundColor Yellow
} else {
    Write-Host "  Using FrameworkPathOverride: $frameworkPathOverride" -ForegroundColor Gray
}

# --- Step 3: NuGet Restore ---
Write-Host "[3/4] Restoring NuGet packages..." -ForegroundColor Yellow

$solutionFile = Join-Path $PSScriptRoot "OutlookOkan.sln"

if (-not (Test-Path -Path $solutionFile)) {
    Write-Host "  FAIL: Solution file not found: $solutionFile" -ForegroundColor Red
    exit 1
}

$nugetExe = Join-Path $PSScriptRoot "nuget.exe"
if (Test-Path $nugetExe) {
    & $nugetExe restore $solutionFile -Verbosity quiet
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  OK: NuGet packages restored." -ForegroundColor Green
    } else {
        Write-Host "  WARN: NuGet restore returned exit code $LASTEXITCODE" -ForegroundColor Yellow
    }
} else {
    Write-Host "  SKIP: nuget.exe not found, skipping restore." -ForegroundColor Yellow
    Write-Host "  Run manually: nuget.exe restore OutlookOkan.sln" -ForegroundColor Gray
}

# --- Step 4: Build ---
Write-Host "[4/4] Building solution ($Configuration)..." -ForegroundColor Yellow
Write-Host "--------------------------------------------------------" -ForegroundColor Gray
Write-Host ""

$verbosityLevel = if ($Verbose) { "normal" } else { "minimal" }

# Build MSBuild arguments
$msbuildArgs = @(
    $solutionFile,
    "/t:Rebuild",
    "/p:Configuration=$Configuration",
    "/v:$verbosityLevel",
    "/nologo"
)

# Add FrameworkPathOverride if detected
if ($frameworkPathOverride) {
    $msbuildArgs += "/p:FrameworkPathOverride=$frameworkPathOverride"
}

try {
    & $msbuildPath @msbuildArgs
    $buildExitCode = $LASTEXITCODE
}
catch {
    Write-Host ""
    Write-Host "  FAIL: Failed to execute MSBuild" -ForegroundColor Red
    Write-Host "  Message: $_" -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "--------------------------------------------------------" -ForegroundColor Gray

# Check build result
if ($buildExitCode -eq 0) {
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Green
    Write-Host "        BUILD SUCCESSFUL ($Configuration)               " -ForegroundColor Green
    Write-Host "========================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Output:" -ForegroundColor Cyan
    Write-Host "  Main DLL: OutlookOkan\bin\$Configuration\OutlookOkan.dll" -ForegroundColor Gray
    Write-Host ""
    exit 0
}
else {
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host "        BUILD FAILED (Exit Code: $buildExitCode)        " -ForegroundColor Red
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Cyan
    Write-Host "  1. Run: nuget.exe restore OutlookOkan.sln" -ForegroundColor Gray
    Write-Host "  2. Verify VSTO OfficeTools installed in VS" -ForegroundColor Gray
    Write-Host "  3. Install .NET Framework 4.6.2 Developer Pack" -ForegroundColor Gray
    Write-Host "  4. Run: .\build.ps1 -Verbose for more details" -ForegroundColor Gray
    Write-Host ""
    exit $buildExitCode
}
