<#
.SYNOPSIS
    Build script for OutlookOkan VSTO Add-in.
    
.DESCRIPTION
    Uses the specific MSBuild instance found in "C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
    to rebuild the solution.
#>

$ErrorActionPreference = "Stop"

$msbuildPath = "C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
$solutionFile = Join-Path $PSScriptRoot "OutlookOkan.sln"

Write-Host "Checking for MSBuild at: $msbuildPath" -ForegroundColor Cyan

if (-not (Test-Path -Path $msbuildPath)) {
    Write-Error "MSBuild executable not found at '$msbuildPath'. Please check your Visual Studio installation path."
    exit 1
}

Write-Host "Starting Rebuild for: $solutionFile" -ForegroundColor Cyan

# Execute MSBuild
& $msbuildPath $solutionFile /t:Rebuild /p:Configuration=Release /v:m

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n----------------------------------------" -ForegroundColor Green
    Write-Host "✅ BUILD SUCCESSFUL" -ForegroundColor Green
    Write-Host "----------------------------------------" -ForegroundColor Green
}
else {
    Write-Host "`n----------------------------------------" -ForegroundColor Red
    Write-Host "❌ BUILD FAILED (Exit Code: $LASTEXITCODE)" -ForegroundColor Red
    Write-Host "----------------------------------------" -ForegroundColor Red
    exit $LASTEXITCODE
}
