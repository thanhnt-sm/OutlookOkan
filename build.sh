#!/bin/bash
# ==============================================================================
# OutlookOkan Cross-Platform Build Script
# ==============================================================================
# SYNOPSIS
#   Builds the OutlookOkan VSTO Add-in on Windows (natively) or macOS/Linux (via Mono/Dotnet).
#
# DESCRIPTION
#   This script detects the operating system and available build tools to attempt
#   a successful build.
#
#   - WINDOWS: Uses specific MSBuild path or auto-detects.
#   - MAC/LINUX: Prioritizes Mono `msbuild` (best for legacy .NET Framework).
#                Falls back to `dotnet build` (requires SDK-style projects, may fail for VSTO).
#
# USAGE
#   ./build.sh
# ==============================================================================

# User-defined MSBuild path (Windows only)
WINDOWS_MSBUILD_PATH="C:/Program Files/Microsoft Visual Studio/18/Enterprise/MSBuild/Current/Bin/MSBuild.exe"
SOLUTION_FILE="./OutlookOkan.sln"
CONFIGURATION="Release"

echo "================================================================"
echo "  OutlookOkan Build Wrapper"
echo "================================================================"

OS_NAME=$(uname -s)
echo "[INFO] Detected OS: $OS_NAME"

# ------------------------------------------------------------------------------
# 1. WINDOWS EXECUTION (Git Bash, WSL, Cygwin)
# ------------------------------------------------------------------------------
if [[ "$OS_NAME" == *"MINGW"* ]] || [[ "$OS_NAME" == *"CYGWIN"* ]] || [[ "$OS_NAME" == *"MSYS"* ]]; then
    echo "[INFO] Window-like environment detected."
    
    if [ -f "$WINDOWS_MSBUILD_PATH" ]; then
        echo "[INFO] Found specific MSBuild: $WINDOWS_MSBUILD_PATH"
        "$WINDOWS_MSBUILD_PATH" "$SOLUTION_FILE" -t:Rebuild -p:Configuration=$CONFIGURATION -v:m
    else
        echo "[WARN] Specific MSBuild not found. Checking PATH..."
        if command -v msbuild &> /dev/null; then
            msbuild "$SOLUTION_FILE" -t:Rebuild -p:Configuration=$CONFIGURATION -v:m
        else
            echo "[ERROR] MSBuild not found. Please install Visual Studio."
            exit 1
        fi
    fi

# ------------------------------------------------------------------------------
# 2. MACOS / LINUX EXECUTION
# ------------------------------------------------------------------------------
else
    echo "[INFO] Unix-like environment detected."

    # A. Check for Mono (Recommended for VSTO/.NET Framework legacy projects)
    if command -v msbuild &> /dev/null; then
        echo "[INFO] Found Mono 'msbuild'. This is the best option for .NET Framework on macOS."
        echo "       Attempting build with Mono..."
        
        # Check for FrameworkPathOverride (common fix for Mono finding assemblies)
        if [ -z "$FrameworkPathOverride" ]; then
             echo "[HINT] If build fails on 'mscorlib', try setting: export FrameworkPathOverride=\$(dirname \$(which mono))/../lib/mono/4.5/"
        fi

        msbuild "$SOLUTION_FILE" /t:Rebuild /p:Configuration=$CONFIGURATION /v:m
        
    # B. Check for Dotnet SDK (Fallback)
    elif command -v dotnet &> /dev/null; then
        echo "[WARN] Mono not found. Fallback to 'dotnet' CLI."
        echo "[WARN] This is a legacy .NET Framework VSTO project."
        echo "       'dotnet build' will likely FAIL unless the project is migrated to SDK-style."
        echo "       Current .NET SDK: $(dotnet --version)"
        
        echo "Attempting 'dotnet restore'..."
        dotnet restore "$SOLUTION_FILE"
        
        echo "Attempting 'dotnet build'..."
        dotnet build "$SOLUTION_FILE" --configuration $CONFIGURATION
        
    else
        echo "[ERROR] No build tools found."
        echo "       Please install 'Mono' (recommended) or '.NET SDK'."
        echo "       Brew: brew install mono"
        exit 1
    fi
fi

EXIT_CODE=$?
echo ""
if [ $EXIT_CODE -eq 0 ]; then
    echo "✅ BUILD SUCCESSFUL"
else
    echo "❌ BUILD FAILED (Exit Code: $EXIT_CODE)"
    if [[ "$OS_NAME" != *"MINGW"* ]]; then
        echo "NOTE: VSTO projects heavily rely on Windows COM targets."
        echo "      A failure on macOS is expected if missing 'Microsoft.VisualStudio.Tools.Office.targets'."
    fi
fi

exit $EXIT_CODE
