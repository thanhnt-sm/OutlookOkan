#!/bin/bash
# ==============================================================================
# OutlookOkan Cross-Platform Build Script
# ==============================================================================
# SYNOPSIS
#   Builds the OutlookOkan VSTO Add-in on Windows (natively) or macOS/Linux (via Mono).
#
# DESCRIPTION
#   This script detects the operating system and available build tools.
#
#   - WINDOWS: Uses MSBuild from Visual Studio
#   - MAC/LINUX: Uses Mono msbuild (required for .NET Framework VSTO projects)
#
# REQUIREMENTS (macOS)
#   brew install mono
#   brew install --cask visual-studio  # For Office Interop assemblies
#
# USAGE
#   ./build.sh [Debug|Release]
# ==============================================================================

set -e  # Exit on error

# Configuration
SOLUTION_FILE="./OutlookOkan.sln"
CONFIGURATION="${1:-Release}"
EXIT_CODE=0

echo "╔════════════════════════════════════════════════════════════╗"
echo "║            OutlookOkan Build Script                        ║"
echo "╚════════════════════════════════════════════════════════════╝"
echo ""

OS_NAME=$(uname -s)
echo "📍 OS: $OS_NAME"
echo "📦 Configuration: $CONFIGURATION"
echo ""

# ------------------------------------------------------------------------------
# WINDOWS EXECUTION (Git Bash, WSL, Cygwin)
# ------------------------------------------------------------------------------
if [[ "$OS_NAME" == *"MINGW"* ]] || [[ "$OS_NAME" == *"CYGWIN"* ]] || [[ "$OS_NAME" == *"MSYS"* ]]; then
    echo "  Windows-like environment detected"
    
    # Try to find MSBuild in common locations
    MSBUILD_PATHS=(
        "C:/Program Files/Microsoft Visual Studio/18/Enterprise/MSBuild/Current/Bin/MSBuild.exe"
        "C:/Program Files/Microsoft Visual Studio/2022/Enterprise/MSBuild/Current/Bin/MSBuild.exe"
        "C:/Program Files/Microsoft Visual Studio/2022/Professional/MSBuild/Current/Bin/MSBuild.exe"
        "C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe"
        "C:/Program Files (x86)/Microsoft Visual Studio/2019/Enterprise/MSBuild/Current/Bin/MSBuild.exe"
        "C:/Program Files (x86)/Microsoft Visual Studio/2019/Professional/MSBuild/Current/Bin/MSBuild.exe"
        "C:/Program Files (x86)/Microsoft Visual Studio/2019/Community/MSBuild/Current/Bin/MSBuild.exe"
    )

    # Also check msbuild_path.txt
    if [ -f "./msbuild_path.txt" ]; then
        CUSTOM_PATH=$(head -1 ./msbuild_path.txt | tr -d '\r\n')
        if [ -n "$CUSTOM_PATH" ]; then
            MSBUILD_PATHS=("$CUSTOM_PATH" "${MSBUILD_PATHS[@]}")
        fi
    fi
    
    MSBUILD_PATH=""
    for path in "${MSBUILD_PATHS[@]}"; do
        if [ -f "$path" ]; then
            MSBUILD_PATH="$path"
            break
        fi
    done

    # Find .NET Framework reference assemblies
    FW_OVERRIDE=""
    for ver in "v4.6.2" "v4.6.1" "v4.6"; do
        FW_PATH="C:/Program Files (x86)/Reference Assemblies/Microsoft/Framework/.NETFramework/$ver"
        if [ -d "$FW_PATH" ]; then
            FW_OVERRIDE="$FW_PATH"
            echo "  Found .NET Framework $ver"
            break
        fi
    done

    # NuGet restore
    if [ -f "./nuget.exe" ]; then
        echo "  Restoring NuGet packages..."
        ./nuget.exe restore "$SOLUTION_FILE" -Verbosity quiet
    fi

    FW_ARG=""
    if [ -n "$FW_OVERRIDE" ]; then
        FW_ARG="-p:FrameworkPathOverride=$FW_OVERRIDE"
    fi
    
    if [ -n "$MSBUILD_PATH" ]; then
        echo "  Found MSBuild: $MSBUILD_PATH"
        "$MSBUILD_PATH" "$SOLUTION_FILE" -t:Rebuild -p:Configuration=$CONFIGURATION $FW_ARG -v:m
        EXIT_CODE=$?
    elif command -v msbuild &> /dev/null; then
        echo "  Using msbuild from PATH"
        msbuild "$SOLUTION_FILE" -t:Rebuild -p:Configuration=$CONFIGURATION $FW_ARG -v:m
        EXIT_CODE=$?
    else
        echo "  ERROR: MSBuild not found. Please install Visual Studio."
        exit 1
    fi

# ------------------------------------------------------------------------------
# MACOS / LINUX EXECUTION
# ------------------------------------------------------------------------------
else
    echo "🍎 Unix-like environment detected (macOS/Linux)"
    echo ""
    
    # Check for Mono msbuild (REQUIRED for .NET Framework VSTO projects)
    if command -v msbuild &> /dev/null; then
        echo "✅ Found Mono msbuild"
        
        # Set FrameworkPathOverride if not set (helps Mono find assemblies)
        if [ -z "$FrameworkPathOverride" ]; then
            MONO_LIB=$(dirname $(which mono))/../lib/mono/4.5/
            if [ -d "$MONO_LIB" ]; then
                export FrameworkPathOverride="$MONO_LIB"
                echo "📌 Set FrameworkPathOverride=$FrameworkPathOverride"
            fi
        fi
        
        echo ""
        echo "🔨 Building with Mono..."
        echo "────────────────────────────────────────────────────────────"
        msbuild "$SOLUTION_FILE" /t:Rebuild /p:Configuration=$CONFIGURATION /v:m
        EXIT_CODE=$?
        
    else
        echo "❌ ERROR: Mono msbuild not found."
        echo ""
        echo "💡 To install on macOS:"
        echo "   brew install mono"
        echo ""
        echo "⚠️  Note: VSTO projects require .NET Framework which is Windows-only."
        echo "   Mono provides partial compatibility but may not build all targets."
        echo "   For full build support, use Windows with Visual Studio."
        exit 1
    fi
fi

# ------------------------------------------------------------------------------
# BUILD RESULT
# ------------------------------------------------------------------------------
echo ""
echo "────────────────────────────────────────────────────────────"
if [ $EXIT_CODE -eq 0 ]; then
    echo "╔════════════════════════════════════════════════════════════╗"
    echo "║              ✅ BUILD SUCCESSFUL                           ║"
    echo "╚════════════════════════════════════════════════════════════╝"
    echo ""
    echo "📍 Output: OutlookOkan/bin/$CONFIGURATION/OutlookOkan.dll"
else
    echo "╔════════════════════════════════════════════════════════════╗"
    echo "║              ❌ BUILD FAILED (Exit: $EXIT_CODE)               ║"
    echo "╚════════════════════════════════════════════════════════════╝"
    echo ""
    if [[ "$OS_NAME" == "Darwin" ]]; then
        echo "⚠️  Note: VSTO projects have Windows-specific dependencies."
        echo "   Common issues on macOS:"
        echo "   - Missing 'Microsoft.VisualStudio.Tools.Office.targets'"
        echo "   - Missing Office Interop assemblies"
        echo ""
        echo "💡 Consider building on Windows for full VSTO support."
    fi
fi

exit $EXIT_CODE
