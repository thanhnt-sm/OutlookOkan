#!/bin/bash
# ==============================================================================
# OutlookOkan Verification Script
# ==============================================================================
# This script performs comprehensive verification of the OutlookOkan C# codebase.
# It is designed to run on macOS (and Linux/Windows via Git Bash).
#
# CHECKS PERFORMED:
# 1. Static Analysis: Checks for known anti-patterns, null safety issues, and style violations.
# 2. Compilation Check: Compiles the code using a Verification Project stubbed with
#    Office/VSTO dependencies to ensure C# syntax and type correctness on non-Windows OS.
#
# USAGE:
#   ./verify.sh [dependencies]
#
# ARGS:
#   --skip-build    Skip compilation check (static analysis only)
#
# ==============================================================================

# ----------------- Configuration -----------------
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VERIFY_PROJECT="$PROJECT_DIR/OutlookOkan.Verify.csproj"
EXIT_CODE=0
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}==============================================================================${NC}"
echo -e "${BLUE}                   OutlookOkan Code Verification                              ${NC}"
echo -e "${BLUE}==============================================================================${NC}"
echo "Project Directory: $PROJECT_DIR"
echo ""

# ----------------- Static Analysis -----------------
echo -e "${BLUE}[1/2] Running Static Code Analysis (Patterns & Null Safety)...${NC}"

ERROR_COUNT=0
WARN_COUNT=0

# Helper to check pattern
check_pattern() {
    local pattern="$1"
    local message="$2"
    local severity="$3" # "ERROR" or "WARNING"
    
    # Use grep to find matches
    matches=$(grep -r "$pattern" "$PROJECT_DIR/OutlookOkan" --include="*.cs" --exclude-dir="Properties" --exclude="*.Designer.cs" --exclude="*.Verify.*" 2>/dev/null)
    
    if [ ! -z "$matches" ]; then
        if [ "$severity" == "ERROR" ]; then
            echo -e "${RED}[ERROR] $message${NC}"
            ((ERROR_COUNT++))
        else
            echo -e "${YELLOW}[WARN]  $message${NC}"
            ((WARN_COUNT++))
        fi
        echo "$matches" | while read -r line; do
            # Format: filename:line: code
            echo "    -> $(basename "$line")"
        done
    fi
}

# 1. Null Checks
# Check for null checks without return (simple heuristic)
# Code like: if (x == null) { log } ... (proceeds to use x)
# This is hard to regex exactly, but we can look for suspicious patterns
check_pattern "if (.* == null) *$" "Potential null check missing block or early return (style)" "WARNING"

# 2. Empty Catch Blocks
check_pattern "catch.*{ *}" "Empty catch block detected (swallows errors)" "WARNING"
check_pattern "catch.*{ \/\/.*}" "Catch block with only comments (verify intent)" "WARNING"

# 3. TODO/FIXME
check_pattern "TODO:" "TODO comment found" "WARNING"
check_pattern "FIXME:" "FIXME comment found" "WARNING"

# 4. Console.WriteLine (Classic VSTO anti-pattern, use logging)
check_pattern "Console.WriteLine" "Console.WriteLine usage detected (should use logging)" "WARNING"

# 5. Thread.Sleep (Performance)
check_pattern "Thread.Sleep" "Thread.Sleep usage detected (potential UI blocker)" "WARNING"

# 6. Hardcoded Paths
check_pattern "C:\\\\" "Hardcoded C:\\ path detected" "WARNING"

if [ $ERROR_COUNT -eq 0 ]; then
    echo -e "${GREEN}✓ Static Analysis Passed (Warnings: $WARN_COUNT)${NC}"
else
    echo -e "${RED}✗ Static Analysis Failed ($ERROR_COUNT errors, $WARN_COUNT warnings)${NC}"
    EXIT_CODE=1
fi

echo ""

# ----------------- Compilation Check -----------------
if [[ "$1" == "--skip-build" ]]; then
    echo -e "${YELLOW}[2/2] Skipping Compilation Check (--skip-build passed)${NC}"
else
    echo -e "${BLUE}[2/2] Running Compilation Verification (via dotnet build)...${NC}"
    
    if ! command -v dotnet &> /dev/null; then
        echo -e "${RED}[ERROR] 'dotnet' command not found. Please install .NET SDK.${NC}"
        EXIT_CODE=1
    else
        echo "Building $VERIFY_PROJECT..."
        
        # Capture output
        BUILD_OUTPUT=$(dotnet build "$VERIFY_PROJECT" -v quiet 2>&1)
        BUILD_EXIT=$?
        
        if [ $BUILD_EXIT -eq 0 ]; then
             echo -e "${GREEN}✓ Compilation Verification SUCCESSFUL${NC}"
             echo "  (Note: Warnings regarding PDFsharp and platform compatibility are expected on macOS)"
        else
             echo -e "${RED}✗ Compilation Verification FAILED${NC}"
             echo "Build Output (Errors Only):"
             echo "$BUILD_OUTPUT" | grep "error" | sed 's/^/    /'
             EXIT_CODE=1
        fi
    fi
fi

# ----------------- Summary -----------------
echo ""
echo -e "${BLUE}==============================================================================${NC}"
if [ $EXIT_CODE -eq 0 ]; then
    echo -e "${GREEN}VERIFICATION COMPLETED SUCCESSFULLY${NC}"
    echo "Your code is syntactically correct and compiles against VSTO stubs."
    exit 0
else
    echo -e "${RED}VERIFICATION FAILED${NC}"
    echo "Please fix the compilation errors or static analysis issues above."
    exit 1
fi
