#!/bin/bash
# run.sh - Local development and testing script

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "Building project..."
dotnet build

echo ""
echo "Running PplExtractor..."
echo ""

# If arguments are provided, pass them through
if [ $# -gt 0 ]; then
    dotnet run -- "$@"
else
    # Auto-detect Excel files in parent directory
    PARENT_DIR="$(dirname "$SCRIPT_DIR")"
    if [ -f "$PARENT_DIR/OPC W2 - Hamilton - Distribution 2.xls" ]; then
        dotnet run -- "$PARENT_DIR/OPC W2 - Hamilton - Distribution 2.xls"
    else
        # Try to find any Excel file in parent directory
        EXCEL_FILE=$(find "$PARENT_DIR" -maxdepth 1 -name "*.xls" -o -name "*.xlsx" | head -1)
        if [ -n "$EXCEL_FILE" ]; then
            dotnet run -- "$EXCEL_FILE"
        else
            echo "No Excel file specified and no default file found."
            echo "Usage: ./run.sh [excel_file] [options]"
            echo "Example: ./run.sh \"../OPC W2 - Hamilton - Distribution 2.xls\""
            exit 1
        fi
    fi
fi

