#!/bin/bash
# run.sh - Local development and testing script

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "Building project..."
dotnet build

echo ""
echo "Running PplExtractor (GUI application)..."
echo ""

# Run the GUI application (no command-line arguments needed)
dotnet run

