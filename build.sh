#!/bin/bash
# build.sh - Build script for local development

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "Building PplExtractor in Release mode..."
dotnet build -c Release

echo ""
echo "Build completed successfully!"
echo "Output: bin/Release/net9.0/PplExtractor"


