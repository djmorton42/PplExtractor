#!/bin/bash
# build-dist.sh - Build single-file executables for distribution (Mac and Windows)

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

DIST_DIR="dist"

# Clean previous builds
echo "Cleaning previous builds..."
rm -rf "$DIST_DIR"
rm -f PplExtractor-*.zip
mkdir -p "$DIST_DIR"

# Detect current platform
PLATFORM=$(uname -s)
ARCH=$(uname -m)

echo "Building distribution packages..."
echo ""

# Build for macOS (ARM64)
echo "Building for macOS (ARM64)..."
dotnet publish -c Release -r osx-arm64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
mkdir -p "$DIST_DIR/macos-arm64"
# Copy only the executable, excluding .pdb and other debug files
find bin/Release/net9.0/osx-arm64/publish -maxdepth 1 -type f ! -name "*.pdb" ! -name "*.pdf" -exec cp {} "$DIST_DIR/macos-arm64/" \;
echo "✓ macOS ARM64 build complete"

# Build for macOS (x64 - Intel)
echo "Building for macOS (x64 - Intel)..."
dotnet publish -c Release -r osx-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
mkdir -p "$DIST_DIR/macos-x64"
# Copy only the executable, excluding .pdb and other debug files
find bin/Release/net9.0/osx-x64/publish -maxdepth 1 -type f ! -name "*.pdb" ! -name "*.pdf" -exec cp {} "$DIST_DIR/macos-x64/" \;
echo "✓ macOS x64 build complete"

# Build for Windows (x64)
echo "Building for Windows (x64)..."
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
mkdir -p "$DIST_DIR/win-x64"
# Copy only the executable, excluding .pdb and other debug files
find bin/Release/net9.0/win-x64/publish -maxdepth 1 -type f ! -name "*.pdb" ! -name "*.pdf" -exec cp {} "$DIST_DIR/win-x64/" \;
echo "✓ Windows x64 build complete"

# Copy VERSION.txt to each distribution if it exists
if [ -f "VERSION.txt" ]; then
  echo "Including VERSION.txt in distributions..."
  cp VERSION.txt "$DIST_DIR/macos-arm64/"
  cp VERSION.txt "$DIST_DIR/macos-x64/"
  cp VERSION.txt "$DIST_DIR/win-x64/"
fi

# Create zip files
echo ""
echo "Creating distribution zip files..."

# macOS ARM64
cd "$DIST_DIR/macos-arm64"
zip -q -r "../../PplExtractor-macos-arm64.zip" .
cd "$SCRIPT_DIR"
echo "✓ Created PplExtractor-macos-arm64.zip"

# macOS x64
cd "$DIST_DIR/macos-x64"
zip -q -r "../../PplExtractor-macos-x64.zip" .
cd "$SCRIPT_DIR"
echo "✓ Created PplExtractor-macos-x64.zip"

# Windows x64
cd "$DIST_DIR/win-x64"
zip -q -r "../../PplExtractor-win-x64.zip" .
cd "$SCRIPT_DIR"
echo "✓ Created PplExtractor-win-x64.zip"

echo ""
echo "Distribution build complete!"
echo ""
echo "Created files:"
echo "  - PplExtractor-macos-arm64.zip"
echo "  - PplExtractor-macos-x64.zip"
echo "  - PplExtractor-win-x64.zip"
echo ""
echo "Each zip contains a standalone single-file executable with .NET runtime included."


