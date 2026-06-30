#!/bin/bash
set -e

# Build Mac PPTX Converter as a macOS .app bundle
# Run this on your Mac: ./build_app.sh

APP_NAME="Mac PPTX Converter"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "=== Building ${APP_NAME}.app ==="

# Create venv if needed
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
fi

source .venv/bin/activate

echo "Installing dependencies..."
pip install -q --upgrade pip
pip install -q olefile pyinstaller

# Optional: better drag-and-drop support
pip install -q tkinterdnd2 2>/dev/null || echo "(tkinterdnd2 not available — file dialog will be used instead of drag-and-drop)"

echo "Building .app bundle..."
pyinstaller \
    --name "Mac PPTX Converter" \
    --windowed \
    --onedir \
    --icon icon.icns 2>/dev/null || \
pyinstaller \
    --name "Mac PPTX Converter" \
    --windowed \
    --onedir \
    --add-data "mac_pptx_converter:mac_pptx_converter" \
    --hidden-import mac_pptx_converter \
    --hidden-import mac_pptx_converter.converter \
    --hidden-import mac_pptx_converter.transforms \
    --hidden-import mac_pptx_converter.vba_parser \
    --hidden-import mac_pptx_converter.vba_compress \
    --hidden-import mac_pptx_converter.ole_builder \
    --hidden-import olefile \
    --osx-bundle-identifier com.macpptx.converter \
    mac_pptx_converter/gui.py

# Add file type associations to Info.plist
APP_BUNDLE="dist/Mac PPTX Converter.app"
PLIST="${APP_BUNDLE}/Contents/Info.plist"

if [ -f "$PLIST" ] && command -v /usr/libexec/PlistBuddy &>/dev/null; then
    echo "Adding file type associations..."
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes array" "$PLIST" 2>/dev/null || true
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0 dict" "$PLIST" 2>/dev/null || true
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:CFBundleTypeName string 'Office Macro Files'" "$PLIST" 2>/dev/null || true
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:CFBundleTypeRole string 'Editor'" "$PLIST" 2>/dev/null || true
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:LSItemContentTypes array" "$PLIST" 2>/dev/null || true
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:LSItemContentTypes:0 string 'org.openxmlformats.presentationml.presentation.macroenabled'" "$PLIST" 2>/dev/null || true
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions array" "$PLIST" 2>/dev/null || true
    for ext in pptm pptx xlsm xlsx docm docx; do
        /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions: string '${ext}'" "$PLIST" 2>/dev/null || true
    done
fi

echo ""
echo "=== Build complete ==="
echo ""
echo "  App location: dist/Mac PPTX Converter.app"
echo ""
echo "  To install, drag it to /Applications:"
echo "    cp -R \"dist/Mac PPTX Converter.app\" /Applications/"
echo ""
echo "  Or run directly:"
echo "    open \"dist/Mac PPTX Converter.app\""
echo ""
