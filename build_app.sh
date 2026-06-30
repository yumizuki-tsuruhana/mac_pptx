#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo ""
echo "=== Mac PPTX Converter — .app build ==="
echo ""

# ---- venv ----
if [ ! -d ".venv" ]; then
    echo "[1/4] Creating virtual environment..."
    python3 -m venv .venv
else
    echo "[1/4] Virtual environment exists"
fi
source .venv/bin/activate

# ---- deps ----
echo "[2/4] Installing dependencies..."
pip install -q --upgrade pip
pip install -q olefile pyinstaller
pip install -q -e .

# ---- clean previous build ----
rm -rf build dist

# ---- build ----
echo "[3/4] Building .app bundle..."
pyinstaller \
    --name "Mac PPTX Converter" \
    --windowed \
    --onedir \
    --noconfirm \
    --collect-submodules mac_pptx_converter \
    --hidden-import mac_pptx_converter.converter \
    --hidden-import mac_pptx_converter.transforms \
    --hidden-import mac_pptx_converter.vba_parser \
    --hidden-import mac_pptx_converter.vba_compress \
    --hidden-import mac_pptx_converter.ole_builder \
    --hidden-import mac_pptx_converter.gui \
    --hidden-import olefile \
    --osx-bundle-identifier com.macpptx.converter \
    launcher.py

# ---- post-build: file type associations ----
echo "[4/4] Configuring .app metadata..."
APP_BUNDLE="dist/Mac PPTX Converter.app"
PLIST="${APP_BUNDLE}/Contents/Info.plist"

if [ -f "$PLIST" ] && command -v /usr/libexec/PlistBuddy &>/dev/null; then
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes array" "$PLIST" 2>/dev/null || true
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0 dict" "$PLIST" 2>/dev/null || true
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:CFBundleTypeName string 'Office Macro Files'" "$PLIST" 2>/dev/null || true
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:CFBundleTypeRole string 'Editor'" "$PLIST" 2>/dev/null || true
    /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions array" "$PLIST" 2>/dev/null || true
    for ext in pptm pptx xlsm xlsx docm docx; do
        /usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions: string '${ext}'" "$PLIST" 2>/dev/null || true
    done
fi

echo ""
echo "=== Build complete ==="
echo ""
echo "  .app is at:"
echo "    dist/Mac PPTX Converter.app"
echo ""
echo "  To open:"
echo "    open \"dist/Mac PPTX Converter.app\""
echo ""
echo "  To install to Applications:"
echo "    cp -R \"dist/Mac PPTX Converter.app\" /Applications/"
echo ""
echo "  -------------------------------------------------------"
echo "  If macOS blocks the app (\"damaged\" or \"unidentified\"):"
echo ""
echo "    xattr -cr \"dist/Mac PPTX Converter.app\""
echo ""
echo "  Then open again."
echo "  -------------------------------------------------------"
echo ""
