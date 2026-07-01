#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_NAME="Mac PPTX Converter"
APP_DIR="${SCRIPT_DIR}/dist/${APP_NAME}.app"

echo ""
echo "=== Creating ${APP_NAME}.app ==="
echo ""

# Clean previous
rm -rf "$APP_DIR"

# Create .app directory structure
mkdir -p "${APP_DIR}/Contents/MacOS"
mkdir -p "${APP_DIR}/Contents/Resources/app/mac_pptx_converter"

# ---- Info.plist ----
cat > "${APP_DIR}/Contents/Info.plist" << 'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key>
    <string>Mac PPTX Converter</string>
    <key>CFBundleDisplayName</key>
    <string>Mac PPTX Converter</string>
    <key>CFBundleIdentifier</key>
    <string>com.macpptx.converter</string>
    <key>CFBundleVersion</key>
    <string>1.0.0</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0.0</string>
    <key>CFBundleExecutable</key>
    <string>MacPPTXConverter</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.15</string>
    <key>NSHighResolutionCapable</key>
    <true/>
    <key>CFBundleDocumentTypes</key>
    <array>
        <dict>
            <key>CFBundleTypeName</key>
            <string>Office Macro Files</string>
            <key>CFBundleTypeRole</key>
            <string>Editor</string>
            <key>CFBundleTypeExtensions</key>
            <array>
                <string>pptm</string>
                <string>pptx</string>
                <string>xlsm</string>
                <string>xlsx</string>
                <string>docm</string>
                <string>docx</string>
            </array>
        </dict>
    </array>
</dict>
</plist>
PLIST

# ---- Launcher shell script ----
cat > "${APP_DIR}/Contents/MacOS/MacPPTXConverter" << 'LAUNCHER'
#!/bin/bash

# Find the Resources directory
RESOURCES="$(dirname "$0")/../Resources"
APP_ROOT="${RESOURCES}/app"
VENV="${APP_ROOT}/.venv"
LOG="${APP_ROOT}/launch.log"

# Find Python 3
find_python() {
    # Homebrew (Apple Silicon)
    if [ -x "/opt/homebrew/bin/python3" ]; then
        echo "/opt/homebrew/bin/python3"; return
    fi
    # Homebrew (Intel)
    if [ -x "/usr/local/bin/python3" ]; then
        echo "/usr/local/bin/python3"; return
    fi
    # Xcode Command Line Tools / system
    if command -v python3 &>/dev/null; then
        echo "python3"; return
    fi
    return 1
}

PYTHON=$(find_python) || {
    osascript -e 'display dialog "Python 3 が見つかりません。\n\nHomebrew でインストールしてください:\n  brew install python3\n\nまたは python.org からダウンロード:\n  https://www.python.org/downloads/" buttons {"OK"} default button "OK" with title "Mac PPTX Converter" with icon stop'
    exit 1
}

# Create venv on first run
if [ ! -d "$VENV" ]; then
    osascript -e 'display notification "初回セットアップ中... 数秒お待ちください" with title "Mac PPTX Converter"' &
    "$PYTHON" -m venv "$VENV" 2>"$LOG" || {
        osascript -e "display dialog \"venv の作成に失敗しました。\n\nログ: ${LOG}\" buttons {\"OK\"} default button \"OK\" with title \"Mac PPTX Converter\" with icon stop"
        exit 1
    }
    "$VENV/bin/pip" install -q olefile >>"$LOG" 2>&1 || {
        osascript -e "display dialog \"依存パッケージのインストールに失敗しました。\n\nログ: ${LOG}\" buttons {\"OK\"} default button \"OK\" with title \"Mac PPTX Converter\" with icon stop"
        rm -rf "$VENV"
        exit 1
    }
fi

# Run the GUI
exec "$VENV/bin/python3" "${APP_ROOT}/launcher.py" "$@" 2>"$LOG"
LAUNCHER

chmod +x "${APP_DIR}/Contents/MacOS/MacPPTXConverter"

# ---- Copy Python source into the .app ----
cp "${SCRIPT_DIR}/launcher.py" "${APP_DIR}/Contents/Resources/app/"

for f in __init__.py cli.py converter.py gui.py ole_builder.py transforms.py vba_compress.py vba_parser.py; do
    cp "${SCRIPT_DIR}/mac_pptx_converter/${f}" "${APP_DIR}/Contents/Resources/app/mac_pptx_converter/"
done

# Copy __main__.py if it exists
[ -f "${SCRIPT_DIR}/mac_pptx_converter/__main__.py" ] && \
    cp "${SCRIPT_DIR}/mac_pptx_converter/__main__.py" "${APP_DIR}/Contents/Resources/app/mac_pptx_converter/"

# ---- Remove quarantine attribute ----
xattr -cr "$APP_DIR" 2>/dev/null || true

echo "=== Done ==="
echo ""
echo "  App: dist/Mac PPTX Converter.app"
echo ""
echo "  Open:"
echo "    open \"dist/Mac PPTX Converter.app\""
echo ""
echo "  Install to Applications:"
echo "    cp -R \"dist/Mac PPTX Converter.app\" /Applications/"
echo ""
echo "  If macOS still blocks it:"
echo "    xattr -cr \"dist/Mac PPTX Converter.app\""
echo ""
