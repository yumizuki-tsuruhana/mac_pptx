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

RESOURCES="$(cd "$(dirname "$0")/../Resources" && pwd)"
APP_ROOT="${RESOURCES}/app"
VENV="${APP_ROOT}/.venv"
LOG="${APP_ROOT}/launch.log"

echo "=== $(date) ===" > "$LOG"

show_error() {
    osascript -e "display dialog \"$1\" buttons {\"OK\"} default button \"OK\" with title \"Mac PPTX Converter\" with icon stop"
    open -e "$LOG" 2>/dev/null || true
}

# Find a python3 that actually has tkinter (Homebrew python3 often lacks it
# unless python-tk is installed separately, and fails import silently).
CANDIDATES=(
    "/opt/homebrew/bin/python3"
    "/usr/local/bin/python3"
    "/usr/bin/python3"
    "python3"
)

PYTHON=""
for c in "${CANDIDATES[@]}"; do
    if command -v "$c" &>/dev/null; then
        if "$c" -c "import tkinter" >>"$LOG" 2>&1; then
            PYTHON="$c"
            echo "Using python: $c" >> "$LOG"
            break
        else
            echo "Skipping $c (no tkinter)" >> "$LOG"
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    show_error "Python3 + tkinter が見つかりません。\n\nターミナルで以下を実行してください:\n\n  brew install python-tk\n\nHomebrew が無い場合は python.org のインストーラーを使ってください:\n  https://www.python.org/downloads/"
    exit 1
fi

# (Re)create venv if missing or broken
if [ ! -x "${VENV}/bin/python3" ]; then
    rm -rf "$VENV"
    "$PYTHON" -m venv "$VENV" >>"$LOG" 2>&1
fi

if [ ! -x "${VENV}/bin/python3" ]; then
    show_error "セットアップ(venv作成)に失敗しました。ログを開きます。"
    exit 1
fi

if ! "${VENV}/bin/python3" -c "import tkinter" >>"$LOG" 2>&1; then
    show_error "venv内でtkinterが使えません。ログを開きます。"
    exit 1
fi

if ! "${VENV}/bin/python3" -c "import olefile" >>"$LOG" 2>&1; then
    "${VENV}/bin/pip" install -q olefile >>"$LOG" 2>&1
    if ! "${VENV}/bin/python3" -c "import olefile" >>"$LOG" 2>&1; then
        show_error "依存パッケージ(olefile)のインストールに失敗しました。ログを開きます。"
        exit 1
    fi
fi

"${VENV}/bin/python3" "${APP_ROOT}/launcher.py" "$@" >>"$LOG" 2>&1
EXIT_CODE=$?

if [ $EXIT_CODE -ne 0 ]; then
    show_error "アプリの起動に失敗しました (exit code ${EXIT_CODE})。ログを開きます。"
    exit 1
fi
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
