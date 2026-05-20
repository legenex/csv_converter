#!/usr/bin/env bash
# Build an unsigned .dmg installer for the CSV to XLSX Converter.
#
# Stages 1 and 2 (PyInstaller build + DMG) run by default.
# Stages "sign" and "notarize" are commented out and ready to enable
# once you have an Apple Developer ID.
#
# Requirements (macOS only):
#   - Xcode command-line tools:  xcode-select --install
#   - Homebrew + create-dmg:     brew install create-dmg
# PyInstaller is installed into .venv by this script if missing.

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$HERE"

APP_NAME="CSV to XLSX Converter"
BUNDLE_ID="com.legenex.csvconverter"

# Uncomment once you have a Developer ID set up:
# CODESIGN_IDENTITY="Developer ID Application: Your Name (TEAMID)"
# NOTARY_PROFILE="legenex-notary"   # see `xcrun notarytool store-credentials`

# ---- Sanity checks --------------------------------------------------------
if [[ "$(uname)" != "Darwin" ]]; then
    echo "error: this script only runs on macOS." >&2
    exit 1
fi

if [[ ! -d ".venv" ]]; then
    echo "Creating venv ..."
    python3 -m venv .venv
fi
# Invoke tools via the venv's python directly instead of `source activate`.
# This is robust even if the venv has been moved (which breaks the absolute
# paths in activate).
VENV_PY="$HERE/.venv/bin/python3"
if [[ ! -x "$VENV_PY" ]]; then
    echo "error: $VENV_PY not found or not executable." >&2
    exit 1
fi

echo "Ensuring runtime + build dependencies ..."
"$VENV_PY" -m pip install --upgrade pip >/dev/null
"$VENV_PY" -m pip install -r requirements.txt
"$VENV_PY" -m pip install --upgrade pyinstaller >/dev/null

USE_CREATE_DMG=1
if ! command -v create-dmg >/dev/null 2>&1; then
    echo "note: create-dmg not found; will use macOS-built-in hdiutil instead."
    echo "      (for nicer DMG layout: brew install create-dmg)"
    USE_CREATE_DMG=0
fi

# ---- Stage 1: PyInstaller -------------------------------------------------
echo "Cleaning previous build artifacts ..."
rm -rf build dist "$APP_NAME.spec"

echo "Stage 1/2: Building $APP_NAME.app with PyInstaller ..."
"$VENV_PY" -m PyInstaller --windowed \
    --name "$APP_NAME" \
    --osx-bundle-identifier "$BUNDLE_ID" \
    --collect-all PyQt6 \
    --collect-all polars \
    --noconfirm \
    main.py

APP_PATH="dist/$APP_NAME.app"
if [[ ! -d "$APP_PATH" ]]; then
    echo "error: build failed; $APP_PATH was not produced." >&2
    exit 1
fi

# ---- (disabled) code-sign -------------------------------------------------
# Uncomment CODESIGN_IDENTITY above and this block:
#
# echo "Code-signing $APP_PATH ..."
# codesign --deep --force --options runtime --timestamp \
#     --sign "$CODESIGN_IDENTITY" \
#     "$APP_PATH"

# ---- (disabled) notarize --------------------------------------------------
# First, store credentials once on this machine:
#   xcrun notarytool store-credentials "legenex-notary" \
#       --apple-id team@legenex.com --team-id TEAMID --password APP_SPECIFIC_PW
# Then uncomment:
#
# ZIP_PATH="dist/${APP_NAME// /_}.zip"
# ditto -c -k --keepParent "$APP_PATH" "$ZIP_PATH"
# xcrun notarytool submit "$ZIP_PATH" --keychain-profile "$NOTARY_PROFILE" --wait
# xcrun stapler staple "$APP_PATH"

# ---- Stage 2: DMG ---------------------------------------------------------
DMG_PATH="dist/$APP_NAME.dmg"
rm -f "$DMG_PATH"

echo "Stage 2/2: Building $DMG_PATH ..."

if [[ "$USE_CREATE_DMG" == "1" ]]; then
    # create-dmg can return non-zero on benign warnings (e.g. when no
    # codesign identity is available); ignore and verify by file existence.
    create-dmg \
        --volname "$APP_NAME" \
        --window-pos 200 120 \
        --window-size 600 400 \
        --icon-size 100 \
        --icon "$APP_NAME.app" 175 190 \
        --hide-extension "$APP_NAME.app" \
        --app-drop-link 425 190 \
        "$DMG_PATH" \
        "$APP_PATH" || true
else
    # Fallback: assemble a staging folder containing the app + a symlink
    # to /Applications, then turn it into a compressed read-only DMG using
    # the macOS-built-in hdiutil. This gives users the familiar
    # "drag to Applications" install flow without needing Homebrew.
    STAGE="dist/dmg_stage"
    rm -rf "$STAGE"
    mkdir -p "$STAGE"
    cp -R "$APP_PATH" "$STAGE/"
    ln -s /Applications "$STAGE/Applications"
    hdiutil create \
        -volname "$APP_NAME" \
        -srcfolder "$STAGE" \
        -ov \
        -format UDZO \
        "$DMG_PATH"
    rm -rf "$STAGE"
fi

if [[ ! -f "$DMG_PATH" ]]; then
    echo "error: DMG creation failed; $DMG_PATH was not produced." >&2
    exit 1
fi

# Optionally sign + notarize + staple the DMG itself once you have a Dev ID:
# codesign --sign "$CODESIGN_IDENTITY" --timestamp "$DMG_PATH"
# xcrun notarytool submit "$DMG_PATH" --keychain-profile "$NOTARY_PROFILE" --wait
# xcrun stapler staple "$DMG_PATH"

SIZE=$(du -h "$DMG_PATH" | awk '{print $1}')
echo
echo "Built: $DMG_PATH ($SIZE)"
echo
echo "This is an UNSIGNED build. On first launch users must either:"
echo "  1. Right-click the installed app in Finder -> Open -> Open, or"
echo "  2. Run: xattr -dr com.apple.quarantine /Applications/\"$APP_NAME.app\""
echo
echo "For a signed/notarized release, set CODESIGN_IDENTITY + NOTARY_PROFILE"
echo "at the top of this script and uncomment the sign/notarize blocks."
