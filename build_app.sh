#!/usr/bin/env bash
# ============================================================
#  MP&L Hub  –  Desktop Application Build Script (Linux/Mac)
#  Run from the project root:  bash build_app.sh
# ============================================================

set -e

echo "=== MP&L Hub Build Script ==="
echo

# -- Install / upgrade build dependencies --
echo "Installing build dependencies..."
pip install --upgrade pyinstaller

# -- Install app runtime dependencies --
echo "Installing runtime dependencies..."
pip install PyQt5 pandas numpy bcrypt

# -- Clean previous build artifacts --
echo "Cleaning previous build..."
rm -rf build dist

# -- Run PyInstaller --
echo "Building executable..."
pyinstaller "MP&L_Hub.spec"

echo
echo "============================================================"
echo " Build complete!"
echo " Distributable folder:  dist/MPL_Hub/"
echo " Share the entire MPL_Hub folder with your users."
echo " Users launch the app by running:  ./MPL_Hub"
echo "============================================================"
echo
