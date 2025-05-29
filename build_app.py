#!/usr/bin/env bash
# build_app.sh — one-step rebuild of your PredictionApp

set -euo pipefail

# 1. cd into your project folder
cd ~/Documents/CO2\ PAPER\ FINAL/CO2\ CODE

# 2. Create (if needed) and activate the virtualenv
if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi
source .venv/bin/activate

# 3. Install/update tooling + deps
pip install --upgrade pip setuptools wheel
pip install pyinstaller
pip install -r requirements.txt

# 4. Clean out any old build artifacts
rm -rf build dist *.spec

# 5. Run PyInstaller to produce the .app bundle
pyinstaller \
  --windowed \
  --name PredictionApp \
  --hidden-import docx \
  --hidden-import openpyxl \
  --hidden-import xgboost \
  --hidden-import sklearn \
  prediction_app.py

# 6. Done!
echo
echo "✅  Build complete! Your app bundle is here:"
echo "   dist/PredictionApp.app"
