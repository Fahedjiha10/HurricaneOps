#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
RUNTIME_DIR="$HOME/Library/Application Support/HurricaneOps"

mkdir -p \
  "$RUNTIME_DIR/05_SCRIPTS" \
  "$RUNTIME_DIR/01_JOBS" \
  "$HOME/Library/Mobile Documents/com~apple~CloudDocs/F&T Contruction/HurricaneOps Incoming" \
  "$HOME/Library/Mobile Documents/com~apple~CloudDocs/F&T Contruction/HurricaneOps Processed"

rsync -a "$PROJECT_ROOT/05_SCRIPTS/" "$RUNTIME_DIR/05_SCRIPTS/"
cp "$PROJECT_ROOT/requirements.txt" "$RUNTIME_DIR/requirements.txt"

if [[ ! -x "$RUNTIME_DIR/.venv/bin/python" ]]; then
  python3 -m venv "$RUNTIME_DIR/.venv"
fi
"$RUNTIME_DIR/.venv/bin/python" -m pip install -r "$RUNTIME_DIR/requirements.txt"

if ! command -v tesseract >/dev/null 2>&1; then
  echo "NOTE: Tesseract OCR is not installed. PDF table extraction still works."
  echo "Image-only OCR fallback will write audit warnings until the tesseract executable is installed."
fi

echo "HurricaneOps iCloud dashboard automation installed."
echo "Open the local dashboard to poll the queue every two minutes."
echo "Incoming queue: $HOME/Library/Mobile Documents/com~apple~CloudDocs/F&T Contruction/HurricaneOps Incoming"
echo "Processed results: $HOME/Library/Mobile Documents/com~apple~CloudDocs/F&T Contruction/HurricaneOps Processed"
