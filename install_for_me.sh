#!/usr/bin/env bash
set -euo pipefail

TOOL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
USER_BIN="$HOME/bin"

mkdir -p "$USER_BIN"
ln -sf "$TOOL_DIR/latex2reviewdocx" "$USER_BIN/latex2reviewdocx"

echo "latex2reviewdocx is ready."
echo "Open a new terminal or run: hash -r"
echo "Then run: latex2reviewdocx --help"
