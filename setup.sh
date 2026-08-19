#!/usr/bin/env bash
set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "Setting up Outlook MCP Server on Linux..."
if [ ! -d ".venv" ]; then
    echo "Creating Python virtual environment (.venv)..."
    python3 -m venv .venv
fi

echo "Installing required dependencies..."
.venv/bin/pip install -r requirements.txt

echo ""
echo "Setup complete!"
echo "Next steps:"
echo "1. Run setup: .venv/bin/python3 server.py --setup"
echo "2. Log in:    .venv/bin/python3 server.py --login"
echo "3. Test:      .venv/bin/python3 server.py --test"
