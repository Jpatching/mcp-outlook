#!/usr/bin/env bash
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
if [ -f "$SCRIPT_DIR/.venv/bin/python3" ]; then
    PYTHON_EXEC="$SCRIPT_DIR/.venv/bin/python3"
else
    PYTHON_EXEC="python3"
fi
exec "$PYTHON_EXEC" "$SCRIPT_DIR/server.py" "$@"
