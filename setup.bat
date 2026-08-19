@echo off
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

echo Setting up Outlook MCP Server on Windows...
if not exist ".venv" (
    echo Creating Python virtual environment (.venv)...
    python -m venv .venv
)

echo Installing required dependencies...
.venv\Scripts\pip.exe install -r requirements.txt

echo.
echo Setup complete!
echo Next steps:
echo 1. Run setup: .venv\Scripts\python.exe server.py --setup
echo 2. Log in:    .venv\Scripts\python.exe server.py --login
echo 3. Test:      .venv\Scripts\python.exe server.py --test
