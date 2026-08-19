@echo off
set SCRIPT_DIR=%~dp0
if exist "%SCRIPT_DIR%.venv\Scripts\python.exe" (
    "%SCRIPT_DIR%.venv\Scripts\python.exe" "%SCRIPT_DIR%server.py" %*
) else (
    python "%SCRIPT_DIR%server.py" %*
)
