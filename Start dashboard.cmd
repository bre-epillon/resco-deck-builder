@echo off
rem Opens the ResCo deck builder dashboard (local only, http://127.0.0.1:8765) in your browser.
rem Close this window (or press Ctrl+C) to stop it.
cd /d "%~dp0"
set PYTHONPATH=%~dp0src
set VENV_PY=C:\Users\ErikPillon\venvs\resco\Scripts\python.exe
if exist "%VENV_PY%" (
  "%VENV_PY%" -m resco serve --out output
) else (
  echo Virtual environment not found at %VENV_PY%, trying "python" from PATH...
  python -m resco serve --out output
)
pause
