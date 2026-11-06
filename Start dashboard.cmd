@echo off
rem Opens the ResCo deck builder dashboard (local only, http://127.0.0.1:8765) in your browser.
rem Close this window (or press Ctrl+C) to stop it.
rem Optional local.cmd (not in git, see local.cmd.example): the quarter's reserving pack folder and the Python to use.
cd /d "%~dp0"
set PYTHONPATH=%~dp0src
if exist "%~dp0local.cmd" call "%~dp0local.cmd"
if not defined VENV_PY set VENV_PY=C:\Users\ErikPillon\venvs\resco\Scripts\python.exe
if defined RESCO_PACK_DIR echo Reserving pack folder: %RESCO_PACK_DIR%
if exist "%VENV_PY%" (
  "%VENV_PY%" -m resco serve --out output
) else (
  echo Virtual environment not found at %VENV_PY%, trying "python" from PATH...
  python -m resco serve --out output
)
pause
