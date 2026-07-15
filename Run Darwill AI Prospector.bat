@echo off
setlocal
cd /d "%~dp0"
if not exist .venv (
  echo The application has not been installed yet.
  echo Running first-time setup...
  call install_and_run.bat
  exit /b
)
call .venv\Scripts\activate.bat
python -m darwill_ai_prospector
if errorlevel 1 pause
