@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File "%~dp0install_shortcuts.ps1"
echo Desktop and Start Menu shortcuts created.
pause
