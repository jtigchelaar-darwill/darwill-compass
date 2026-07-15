@echo off
setlocal
cd /d "%~dp0"
git pull --ff-only
pause
