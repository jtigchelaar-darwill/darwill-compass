@echo off
setlocal
cd /d "%~dp0"

where git >nul 2>nul
if errorlevel 1 (
  echo Git is not installed.
  echo Download Git for Windows, install it, and run this script again.
  pause
  exit /b 1
)

where gh >nul 2>nul
if errorlevel 1 (
  echo GitHub CLI is not installed.
  echo Install GitHub CLI, then run this script again.
  echo Command with winget:
  echo   winget install --id GitHub.cli
  pause
  exit /b 1
)

echo.
echo You will sign in to GitHub in your browser if needed.
gh auth status >nul 2>nul
if errorlevel 1 gh auth login

set /p REPO_NAME=GitHub repository name [Darwill-AI-Prospector]:
if "%REPO_NAME%"=="" set REPO_NAME=Darwill-AI-Prospector

set /p VISIBILITY=Visibility: private or public [private]:
if "%VISIBILITY%"=="" set VISIBILITY=private
if /I not "%VISIBILITY%"=="public" set VISIBILITY=private

if not exist .git (
  git init
  git branch -M main
)

git add .
git commit -m "Darwill AI Prospector v1.4"

gh repo create "%REPO_NAME%" --%VISIBILITY% --source=. --remote=origin --push

echo.
echo Repository created and pushed successfully.
pause
