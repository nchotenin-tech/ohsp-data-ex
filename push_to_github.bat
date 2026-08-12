@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set REPO_URL=https://github.com/nchotenin/ohsp-data-ex.git

echo ==================================================================
echo   Push project to GitHub
echo   %REPO_URL%
echo ==================================================================
echo.

where git >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Git is not installed on this computer.
  echo Download from https://git-scm.com/download/win then run this file again.
  pause & exit /b 1
)

where python >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python is not installed on this computer.
  echo Download from https://www.python.org/downloads/ and tick "Add Python to PATH".
  pause & exit /b 1
)

if not exist ".git" (
  echo [1/6] Initialising git repository
  git init -q || goto :error
  git branch -M main
) else (
  echo [1/6] Git repository already exists, skipping
)

echo [2/6] Setting remote
git remote remove origin >nul 2>&1
git remote add origin %REPO_URL% || goto :error

echo [3/6] Staging files ^(.gitignore excludes real patient data^)
git add -A || goto :error

echo [4/6] Safety check - make sure no patient data is included
python packaging\check_safe.py
if errorlevel 1 (
  git reset -q
  echo.
  echo Push cancelled. Nothing was sent to GitHub.
  pause & exit /b 1
)

echo.
echo [5/6] Creating commit
git commit -q -m "Oral health dashboard from HDC Data Exchange" 2>nul
if errorlevel 1 echo       No new changes, skipping

echo [6/6] Pushing to GitHub
echo       A GitHub sign-in window may appear. Please sign in yourself.
git push -u origin main || goto :error

echo.
echo ==================================================================
echo   Done. Your code is on GitHub.
echo ==================================================================
echo.
echo Next step - let GitHub build the .exe for you:
echo       git tag v1.0.0
echo       git push origin v1.0.0
echo Then wait a few minutes and check the Releases page of your repo.
echo.
pause
exit /b 0

:error
echo.
echo [ERROR] Something went wrong. Please read the message above.
pause
exit /b 1
