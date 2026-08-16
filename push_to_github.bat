@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ==================================================================
echo   Push project to GitHub
echo ==================================================================
echo.

REM Remember the repository URL so you only type it once
set REPO_URL=
if exist "repo_url.txt" set /p REPO_URL=<repo_url.txt

if "%REPO_URL%"=="" (
  echo Open your repository page on GitHub and copy the address bar URL.
  echo Example:  https://github.com/YOUR-USERNAME/ohsp-data-ex
  echo.
  set /p REPO_URL=Paste your repository URL here:
  echo.
)

if "%REPO_URL%"=="" (
  echo [ERROR] No repository URL given.
  pause & exit /b 1
)

echo Repository: %REPO_URL%
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

echo [2/6] Checking the repository exists and you can access it
git ls-remote "%REPO_URL%" >nul 2>&1
if errorlevel 1 (
  echo.
  echo [ERROR] Cannot reach that repository. Common causes:
  echo   1. The username in the URL is wrong - check the address bar on GitHub
  echo   2. The repository has not been created yet on GitHub
  echo      Create it at https://github.com/new  name: ohsp-data-ex
  echo      Do NOT tick "Add a README file"
  echo   3. You are signed in to a different GitHub account
  echo.
  echo URL tried: %REPO_URL%
  del /q repo_url.txt >nul 2>&1
  echo The saved URL has been cleared. Run this file again to enter a new one.
  pause & exit /b 1
)
echo       OK
> repo_url.txt echo %REPO_URL%
git remote remove origin >nul 2>&1
git remote add origin "%REPO_URL%" || goto :error

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
echo Next step - to publish a new version for the field offices, run:
echo       release.bat
echo It reads VERSION from src\build_dashboard.py, creates the tag,
echo and GitHub builds the .exe and attaches it to the Releases page.
echo.
pause
exit /b 0

:error
echo.
echo [ERROR] Something went wrong. Please read the message above.
pause
exit /b 1
