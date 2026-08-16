@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ==================================================================
echo   Publish a new release - GitHub will build the .exe for you
echo ==================================================================
echo.

where git >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Git is not installed. https://git-scm.com/download/win
  pause & exit /b 1
)

REM Read VERSION from the source file so the tag always matches the program
for /f "usebackq tokens=2 delims== " %%V in (`findstr /b /c:"VERSION = " src\build_dashboard.py`) do set RAW=%%V
set VER=%RAW:"=%
if "%VER%"=="" (
  echo [ERROR] Could not read VERSION from src\build_dashboard.py
  pause & exit /b 1
)
set TAG=v%VER%

echo Version in source : %VER%
echo Tag to publish    : %TAG%
echo.

git rev-parse "%TAG%" >nul 2>&1
if not errorlevel 1 (
  echo [ERROR] Tag %TAG% already exists.
  echo Change VERSION in src\build_dashboard.py to a higher number, then run again.
  pause & exit /b 1
)

echo [1/4] Committing any pending changes
git add -A
python packaging\check_safe.py
if errorlevel 1 (
  git reset -q
  echo Cancelled - nothing was sent.
  pause & exit /b 1
)
git commit -q -m "Release %TAG%" 2>nul
if errorlevel 1 echo       Nothing new to commit

echo [2/4] Pushing code
git push || goto :error

echo [3/4] Creating tag %TAG%
git tag %TAG% || goto :error

echo [4/4] Pushing tag - this starts the build on GitHub
git push origin %TAG% || goto :error

echo.
echo ==================================================================
echo   Done. GitHub is building the .exe now.
echo ==================================================================
echo.
echo Watch progress here:
for /f "delims=" %%U in ('git config --get remote.origin.url') do set URL=%%U
set PAGE=%URL:.git=%
echo    %PAGE%/actions
echo.
echo When the tick turns green, download the zip from:
echo    %PAGE%/releases
echo.
echo It usually takes about 3 to 5 minutes.
pause
exit /b 0

:error
echo.
echo [ERROR] Something went wrong. Please read the message above.
pause
exit /b 1
