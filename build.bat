@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ==================================================================
echo   Build OralHealthDashboard.exe
echo ==================================================================
echo.

where python >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python is not installed on this computer.
  echo Download from https://www.python.org/downloads/ and tick "Add Python to PATH".
  pause & exit /b 1
)

echo [1/4] Installing required libraries
python -m pip install --quiet --upgrade openpyxl pyinstaller || goto :error

echo [2/4] Setting repository name for the update notice
python packaging\set_repo.py

echo [3/4] Building the .exe file
python -m PyInstaller --onefile --noconfirm --clean ^
  --name OralHealthDashboard ^
  --add-data "src/dashboard_template.html;." ^
  --add-data "src/hospitals.csv;." ^
  --hidden-import openpyxl ^
  src/build_dashboard.py || goto :error

echo [4/4] Creating the distribution zip
python packaging\make_package.py || goto :error

echo.
echo ==================================================================
echo   Done
echo ==================================================================
echo   Program : dist\OralHealthDashboard.exe
echo   Package : OralHealthDashboard-v*.zip
echo.
pause
exit /b 0

:error
echo.
echo [ERROR] Build failed. Please read the message above.
pause
exit /b 1
