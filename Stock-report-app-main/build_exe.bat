@echo off
setlocal enabledelayedexpansion

REM ============================================================
REM Build EXE for Stock Report App (uses existing venv)
REM - Activates .venv or venv in current folder
REM - Installs/updates requirements
REM - Bundles Checklist\Fundamental_Checklist... into the EXE
REM ============================================================

cd /d "%~dp0"
echo.
echo === Stock Report App: Build EXE ===
echo Working dir: %cd%
echo.

REM -------- Locate existing virtual environment --------
set VENV_DIR=
if exist ".venv\Scripts\activate.bat" set VENV_DIR=.venv
if exist "venv\Scripts\activate.bat" set VENV_DIR=venv

if "%VENV_DIR%"=="" (
  echo [ERROR] No virtual environment found.
  echo        Expected: .venv\Scripts\activate.bat  OR  venv\Scripts\activate.bat
  pause
  exit /b 1
)

echo Using venv: %VENV_DIR%
call "%VENV_DIR%\Scripts\activate.bat"
if errorlevel 1 (
  echo [ERROR] Failed to activate virtual environment.
  pause
  exit /b 1
)

REM -------- Install dependencies --------
echo.
echo Installing/updating dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
if errorlevel 1 (
  echo [ERROR] Dependency install failed.
  pause
  exit /b 1
)

REM -------- Check required files --------
if not exist "main.py" (
  echo [ERROR] main.py not found in %cd%
  pause
  exit /b 1
)

if not exist "Checklist\Fundamental_Checklist_v2_with_sector_adjustments.xlsx" (
  echo [ERROR] Checklist file not found:
  echo         Checklist\Fundamental_Checklist_v2_with_sector_adjustments.xlsx
  echo Put it in the Checklist folder next to main.py and try again.
  pause
  exit /b 1
)

REM -------- Clean old build artifacts --------
echo.
echo Cleaning old builds...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "main.spec" del /q "main.spec"

REM -------- Build EXE --------
echo.
echo Building EXE with PyInstaller...

REM Use --noconsole if you want no terminal window:
REM pyinstaller --onefile --noconsole ...

pyinstaller --onefile ^
  --name StockReportApp ^
  --hidden-import=tkinter ^
  --add-data "Checklist\Fundamental_Checklist_v2_with_sector_adjustments.xlsx;Checklist" ^
  main.py

if errorlevel 1 (
  echo [ERROR] PyInstaller build failed.
  pause
  exit /b 1
)

echo.
echo ✅ Build complete!
echo EXE location: %cd%\dist\StockReportApp.exe
echo.
pause
endlocal
