@echo off
setlocal enabledelayedexpansion

REM ============================================================
REM Run StockReportApp.exe (portable launcher)
REM - Looks for dist\StockReportApp.exe first
REM - Then tries StockReportApp.exe in current folder
REM - Warns if external Checklist folder is missing (optional)
REM ============================================================

cd /d "%~dp0"

set EXE_PATH=
if exist "dist\StockReportApp.exe" set EXE_PATH=dist\StockReportApp.exe
if exist "StockReportApp.exe" set EXE_PATH=StockReportApp.exe

if "%EXE_PATH%"=="" (
  echo [ERROR] Could not find StockReportApp.exe
  echo.
  echo Looked in:
  echo   - %cd%\dist\StockReportApp.exe
  echo   - %cd%\StockReportApp.exe
  echo.
  echo Run build_exe.bat first to build it.
  echo.
  pause
  exit /b 1
)

echo.
echo === Stock Report App Launcher ===
echo Executable: %cd%\%EXE_PATH%
echo.

if not exist "Checklist\Fundamental_Checklist_v2_with_sector_adjustments.xlsx" (
  echo [INFO] External checklist not found in:
  echo        %cd%\Checklist\
  echo        That's OK if the EXE was built with --add-data (bundled checklist).
  echo.
)

"%EXE_PATH%"

echo.
echo === App finished ===
pause
endlocal
