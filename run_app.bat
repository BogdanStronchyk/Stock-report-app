@echo off
setlocal enabledelayedexpansion

REM ============================================================
REM Run StockReportApp.exe (portable launcher)
REM - Looks for dist\StockReportApp.exe first
REM - Then tries StockReportApp.exe in current folder
REM - Provides helpful error messages
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

REM If you DID NOT bundle the checklist into the EXE, it must exist next to it.
REM With the current build_exe.bat, the checklist IS bundled, so this is just a warning.
if not exist "Fundamental_Checklist_v2_with_sector_adjustments.xlsx" (
  echo [INFO] Checklist file not found next to launcher:
  echo        Fundamental_Checklist_v2_with_sector_adjustments.xlsx
  echo        (This is OK if you built the EXE with --add-data, which bundles it.)
  echo.
)

REM Launch the app
"%EXE_PATH%"

REM If the app exits immediately, keep the window open so you can read messages
echo.
echo === App finished ===
pause
endlocal
