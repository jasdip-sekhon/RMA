@echo off
rem Runs the 200G RMA test. Double-click this file to start it.

rem Run from this file's folder, so main.py finds the templates.
cd /d "%~dp0"

".venv\Scripts\python.exe" main.py 200G
if errorlevel 1 (
    echo.
    echo The run stopped. The message above says why, and nothing was saved.
) else (
    echo.
    echo Done. The result is open in Excel. Save it once before sharing.
)
echo.
pause
