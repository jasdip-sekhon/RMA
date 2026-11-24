@echo off
rem Runs the 400G RMA test. Double-click this file to start it.

rem Run from this file's folder, so main.py finds the templates.
cd /d "%~dp0"

".venv\Scripts\python.exe" main.py 400G
if errorlevel 1 (
    echo.
    echo The run stopped. The message above says why, and nothing was saved.
) else (
    echo.
    echo Done. Open the file in Excel and save it once before sharing.
)
echo.
pause
