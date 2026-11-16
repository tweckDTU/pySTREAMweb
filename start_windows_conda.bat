@echo off
cd /d "%~dp0"

rem Check that the pySTREAM Conda environment exists
call conda run -n pySTREAM_class python --version >nul 2>nul
if errorlevel 1 (
    echo pySTREAM is not installed yet. Double-click install_windows.bat first.
    pause
    exit /b 1
)

echo Starting pySTREAM. The window can take up to a minute to appear...
call conda run -n pySTREAM_class python pystream.py
rem Keep this window open if pySTREAM stopped with an error, so the message can be read

if errorlevel 1 (
    echo.
    echo pySTREAM could not be started.
    pause
)
