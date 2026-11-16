@echo off
cd /d "%~dp0"

echo === Checking Conda ===
call conda --version
if errorlevel 1 (
    echo Conda was not found.
    pause
    exit /b 1
)

echo.
echo === Creating pySTREAM_class ===
call conda create -n pySTREAM_class python=3.12 -y
if errorlevel 1 goto failed

echo.
echo === Installing required packages ===
call conda install -n pySTREAM_class --file requirements.txt -y
if errorlevel 1 goto failed

echo.
echo === Installation complete ===
call conda env list

pause
exit /b 0

:failed
echo.
echo Installation failed. See the messages above.
pause
exit /b 1