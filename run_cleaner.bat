@echo off
:: C Drive Cleanup Tool Launcher
:: Requests UAC elevation then runs the Python script.

:: ---------------------------------------------------------------
:: Check if we are already running as Administrator
:: ---------------------------------------------------------------
net session >nul 2>&1
if %errorLevel% == 0 (
    goto :run
)

:: ---------------------------------------------------------------
:: Re-launch this script with elevated privileges via PowerShell
:: ---------------------------------------------------------------
echo Requesting administrator privileges...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Start-Process -FilePath '%~f0' -Verb RunAs"
exit /b

:run
:: ---------------------------------------------------------------
:: Find Python (tries 'python', then 'py' launcher, then common paths)
:: ---------------------------------------------------------------
set SCRIPT=%~dp0c_drive_cleaner.py

where python >nul 2>&1
if %errorLevel% == 0 (
    set PYTHON=python
    goto :launch
)

where py >nul 2>&1
if %errorLevel% == 0 (
    set PYTHON=py
    goto :launch
)

if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set PYTHON="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    goto :launch
)
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set PYTHON="%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    goto :launch
)
if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" (
    set PYTHON="%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    goto :launch
)
if exist "C:\Python312\python.exe" (
    set PYTHON="C:\Python312\python.exe"
    goto :launch
)
if exist "C:\Python311\python.exe" (
    set PYTHON="C:\Python311\python.exe"
    goto :launch
)

echo.
echo ERROR: Python not found.
echo Please install Python 3.10+ from https://www.python.org/downloads/
echo and ensure it is added to your PATH.
echo.
pause
exit /b 1

:launch
echo Starting C Drive Cleanup Tool...
%PYTHON% "%SCRIPT%"
if %errorLevel% neq 0 (
    echo.
    echo The script exited with an error (code %errorLevel%).
    pause
)
