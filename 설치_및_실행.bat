@echo off
title Calendar and To do list - Setup
color 0A
cls
echo.
echo  ================================================
echo   Calendar and To do list  -  Setup
echo  ================================================
echo.

:: Log file for debugging
set "LOG=%~dp0install_log.txt"
echo Setup started > %LOG%

:: 式式 Search for Python 式式式式式式式式式式式式式式式式式式式式式式式式式式
set "PYTHON_CMD="

where python >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
    goto :check_version
)
where py >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=py"
    goto :check_version
)
goto :install_python

:: 式式 Version check (3.10+) 式式式式式式式式式式式式式式式式式式式式式式
:check_version
for /f "tokens=2 delims= " %%v in ('%PYTHON_CMD% --version 2^>^&1') do set "PY_VER=%%v"
for /f "tokens=1 delims=." %%m in ("%PY_VER%") do set "PY_MAJOR=%%m"
for /f "tokens=2 delims=." %%n in ("%PY_VER%") do set "PY_MINOR=%%n"
echo  Python %PY_VER% found
echo Python %PY_VER% found >> %LOG%
if %PY_MAJOR% LSS 3 goto :install_python
if %PY_MAJOR%==3 if %PY_MINOR% LSS 10 goto :install_python
goto :install_packages

:: 式式 Download and install Python 3.13 式式式式式式式式式式式
:install_python
echo  Python not found or version too old. Installing Python 3.13...
echo  (This may take a few minutes)
echo.
set "PY_INSTALLER=%TEMP%\python-3.13.3-setup.exe"
set "PY_URL=https://www.python.org/ftp/python/3.13.3/python-3.13.3-amd64.exe"

echo  [1/3] Downloading Python 3.13.3 (~25MB)...
echo Downloading Python... >> %LOG%
powershell -NoProfile -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%PY_URL%', '%PY_INSTALLER%')"
if %errorlevel% neq 0 (
    echo.
    echo  [ERROR] Download failed.
    echo  Please install Python 3.10+ manually: https://www.python.org/downloads/
    echo  (Check 'Add Python to PATH' during install)
    echo Download failed >> %LOG%
    pause
    exit /b 1
)

echo  [2/3] Installing Python (please wait)...
echo Installing Python... >> %LOG%
"%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=1
if %errorlevel% neq 0 (
    echo.
    echo  [ERROR] Install failed. Try running as Administrator.
    echo Install failed >> %LOG%
    pause
    exit /b 1
)
del "%PY_INSTALLER%" >nul 2>&1
echo  [3/3] Python installed!
echo Python installed >> %LOG%
echo.

:: Refresh PATH from registry
for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v PATH 2^>nul') do set "USR_PATH=%%b"
set "PATH=%PATH%;%USR_PATH%;%LOCALAPPDATA%\Programs\Python\Python313;%LOCALAPPDATA%\Programs\Python\Python313\Scripts"
set "PYTHON_CMD=python"

:: 式式 Install PySide6 式式式式式式式式式式式式式式式式式式式式式式式式式式式式式
:install_packages
echo  Checking packages...
%PYTHON_CMD% -c "import PySide6" >nul 2>&1
if %errorlevel% neq 0 (
    echo  Installing PySide6 (first time only, 2-5 min)...
    echo Installing PySide6... >> %LOG%
    %PYTHON_CMD% -m pip install PySide6 -q --disable-pip-version-check
    if %errorlevel% neq 0 (
        echo.
        echo  [ERROR] PySide6 install failed. Try running as Administrator.
        echo PySide6 failed >> %LOG%
        pause
        exit /b 1
    )
    echo  PySide6 installed!
) else (
    echo  PySide6 already installed.
)
echo PySide6 OK >> %LOG%

:: 式式 Launch widget 式式式式式式式式式式式式式式式式式式式式式式式式式式式式式式式
set "SCRIPT_DIR=%~dp0"
echo.
echo  Launching Calendar and To do list...
echo Launching... >> %LOG%

for /f "tokens=*" %%p in ('where pythonw 2^>nul') do set "PYTHONW=%%p"
if defined PYTHONW (
    start "" "%PYTHONW%" "%SCRIPT_DIR%main.py"
) else (
    start "" %PYTHON_CMD% "%SCRIPT_DIR%main.py"
)

echo  Done! Check the right side of your screen.
echo Done >> %LOG%
echo.
timeout /t 3 >nul
