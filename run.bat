@echo off
REM Vérifie si Python est installé
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not installed. Downloading...
    powershell -Command "Start-BitsTransfer -Source https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe -Destination python-installer.exe"
    echo Run Installing Python...
    start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
    del python-installer.exe
) else (
    echo Python is installed.
)

REM Installe packages
python -m pip install --upgrade pip
python -m pip install pillow requests zmanim hdate

REM Execute the script
python zmanim.py

pause