@echo off
REM Vérifie si Python est installé
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python n'est pas installe. Telechargement...
    powershell -Command "Start-BitsTransfer -Source https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe -Destination python-installer.exe"
    echo Lancement de l'installateur Python...
    start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
    del python-installer.exe
) else (
    echo Python est deja installe.
)

REM Installe les dépendances
python -m pip install --upgrade pip
python -m pip install pillow requests zmanim hdate

REM Lance le script principal
python main.py

pause