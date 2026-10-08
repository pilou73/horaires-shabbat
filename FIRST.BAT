@echo off
setlocal

REM Toujours se placer dans le dossier du .bat (important)
cd /d "%~dp0"

echo === Horaires Shabbat - Setup & Run ===

REM 1) Detecter Python (preferer le launcher "py")
set "PYEXE="

where py >nul 2>&1
if %errorlevel%==0 (
  set "PYEXE=py -3"
) else (
  where python >nul 2>&1
  if %errorlevel%==0 set "PYEXE=python"
)

REM 2) Installer Python si absent
if "%PYEXE%"=="" (
  echo Python not installed. Downloading...

  powershell -NoProfile -Command ^
    "try { Start-BitsTransfer -Source 'https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe' -Destination 'python-installer.exe' } catch { Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe' -OutFile 'python-installer.exe' }"

  if not exist "python-installer.exe" (
    echo ERROR: download failed.
    pause
    exit /b 1
  )

  echo Installing Python...
  start /wait "" python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
  del /q python-installer.exe

  REM Re-detecter Python (dans cette session, "py" est le plus fiable)
  where py >nul 2>&1
  if %errorlevel%==0 (
    set "PYEXE=py -3"
  ) else (
    echo Python installed but not detected yet.
    echo Please close and re-open this window, then run again.
    pause
    exit /b 1
  )
) else (
  echo Python is installed.
)

echo Using: %PYEXE%
%PYEXE% -V

REM 3) Creer venv si absent
if not exist ".venv\Scripts\python.exe" (
  echo Creating venv (.venv)...
  %PYEXE% -m venv .venv
  if not %errorlevel%==0 (
    echo ERROR: could not create venv.
    pause
    exit /b 1
  )
)

REM 4) Installer dependances via requirements.txt
if not exist "requirements.txt" (
  echo ERROR: requirements.txt not found in this folder.
  pause
  exit /b 1
)

echo Installing dependencies from requirements.txt...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if not %errorlevel%==0 (
  echo ERROR: pip install failed.
  pause
  exit /b 1
)

REM 5) Lancer le script
if not exist "horaires_shabbat_for_local.py" (
  echo ERROR: horaires_shabbat_for_local.py not found in this folder.
  pause
  exit /b 1
)

echo Running horaires_shabbat_for_local.py...
".venv\Scripts\python.exe" "horaires_shabbat_for_local.py"

pause