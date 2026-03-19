@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ====================================
echo  Boss DPS Calculator EXE Build
echo ====================================
echo.

set "APP_NAME=BossDPSCalculator_v2_1.6"

set "PY_CMD="
if exist "%LocalAppData%\Programs\Python\Python313-32\python.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python313-32\python.exe"
)

if not "%PY_CMD%"=="" goto :py_ok

where py >nul 2>nul
if %ERRORLEVEL%==0 (
    set "PY_CMD=py"
) else (
    where python >nul 2>nul
    if %ERRORLEVEL%==0 (
        set "PY_CMD=python"
    )
)

if "%PY_CMD%"=="" (
    echo [ERROR] Python launcher not found.
    echo Install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

:py_ok

echo [1/4] Python command: %PY_CMD%

if not exist venv (
    echo [2/4] Creating virtual environment...
    "%PY_CMD%" -m venv venv
    if errorlevel 1 goto :fail
)

echo [3/4] Installing packages...
call venv\Scripts\activate
"%PY_CMD%" -m pip install --upgrade pip --quiet
if errorlevel 1 goto :fail
"%PY_CMD%" -m pip install customtkinter pyinstaller keyboard pyautogui pystray pillow --quiet
if errorlevel 1 goto :fail

if exist "main_image.png" (
    echo [3.5/4] Generating app.ico from main_image.png...
    "%PY_CMD%" -c "from PIL import Image; Image.open('main_image.png').save('app.ico', format='ICO', sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])"
    if errorlevel 1 goto :fail
) else if not exist "app.ico" (
    if exist "1.png" (
        echo [3.5/4] Generating app.ico from 1.png...
        "%PY_CMD%" -c "from PIL import Image; Image.open('1.png').save('app.ico', format='ICO', sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])"
        if errorlevel 1 goto :fail
    ) else (
        echo [ERROR] main_image.png, app.ico, 1.png are all missing.
        goto :fail
    )
)

set "ADD_DATA_ICON="
set "ADD_DATA_MAIN_IMAGE="
set "ADD_DATA_PNG="
if exist "app.ico" set "ADD_DATA_ICON=--add-data app.ico;."
if exist "main_image.png" set "ADD_DATA_MAIN_IMAGE=--add-data main_image.png;."
if exist "1.png" set "ADD_DATA_PNG=--add-data 1.png;."

echo [4/4] Building EXE...
"%PY_CMD%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name "%APP_NAME%" ^
    --icon "app.ico" ^
    %ADD_DATA_ICON% ^
    %ADD_DATA_MAIN_IMAGE% ^
    %ADD_DATA_PNG% ^
    --collect-all customtkinter ^
    --hidden-import keyboard ^
    --hidden-import pyautogui ^
    --hidden-import pystray ^
    --hidden-import PIL ^
    main.py
if errorlevel 1 goto :fail

echo.
echo Done.
echo Output: dist\%APP_NAME%.exe
pause
exit /b 0

:fail
echo.
echo Build failed. Check errors above.
pause
exit /b 1
