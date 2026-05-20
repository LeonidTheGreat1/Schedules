@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Проверка установки

echo ========================================
echo   Проверка установки
echo ========================================
echo.

set OK=1

echo [1] Python:
where py >nul 2>&1
if %errorlevel%==0 (
    py --version
) else (
    python --version 2>nul
    if errorlevel 1 (
        echo    НЕ НАЙДЕН - запустите installers\1_Установить_Python.bat
        set OK=0
    )
)
echo.

echo [2] Библиотеки pandas и openpyxl:
where py >nul 2>&1
if %errorlevel%==0 (
    py -c "import pandas; import openpyxl; print('    OK:', pandas.__version__)" 2>nul
) else (
    python -c "import pandas; import openpyxl; print('    OK:', pandas.__version__)" 2>nul
)
if errorlevel 1 (
    echo    НЕ УСТАНОВЛЕНЫ - запустите 2_Установить_программу.bat
    set OK=0
)
echo.

echo [3] Файлы программы:
if exist gui.py (echo    gui.py - OK) else (echo    gui.py - ОТСУТСТВУЕТ & set OK=0)
if exist schedule_breaks.py (echo    schedule_breaks.py - OK) else (echo    schedule_breaks.py - ОТСУТСТВУЕТ & set OK=0)
if exist data\ (echo    папка data\ - OK) else (echo    папка data\ - создайте или запустите программу)
echo.

echo [4] Установщик Python в пакете:
if exist installers\python-*-amd64.exe (echo    установщик в installers\ - OK) else (echo    нет файла - можно скачать через 1_Установить_Python.bat)
echo.

if %OK%==1 (
    echo ========================================
    echo   ВСЁ ГОТОВО К РАБОТЕ
    echo ========================================
) else (
    echo ========================================
    echo   ТРЕБУЕТСЯ УСТАНОВКА
    echo ========================================
)
echo.
pause
