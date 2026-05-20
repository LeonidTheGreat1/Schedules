@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"
title Установка компонентов программы

echo ========================================
echo   Шаг 2: Установка библиотек Excel
echo ========================================
echo.

where py >nul 2>&1
if !errorlevel!==0 (
    echo [OK] Найден Python Launcher (py)...
    py -m pip install --upgrade pip
    py -m pip install -r requirements.txt
    set PYOK=!errorlevel!
) else (
    echo [INFO] Launcher не найден, пробуем python...
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    set PYOK=!errorlevel!
)

echo.
if !PYOK! neq 0 (
    echo ========================================
    echo   ОШИБКА. Установка не завершена
    echo ========================================
    echo.
    echo 1. Закройте Excel и другие программы.
    echo 2. Запустите этот файл от имени администратора.
    echo 3. Если ошибка повторяется, выполните сначала:
    echo    installers\1_Установить_Python.bat
    echo.
    pause
    exit /b 1
)

echo ========================================
echo   Установка завершена успешно
echo ========================================
echo.
echo Запуск программы через 3 секунды...
timeout /t 3 /nobreak >nul

REM Запускаем следующий скрипт или сам GUI
if exist "3_Запустить_программу.bat" (
    start "" "3_Запустить_программу.bat"
) else (
    REM Если отдельного скрипта запуска нет, запускаем GUI напрямую
    start "" pythonw gui.py
)

exit /b 0