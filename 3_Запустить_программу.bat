@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Планировщик перерывов

if not exist "data" mkdir data

where py >nul 2>&1
if %errorlevel%==0 (
    start "" pyw gui.py 2>nul || py gui.py
) else (
    start "" pythonw gui.py 2>nul || python gui.py
)

if %errorlevel% neq 0 (
    echo.
    echo Не удалось запустить. Выполните:
    echo   2_Установить_программу.bat
    echo   Проверить_установку.bat
    pause
)
