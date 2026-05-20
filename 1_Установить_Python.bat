@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Установка Python

echo ========================================
echo   Шаг 1: Установка Python
echo ========================================
echo.
echo ВАЖНО в окне установщика:
echo   [x] Add python.exe to PATH  (внизу окна)
echo   Затем: Install Now
echo.

set "INSTALLER="
for %%F in (python-*-amd64.exe) do set "INSTALLER=%%F"

if not defined INSTALLER (
    echo Установщик Python не найден в папке installers\
    echo.
    choice /C YN /M "Скачать установщик с python.org сейчас"
    if errorlevel 2 goto :manual
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0download_python_installer.ps1"
    if errorlevel 1 goto :error
    for %%F in (python-*-amd64.exe) do set "INSTALLER=%%F"
)

if not defined INSTALLER (
    :manual
    echo.
    echo Скачайте Python вручную:
    echo https://www.python.org/downloads/
    echo Сохраните файл python-*-amd64.exe в папку:
    echo %~dp0
    start https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Запуск: %INSTALLER%
start /wait "" "%INSTALLER%" InstallAllUsers=0 PrependPath=1 Include_test=0 Shortcuts=0

echo.
echo Python установлен. Следующий шаг: 2_Установить_программу.bat
pause
exit /b 0

:error
echo Ошибка скачивания. Проверьте интернет.
pause
exit /b 1
