@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Сборка ZIP для отправки руководителю

echo ========================================
echo   Сборка архива для отправки
echo ========================================
echo.
echo Будет создан ZIP с программой и установщиком Python.
echo Требуется интернет для скачивания Python (~25 МБ).
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0prepare_package.ps1"
echo.
pause
