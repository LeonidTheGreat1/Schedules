# Скачивает официальный установщик Python для Windows (64-bit)
$ErrorActionPreference = "Stop"
$version = "3.12.8"
$url = "https://www.python.org/ftp/python/$version/python-$version-amd64.exe"
$dest = Join-Path $PSScriptRoot "python-$version-amd64.exe"

Write-Host "Скачивание Python $version ..."
Write-Host $url
Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
Write-Host "Готово: $dest"
Write-Host "Запустите 1_Установить_Python.bat"
