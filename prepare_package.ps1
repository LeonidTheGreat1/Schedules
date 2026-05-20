# Сборка ZIP-пакета для отправки руководителю
$ErrorActionPreference = "Stop"
$version = "1.0"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$releaseName = "Планировщик_перерывов_v${version}"
$staging = Join-Path $env:TEMP $releaseName
$zipPath = Join-Path $root "${releaseName}.zip"

Write-Host "Сборка пакета: $releaseName"

if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
New-Item -ItemType Directory -Path $staging | Out-Null
New-Item -ItemType Directory -Path (Join-Path $staging "data") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $staging "docs") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $staging "installers") | Out-Null

# Основные файлы
$copyRoot = @(
    "gui.py", "schedule_breaks.py", "requirements.txt",
    "НАЧНИТЕ_ЗДЕСЬ.txt", "ВЕРСИЯ.txt",
    "2_Установить_программу.bat", "3_Запустить_программу.bat",
    "Проверить_установку.bat", "СОБРАТЬ_ПАКЕТ_ДЛЯ_ОТПРАВКИ.bat"
)
foreach ($f in $copyRoot) {
    $src = Join-Path $root $f
    if (Test-Path $src) { Copy-Item $src (Join-Path $staging $f) }
}

# Документация
$copyDocs = @(
    "ИНСТРУКЦИЯ_ДЛЯ_РУКОВОДИТЕЛЯ.md",
    "ПАМЯТКА_для_сотрудников.md",
    "УСТАНОВКА.md"
)
foreach ($f in $copyDocs) {
    $src = Join-Path $root "docs\$f"
    if (Test-Path $src) { Copy-Item $src (Join-Path $staging "docs\$f") }
}

# Установщики (скрипты)
$copyInst = @(
    "README.txt", "1_Установить_Python.bat", "download_python_installer.ps1"
)
foreach ($f in $copyInst) {
    $src = Join-Path $root "installers\$f"
    if (Test-Path $src) { Copy-Item $src (Join-Path $staging "installers\$f") }
}

# Скачать установщик Python (~25 МБ)
$pyVer = "3.12.8"
$pyExe = "python-$pyVer-amd64.exe"
$pyUrl = "https://www.python.org/ftp/python/$pyVer/$pyExe"
$pyDest = Join-Path $staging "installers\$pyExe"
if (-not (Test-Path (Join-Path $root "installers\$pyExe"))) {
    Write-Host "Скачивание $pyExe (это может занять 1-2 минуты)..."
    try {
        Invoke-WebRequest -Uri $pyUrl -OutFile $pyDest -UseBasicParsing
    } catch {
        Write-Warning "Не удалось скачать Python: $_"
        Write-Warning "Получатель сможет скачать через 1_Установить_Python.bat"
    }
} else {
    Copy-Item (Join-Path $root "installers\$pyExe") $pyDest
}

# Тестовый input
Push-Location $root
try {
    $env:PYTHONPATH = $root
    if (Get-Command py -ErrorAction SilentlyContinue) {
        py -c "import sys; sys.path.insert(0, r'$root'); from schedule_breaks import generate_dummy_input, INPUT_PATH; generate_dummy_input(INPUT_PATH)"
    } else {
        python -c "import sys; sys.path.insert(0, r'$root'); from schedule_breaks import generate_dummy_input, INPUT_PATH; generate_dummy_input(INPUT_PATH)"
    }
    if (Test-Path (Join-Path $root "data\input.xlsx")) {
        Copy-Item (Join-Path $root "data\input.xlsx") (Join-Path $staging "data\input.xlsx")
    }
} catch {
    Write-Warning "Не удалось создать input.xlsx: $_"
}
Pop-Location

# Пустой шаблон output (опционально - skip)

# ZIP
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $staging -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  ГОТОВО К ОТПРАВКЕ" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "Архив: $zipPath"
Write-Host "Размер: $([math]::Round((Get-Item $zipPath).Length / 1MB, 2)) МБ"
Write-Host ""
Write-Host "Отправьте руководителю файл ZIP."
Write-Host "Внутри: распаковать -> НАЧНИТЕ_ЗДЕСЬ.txt -> шаги 1, 2, 3"
