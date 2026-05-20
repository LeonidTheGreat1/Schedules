#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Графический интерфейс планировщика перерывов.
Запуск: python gui.py
"""

from __future__ import annotations

import os
import sys
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

# 🔥 КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ:
# Заставляет Python работать строго из папки, где лежит этот скрипт.
# Это полностью убирает привязку к C:\Users\leonid\... или любым другим путям.
BASE_DIR = Path(__file__).resolve().parent
os.chdir(BASE_DIR)

# 🔍 Безопасный импорт основного модуля.
# Если schedule_breaks.py отсутствует или не установлены pandas/openpyxl,
# программа покажет понятное окно ошибки вместо падения в консоль.
try:
    from schedule_breaks import (
        DATA_DIR,
        DOCS_DIR,
        INPUT_PATH,
        OUTPUT_PATH,
        generate_dummy_input,
        run_pipeline,
    )
except Exception as e:
    _err_root = tk.Tk()
    _err_root.withdraw()
    messagebox.showerror(
        "Ошибка запуска",
        f"Не удалось запустить планировщик.\n\n"
        f"Причина: {e}\n\n"
        f"Что проверить:\n"
        f"1. Файл schedule_breaks.py должен лежать в той же папке, что и gui.py\n"
        f"2. Запустите installers\\2_Установить_программу.bat (или pip install pandas openpyxl)\n"
        f"3. Закройте Excel, если он открыт"
    )
    _err_root.destroy()
    sys.exit(1)

APP_TITLE = "Планировщик перерывов v1.0"
MEMO_FILE = DOCS_DIR / "ПАМЯТКА_для_сотрудников.md"
INSTALL_FILE = DOCS_DIR / "УСТАНОВКА.md"


class BreakSchedulerApp(tk.Tk):
    """Окно программы расчёта перерывов."""

    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("720x520")
        self.minsize(640, 480)

        # Создаём папку данных, если её нет
        DATA_DIR.mkdir(parents=True, exist_ok=True)

        # Если пути оказались абсолютными и привязанными к старому ПК,
        # заменяем их на относительные от BASE_DIR
        self.input_path = tk.StringVar(value=str(INPUT_PATH) if INPUT_PATH.is_absolute() else str(BASE_DIR / INPUT_PATH))
        self.output_path = tk.StringVar(value=str(OUTPUT_PATH) if OUTPUT_PATH.is_absolute() else str(BASE_DIR / OUTPUT_PATH))
        
        self.status_text = tk.StringVar(value="Готово к расчёту")
        self._busy = False

        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        header = ttk.Label(
            self,
            text="Расчёт перерывов и обедов по сменам",
            font=("Segoe UI", 14, "bold"),
        )
        header.pack(anchor="w", **pad)

        hint = ttk.Label(
            self,
            text="Заполните input.xlsx (см. памятку), закройте Excel и нажмите «Рассчитать».",
            wraplength=680,
        )
        hint.pack(anchor="w", padx=10, pady=(0, 10))

        files_frame = ttk.LabelFrame(self, text="Файлы")
        files_frame.pack(fill="x", padx=10, pady=4)

        self._file_row(files_frame, "Входной файл (смены):", self.input_path)
        self._file_row(files_frame, "Результат:", self.output_path)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=10, pady=8)

        ttk.Button(btn_frame, text="Рассчитать", command=self._on_run).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_frame, text="Тестовый input (45 чел.)", command=self._on_dummy).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_frame, text="Памятка", command=self._open_memo).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_frame, text="Папка программы", command=self._open_folder).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_frame, text="Открыть результат", command=self._open_output).pack(
            side="left"
        )

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=10, pady=4)

        status_bar = ttk.Label(self, textvariable=self.status_text, relief="sunken")
        status_bar.pack(fill="x", padx=10, pady=4)

        log_frame = ttk.LabelFrame(self, text="Журнал")
        log_frame.pack(fill="both", expand=True, padx=10, pady=(4, 10))

        self.log = scrolledtext.ScrolledText(
            log_frame, height=12, state="disabled", font=("Consolas", 10)
        )
        self.log.pack(fill="both", expand=True, padx=6, pady=6)