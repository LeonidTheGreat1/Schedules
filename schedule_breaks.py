#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Планировщик перерывов для смен сотрудников.
Читает input.xlsx, рассчитывает расписание перерывов и сохраняет output_schedule.xlsx.
"""

from __future__ import annotations

import sys
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

# --- Конфигурация ---

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DOCS_DIR = BASE_DIR / "docs"

INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "output_schedule.xlsx"
INPUT_PATH = DATA_DIR / INPUT_FILE
OUTPUT_PATH = DATA_DIR / OUTPUT_FILE

LUNCH_MINUTES = 30          # длительность обеда (мин)
SHORT_MINUTES = 15          # длительность короткого перерыва (мин)
MIN_GAP_MINUTES = 90        # минимум 1.5 ч между перерывами
MAX_GAP_MINUTES = 150       # максимум 2.5 ч между перерывами
FORBIDDEN_EDGE_MINUTES = 60 # первый и последний час смены без перерывов
# Первый перерыв должен начаться не позже чем через 2 ч от начала смены
FIRST_BREAK_MAX_AFTER_START_MINUTES = 120
# Шаг смещения обеда для разнесения графиков сотрудников с одной сменой
STAGGER_STEP_MINUTES = 15

# Допустимые длительности смен (часы)
SHIFT_9H_MIN = 8.5
SHIFT_9H_MAX = 9.5
SHIFT_12H_MIN = 11.5
SHIFT_12H_MAX = 12.5

# Состав перерывов (внутренние коды типов)
BREAK_PLAN_9H = ["Short", "Lunch", "Short"]           # 9 ч: 60 мин перерывов
# 12 ч: 75 мин — сначала 2 коротких перерыва, затем обед, затем 1 короткий
BREAK_PLAN_12H = ["Short", "Short", "Lunch", "Short"]

DURATIONS = {"Lunch": LUNCH_MINUTES, "Short": SHORT_MINUTES}

# Подписи типов перерывов в выходном Excel
TYPE_LABELS_RU = {"Lunch": "Обед", "Short": "Короткий"}

# Коды статуса в колонке Status
STATUS_OK = "ОК"
ERR_UNSUPPORTED_SHIFT = "Ошибка: неподдерживаемая длительность смены"
ERR_WINDOW_TOO_SMALL = "Ошибка: окно для перерывов слишком мало"
ERR_CANNOT_FIT = "Ошибка: не удалось разместить перерывы в окне"
ERR_GAP_VIOLATED = "Ошибка: нарушены интервалы между перерывами"
ERR_FIRST_BREAK_LATE = "Ошибка: первый перерыв слишком поздно"
ERR_INVALID_TIME = "Ошибка: неверный формат времени"
ERR_SLOT_LIMIT = "Ошибка: превышен лимит слота"

# Ограничение одновременных перерывов (см. блок со строки 408)
SLOT_MINUTES = 15           # длина временного слота для лимита одновременных перерывов
MAX_BREAKS_PER_SLOT = 7    # макс. сотрудников на перерыве в обычном слоте
# Первый доступный слот после 1-го часа работы (начало окна w_start) — строже
MAX_BREAKS_FIRST_WINDOW_SLOT = 2


# --- Вспомогательные функции времени ---


def parse_hhmm(value: Any) -> datetime | None:
    """Преобразует значение ячейки в datetime (сегодня + время)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.replace(year=1900, month=1, day=1)
    if isinstance(value, time):
        return datetime(1900, 1, 1, value.hour, value.minute, value.second)

    text = str(value).strip()
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            t = datetime.strptime(text, fmt)
            return datetime(1900, 1, 1, t.hour, t.minute, t.second)
        except ValueError:
            continue
    return None


def format_hhmm(dt: datetime | None) -> str:
    """Форматирует datetime в строку ЧЧ:ММ."""
    if dt is None:
        return ""
    return dt.strftime("%H:%M")


def to_minutes(dt: datetime) -> int:
    """Минуты от полуночи."""
    return dt.hour * 60 + dt.minute


def from_minutes(total: int) -> datetime:
    """Объект datetime из минут от полуночи (базовая дата 1900-01-01)."""
    total = total % (24 * 60)
    return datetime(1900, 1, 1) + timedelta(minutes=total)


def shift_duration_hours(start: datetime, end: datetime) -> float:
    """Длительность смены в часах (с учётом перехода через полночь)."""
    delta = end - start
    if delta.total_seconds() <= 0:
        delta += timedelta(days=1)
    return delta.total_seconds() / 3600.0


def duration_label(minutes: int) -> str:
    """Человекочитаемая длительность для колонки Notes."""
    if minutes >= 60:
        h, m = divmod(minutes, 60)
        return f"{h}ч{m:02d}м" if m else f"{h}ч"
    return f"{minutes}м"


def gaps_label(gaps: list[int]) -> str:
    """Строка вида «Промежутки: 1ч50м, 2ч10м» для колонки Notes."""
    parts = []
    for g in gaps:
        h, m = divmod(g, 60)
        if h and m:
            parts.append(f"{h}ч{m:02d}м")
        elif h:
            parts.append(f"{h}ч")
        else:
            parts.append(f"{m}м")
    return "Промежутки: " + ", ".join(parts)


# --- Ядро планирования ---


def _break_plan(duration_h: float) -> tuple[list[str] | None, str | None]:
    """Возвращает список типов перерывов или None и текст ошибки."""
    if SHIFT_9H_MIN <= duration_h < SHIFT_9H_MAX:
        return list(BREAK_PLAN_9H), None
    if SHIFT_12H_MIN <= duration_h < SHIFT_12H_MAX:
        return list(BREAK_PLAN_12H), None
    return None, ERR_UNSUPPORTED_SHIFT


def _min_window_needed(break_types: list[str]) -> int:
    """Минимальная длина окна (мин): сумма перерывов + (n-1)*MIN_GAP."""
    break_sum = sum(DURATIONS[t] for t in break_types)
    gaps = max(0, len(break_types) - 1) * MIN_GAP_MINUTES
    return break_sum + gaps


def _first_break_start_min(breaks: list[dict[str, Any]]) -> int:
    """Минута начала самого раннего перерыва."""
    return min(b["start_min"] for b in breaks)


def _first_break_within_two_hours(
    breaks: list[dict[str, Any]],
    shift_start_min: int,
) -> bool:
    """Первый перерыв начинается не позже чем через 2 ч от начала смены."""
    return _first_break_start_min(breaks) <= shift_start_min + FIRST_BREAK_MAX_AFTER_START_MINUTES


def _place_breaks(
    break_types: list[str],
    w_start: int,
    w_end: int,
    gaps: list[int],
    lunch_start_override: int | None = None,
) -> list[dict[str, Any]] | None:
    """
    Размещает перерывы в хронологическом порядке с заданными промежутками (мин).
    gaps[i] — интервал между концом i-го перерыва и началом (i+1)-го.
    lunch_start_override — явная минута начала обеда (для разнесения графиков).
    """
    lunch_idx = break_types.index("Lunch")
    lunch_dur = DURATIONS["Lunch"]

    if lunch_start_override is not None:
        lunch_start = max(w_start, min(lunch_start_override, w_end - lunch_dur))
    else:
        window_len = w_end - w_start
        lunch_start = w_start + (window_len - lunch_dur) // 2
        lunch_start = max(w_start, min(lunch_start, w_end - lunch_dur))
    lunch_end = lunch_start + lunch_dur

    segments: list[tuple[int, int, str]] = []

    before_types_rev = list(reversed(break_types[:lunch_idx]))
    before_gaps_rev = list(reversed(gaps[:lunch_idx])) if lunch_idx > 0 else []

    anchor_start = lunch_start
    for i, btype in enumerate(before_types_rev):
        gap = before_gaps_rev[i] if i < len(before_gaps_rev) else MIN_GAP_MINUTES
        dur = DURATIONS[btype]
        b_end = anchor_start - gap
        b_start = b_end - dur
        if b_start < w_start:
            return None
        segments.append((b_start, b_end, btype))
        anchor_start = b_start

    segments.append((lunch_start, lunch_end, "Lunch"))

    after_types = break_types[lunch_idx + 1 :]
    after_gaps = gaps[lunch_idx:] if lunch_idx < len(gaps) else []

    anchor_end = lunch_end
    for i, btype in enumerate(after_types):
        gap = after_gaps[i] if i < len(after_gaps) else MIN_GAP_MINUTES
        dur = DURATIONS[btype]
        b_start = anchor_end + gap
        b_end = b_start + dur
        if b_end > w_end:
            return None
        segments.append((b_start, b_end, btype))
        anchor_end = b_end

    segments.sort(key=lambda x: x[0])

    return [
        {
            "start": from_minutes(s),
            "end": from_minutes(e),
            "type": t,
            "start_min": s,
            "end_min": e,
        }
        for s, e, t in segments
    ]


def _maximize_gaps(
    break_types: list[str],
    w_start: int,
    w_end: int,
    lunch_start: int,
) -> list[int] | None:
    """
    Начинает с минимальных промежутков и увеличивает каждый до 2.5 ч,
    пока раскладка помещается в окно при фиксированном времени обеда.
    """
    n = len(break_types) - 1
    if n == 0:
        return []

    gaps = [MIN_GAP_MINUTES] * n
    if _place_breaks(break_types, w_start, w_end, gaps, lunch_start) is None:
        return None

    changed = True
    while changed:
        changed = False
        for i in range(n):
            if gaps[i] >= MAX_GAP_MINUTES:
                continue
            lo, hi = gaps[i], MAX_GAP_MINUTES
            while lo < hi:
                mid = (lo + hi + 1) // 2
                trial = gaps[:]
                trial[i] = mid
                if _place_breaks(break_types, w_start, w_end, trial, lunch_start) is not None:
                    lo = mid
                else:
                    hi = mid - 1
            if lo > gaps[i]:
                gaps[i] = lo
                changed = True
    return gaps


def _try_schedule(
    break_types: list[str],
    w_start: int,
    w_end: int,
    lunch_start: int,
) -> tuple[list[dict] | None, list[int] | None, str | None]:
    """Пробует построить расписание с обедом в заданную минуту; при неудаче — код ошибки."""
    n_gaps = len(break_types) - 1
    min_gaps = [MIN_GAP_MINUTES] * n_gaps
    min_needed = sum(DURATIONS[t] for t in break_types) + sum(min_gaps)
    window_len = w_end - w_start

    if window_len < min_needed:
        return None, None, ERR_WINDOW_TOO_SMALL

    final_gaps = _maximize_gaps(break_types, w_start, w_end, lunch_start)
    if final_gaps is None:
        return None, None, ERR_CANNOT_FIT

    placed = _place_breaks(break_types, w_start, w_end, final_gaps, lunch_start)
    if placed is None:
        return None, None, ERR_CANNOT_FIT

    for i in range(len(placed) - 1):
        gap = placed[i + 1]["start_min"] - placed[i]["end_min"]
        if gap < MIN_GAP_MINUTES or gap > MAX_GAP_MINUTES:
            return None, None, ERR_GAP_VIOLATED

    return placed, final_gaps, None


def _lunch_positions(w_start: int, w_end: int) -> list[int]:
    """Допустимые минуты начала обеда с шагом STAGGER_STEP_MINUTES."""
    lunch_dur = DURATIONS["Lunch"]
    last = w_end - lunch_dur
    if last < w_start:
        return []
    return list(range(w_start, last + 1, STAGGER_STEP_MINUTES))


def calculate_breaks(
    shift_start: datetime,
    shift_end: datetime,
    stagger_index: int = 0,
) -> dict[str, Any]:
    """
    Рассчитывает расписание перерывов для одной смены.

    Возвращает словарь: status, notes, duration_h, breaks (до 4 перерывов).
    """
    duration_h = shift_duration_hours(shift_start, shift_end)
    result: dict[str, Any] = {
        "status": STATUS_OK,
        "notes": "",
        "duration_h": round(duration_h, 2),
        "breaks": [],
    }

    plan, err = _break_plan(duration_h)
    if err:
        result["status"] = err
        result["notes"] = "Длительность смены вне поддерживаемых диапазонов (9 ч / 12 ч)."
        return result

    start_min = to_minutes(shift_start)
    end_min = to_minutes(shift_end)
    if end_min <= start_min:
        end_min += 24 * 60

    w_start = start_min + FORBIDDEN_EDGE_MINUTES
    w_end = end_min - FORBIDDEN_EDGE_MINUTES

    if w_end <= w_start:
        result["status"] = ERR_WINDOW_TOO_SMALL
        result["notes"] = "Запрещённые зоны съедают всё окно для перерывов."
        return result

    if (w_end - w_start) < _min_window_needed(plan):
        result["status"] = ERR_WINDOW_TOO_SMALL
        result["notes"] = (
            f"Окно {duration_label(w_end - w_start)} меньше минимума "
            f"{duration_label(_min_window_needed(plan))}."
        )
        return result

    first_break_deadline = start_min + FIRST_BREAK_MAX_AFTER_START_MINUTES
    valid_options: list[tuple[int, list[dict[str, Any]], list[int]]] = []
    last_err: str | None = None

    for lunch_start in _lunch_positions(w_start, w_end):
        trial_placed, trial_gaps, err = _try_schedule(plan, w_start, w_end, lunch_start)
        if err:
            last_err = err
            continue
        if not _first_break_within_two_hours(trial_placed, start_min):
            last_err = ERR_FIRST_BREAK_LATE
            continue
        valid_options.append((lunch_start, trial_placed, trial_gaps))

    if not valid_options:
        placed = None
    else:
        # Предпочитаем варианты без перерыва в первом слоте после 1-го часа работы
        without_first_slot = [
            opt
            for opt in valid_options
            if not _breaks_overlap_first_window_slot(opt[1], w_start, SLOT_MINUTES)
        ]
        pool = without_first_slot if without_first_slot else valid_options
        pick = stagger_index % len(pool)
        _, placed, gaps = pool[pick]

    if placed is None:
        result["status"] = last_err or ERR_CANNOT_FIT
        if last_err == ERR_FIRST_BREAK_LATE:
            result["notes"] = (
                f"Первый перерыв должен начаться до "
                f"{format_hhmm(from_minutes(first_break_deadline))} "
                f"(не позже 2 ч от начала смены)."
            )
        else:
            result["notes"] = "Не удалось разместить перерывы с интервалами 1.5–2.5 ч."
        return result

    result["breaks"] = placed
    note_parts = [gaps_label(gaps) if gaps else "Без промежутков между перерывами"]
    note_parts.append(
        f"Первый перерыв до {format_hhmm(from_minutes(first_break_deadline))}; "
        f"смещение графика #{stagger_index}"
    )
    result["notes"] = "; ".join(note_parts)
    result["w_start_min"] = w_start
    result["w_end_min"] = w_end
    result["shift_start_min"] = start_min
    result["break_plan"] = plan
    return result


# ===========================================================================
# СТРОКА 408 — НАЧАЛО: ОГРАНИЧЕНИЕ ПЕРЕРЫВОВ В ВРЕМЕННЫЕ СЛОТЫ
# Обычные слоты: MAX_BREAKS_PER_SLOT; первый слот после 1-го часа работы (w_start):
# MAX_BREAKS_FIRST_WINDOW_SLOT — меньше одновременных перерывов.
# ===========================================================================


def _first_window_slot_indices(w_start: int, slot_minutes: int) -> set[int]:
    """Индексы первого 15-минутного слота, доступного после 1-го часа смены."""
    return set(_slots_covered_by_break(w_start, w_start + slot_minutes, slot_minutes))


def _slot_max_employees(
    slot_idx: int,
    w_start: int,
    max_per_slot: int,
    max_first_window_slot: int,
    slot_minutes: int,
) -> int:
    """Лимит одновременных перерывов для данного слота (первый слот окна — строже)."""
    if slot_idx in _first_window_slot_indices(w_start, slot_minutes):
        return max_first_window_slot
    return max_per_slot


def _breaks_overlap_first_window_slot(
    breaks: list[dict[str, Any]],
    w_start: int,
    slot_minutes: int,
) -> bool:
    """True, если хотя бы один перерыв попадает в первый слот окна (после 1-го часа)."""
    first_slots = _first_window_slot_indices(w_start, slot_minutes)
    for br in breaks:
        for slot_idx in _slots_covered_by_break(br["start_min"], br["end_min"], slot_minutes):
            if slot_idx in first_slots:
                return True
    return False


def _minute_to_slot(minute: int, slot_minutes: int) -> int:
    """Индекс временного слота для абсолютной минуты (с учётом смен через полночь)."""
    return minute // slot_minutes


def _slots_covered_by_break(start_min: int, end_min: int, slot_minutes: int) -> range:
    """Диапазон индексов слотов, пересекаемых перерывом [начало, конец)."""
    first = _minute_to_slot(start_min, slot_minutes)
    last = _minute_to_slot(max(start_min, end_min - 1), slot_minutes)
    return range(first, last + 1)


def _count_employees_in_slot(
    occupancy: dict[int, set[str]],
    slot_idx: int,
) -> int:
    """Сколько сотрудников уже занимают слот."""
    return len(occupancy.get(slot_idx, ()))


def _breaks_violate_slot_limit(
    breaks: list[dict[str, Any]],
    occupancy: dict[int, set[str]],
    employee_id: str,
    max_per_slot: int,
    slot_minutes: int,
    w_start: int | None = None,
    max_first_window_slot: int = MAX_BREAKS_FIRST_WINDOW_SLOT,
) -> bool:
    """
    Возвращает True, если после добавления перерывов сотрудника лимит слота будет превышен.
    Для первого слота после 1-го часа работы действует max_first_window_slot.
    """
    for br in breaks:
        for slot_idx in _slots_covered_by_break(br["start_min"], br["end_min"], slot_minutes):
            present = occupancy.get(slot_idx, set())
            if employee_id in present:
                continue
            limit = (
                _slot_max_employees(
                    slot_idx, w_start, max_per_slot, max_first_window_slot, slot_minutes
                )
                if w_start is not None
                else max_per_slot
            )
            if len(present) >= limit:
                return True
    return False


def _register_breaks_in_slots(
    occupancy: dict[int, set[str]],
    employee_id: str,
    breaks: list[dict[str, Any]],
    slot_minutes: int,
) -> None:
    """Регистрирует перерывы сотрудника в карте занятости слотов."""
    for br in breaks:
        for slot_idx in _slots_covered_by_break(br["start_min"], br["end_min"], slot_minutes):
            occupancy.setdefault(slot_idx, set()).add(employee_id)


def _unregister_breaks_from_slots(
    occupancy: dict[int, set[str]],
    employee_id: str,
    breaks: list[dict[str, Any]],
    slot_minutes: int,
) -> None:
    """Убирает сотрудника из слотов (при откате неудачного сдвига)."""
    for br in breaks:
        for slot_idx in _slots_covered_by_break(br["start_min"], br["end_min"], slot_minutes):
            if slot_idx in occupancy:
                occupancy[slot_idx].discard(employee_id)
                if not occupancy[slot_idx]:
                    del occupancy[slot_idx]


def _shift_break_times(
    breaks: list[dict[str, Any]],
    delta_min: int,
) -> list[dict[str, Any]]:
    """Сдвигает все перерывы на delta_min минут, возвращает новый список."""
    shifted = []
    for br in breaks:
        s = br["start_min"] + delta_min
        e = br["end_min"] + delta_min
        shifted.append(
            {
                "start": from_minutes(s),
                "end": from_minutes(e),
                "type": br["type"],
                "start_min": s,
                "end_min": e,
            }
        )
    return shifted


def _validate_shifted_breaks(
    breaks: list[dict[str, Any]],
    w_start: int,
    w_end: int,
    shift_start_min: int | None = None,
) -> bool:
    """Проверяет окно смены, интервалы 1.5–2.5 ч и правило первого перерыва."""
    if not breaks:
        return False
    ordered = sorted(breaks, key=lambda b: b["start_min"])
    for br in ordered:
        if br["start_min"] < w_start or br["end_min"] > w_end:
            return False
    for i in range(len(ordered) - 1):
        gap = ordered[i + 1]["start_min"] - ordered[i]["end_min"]
        if gap < MIN_GAP_MINUTES or gap > MAX_GAP_MINUTES:
            return False
    if shift_start_min is not None and not _first_break_within_two_hours(ordered, shift_start_min):
        return False
    return True


def _find_shift_avoiding_slots(
    calc: dict[str, Any],
    occupancy: dict[int, set[str]],
    employee_id: str,
    max_per_slot: int,
    slot_minutes: int,
    max_first_window_slot: int = MAX_BREAKS_FIRST_WINDOW_SLOT,
) -> bool:
    """Подбирает сдвиг графика, чтобы не перегружать слоты (в т.ч. первый после 1-го часа)."""
    w_start = calc["w_start_min"]
    w_end = calc["w_end_min"]
    shift_start_min = calc.get("shift_start_min")
    max_shift = w_end - w_start
    first_slot_time = format_hhmm(from_minutes(w_start))

    for step in range(1, (max_shift // slot_minutes) + 1):
        delta = step * slot_minutes
        for trial_delta in (delta, -delta):
            shifted = _shift_break_times(calc["breaks"], trial_delta)
            if not _validate_shifted_breaks(shifted, w_start, w_end, shift_start_min):
                continue
            if _breaks_violate_slot_limit(
                shifted,
                occupancy,
                employee_id,
                max_per_slot,
                slot_minutes,
                w_start=w_start,
                max_first_window_slot=max_first_window_slot,
            ):
                continue
            calc["breaks"] = shifted
            if _breaks_overlap_first_window_slot(shifted, w_start, slot_minutes):
                note = (
                    f"Сдвиг {trial_delta:+d} мин (лимит слота {slot_minutes} мин, "
                    f"макс. {max_per_slot}; первый слот с {first_slot_time}: "
                    f"макс. {max_first_window_slot})"
                )
            else:
                note = (
                    f"Сдвиг {trial_delta:+d} мин (избегание перегрузки первого слота "
                    f"с {first_slot_time}, макс. {max_first_window_slot})"
                )
            calc["notes"] = f"{calc.get('notes', '')}; {note}".strip("; ")
            return True
    return False


def apply_time_slot_limits(
    entries: list[dict[str, Any]],
    max_per_slot: int = MAX_BREAKS_PER_SLOT,
    slot_minutes: int = SLOT_MINUTES,
    max_first_window_slot: int = MAX_BREAKS_FIRST_WINDOW_SLOT,
) -> None:
    """
    Постобработка: лимит перерывов в слотах.
    Первый слот после 1-го часа работы (w_start) — не более max_first_window_slot человек.

    entries — список записей с ключами emp_id и calc (результат calculate_breaks).
    Изменяет calc на месте: status, notes, breaks.
    """
    occupancy: dict[int, set[str]] = {}

    for entry in entries:
        calc = entry["calc"]
        emp_id = str(entry["emp_id"])

        if calc.get("status") != STATUS_OK or not calc.get("breaks"):
            continue

        w_start = calc["w_start_min"]
        breaks = calc["breaks"]
        if _breaks_violate_slot_limit(
            breaks,
            occupancy,
            emp_id,
            max_per_slot,
            slot_minutes,
            w_start=w_start,
            max_first_window_slot=max_first_window_slot,
        ):
            if _find_shift_avoiding_slots(
                calc,
                occupancy,
                emp_id,
                max_per_slot,
                slot_minutes,
                max_first_window_slot=max_first_window_slot,
            ):
                _register_breaks_in_slots(occupancy, emp_id, calc["breaks"], slot_minutes)
                continue
            first_slot_time = format_hhmm(from_minutes(w_start))
            calc["status"] = ERR_SLOT_LIMIT
            calc["notes"] = (
                f"{calc.get('notes', '')}; "
                f"Превышен лимит: слот {slot_minutes} мин — {max_per_slot}, "
                f"первый слот с {first_slot_time} — {max_first_window_slot}"
            ).strip("; ")
            continue

        _register_breaks_in_slots(occupancy, emp_id, breaks, slot_minutes)


def _resolve_max_per_slot(df: pd.DataFrame) -> int:
    """Читает колонку Max_Breaks_Per_Slot из Excel или значение по умолчанию."""
    if "Max_Breaks_Per_Slot" not in df.columns:
        return MAX_BREAKS_PER_SLOT
    values = df["Max_Breaks_Per_Slot"].dropna().unique()
    if len(values) == 0:
        return MAX_BREAKS_PER_SLOT
    try:
        return max(1, int(values[0]))
    except (TypeError, ValueError):
        return MAX_BREAKS_PER_SLOT


# --- Генератор тестовых данных ---


def generate_dummy_input(path: str | Path | None = None) -> Path:
    """
    Создаёт input.xlsx с 45 сотрудниками:
    - 10 × 08:00–17:00 (9 ч)
    - 10 × 09:00–18:00 (9 ч)
    - 10 × 09:00–21:00 (12 ч: 2 коротких → обед → 1 короткий)
    -  6 × 11:00–23:00 (12 ч: 2 коротких → обед → 1 короткий)
    -  5 × 16:00–01:00 (9 ч, через полночь)
    -  4 × 21:00–09:00 (12 ч, через полночь: 2 коротких → обед → 1 короткий)
    """
    groups = [
        ("08:00", "17:00", 10),
        ("09:00", "18:00", 10),
        ("09:00", "21:00", 10),
        ("11:00", "23:00", 6),
        ("16:00", "01:00", 5),
        ("21:00", "09:00", 4),
    ]

    rows: list[dict[str, str]] = []
    emp_num = 1
    for shift_start, shift_end, count in groups:
        for _ in range(count):
            rows.append(
                {
                    "Employee_ID": f"E{emp_num:03d}",
                    "Shift_Start": shift_start,
                    "Shift_End": shift_end,
                }
            )
            emp_num += 1

    df = pd.DataFrame(rows)
    df["Max_Breaks_Per_Slot"] = MAX_BREAKS_PER_SLOT
    out = Path(path) if path is not None else INPUT_PATH
    out.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out, index=False)
    return out


# --- Формирование строки вывода ---


def _empty_break_slots() -> dict[str, str]:
    """Пустые колонки Break_1..Break_4."""
    slots = {}
    for i in range(1, 5):
        slots[f"Break_{i}_Start"] = ""
        slots[f"Break_{i}_End"] = ""
        slots[f"Break_{i}_Type"] = ""
    return slots


def row_to_output(
    employee_id: Any,
    shift_start: datetime,
    shift_end: datetime,
    calc: dict[str, Any],
) -> dict[str, Any]:
    """Собирает одну строку результата для DataFrame."""
    row: dict[str, Any] = {
        "Employee_ID": employee_id,
        "Shift_Start": format_hhmm(shift_start),
        "Shift_End": format_hhmm(shift_end),
        "Shift_Duration_Hours": calc["duration_h"],
        "Status": calc["status"],
        "Notes": calc["notes"],
    }
    row.update(_empty_break_slots())

    for i, br in enumerate(calc.get("breaks") or [], start=1):
        if i > 4:
            break
        row[f"Break_{i}_Start"] = format_hhmm(br["start"])
        row[f"Break_{i}_End"] = format_hhmm(br["end"])
        row[f"Break_{i}_Type"] = TYPE_LABELS_RU.get(br["type"], br["type"])

    return row


# Колонки выходного файла output_schedule.xlsx
OUTPUT_COLUMNS = [
    "Employee_ID",
    "Shift_Start",
    "Shift_End",
    "Shift_Duration_Hours",
    "Break_1_Start",
    "Break_1_End",
    "Break_1_Type",
    "Break_2_Start",
    "Break_2_End",
    "Break_2_Type",
    "Break_3_Start",
    "Break_3_End",
    "Break_3_Type",
    "Break_4_Start",
    "Break_4_End",
    "Break_4_Type",
    "Status",
    "Notes",
]


def process_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Обрабатывает все строки входного DataFrame."""
    required = {"Employee_ID", "Shift_Start", "Shift_End"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Во входном файле отсутствуют колонки: {sorted(missing)}")

    max_per_slot = _resolve_max_per_slot(df)
    pending: list[dict[str, Any]] = []
    ordered: list[tuple[int, dict[str, Any]]] = []
    shift_stagger: dict[tuple[str, str], int] = {}

    for idx, row in df.iterrows():
        emp_id = row["Employee_ID"]
        start = parse_hhmm(row["Shift_Start"])
        end = parse_hhmm(row["Shift_End"])

        if start is None or end is None:
            ordered.append(
                (
                    idx,
                    {
                        "Employee_ID": emp_id,
                        "Shift_Start": row.get("Shift_Start", ""),
                        "Shift_End": row.get("Shift_End", ""),
                        "Shift_Duration_Hours": "",
                        "Status": ERR_INVALID_TIME,
                        "Notes": "Не удалось разобрать Shift_Start или Shift_End (ожидается ЧЧ:ММ).",
                        **_empty_break_slots(),
                    },
                )
            )
            continue

        try:
            shift_key = (format_hhmm(start), format_hhmm(end))
            stagger_index = shift_stagger.get(shift_key, 0)
            shift_stagger[shift_key] = stagger_index + 1
            calc = calculate_breaks(start, end, stagger_index=stagger_index)
            entry = {"emp_id": emp_id, "calc": calc, "start": start, "end": end, "_idx": idx}
            pending.append(entry)
            ordered.append((idx, {"_pending": entry}))
        except Exception as exc:
            ordered.append(
                (
                    idx,
                    {
                        "Employee_ID": emp_id,
                        "Shift_Start": format_hhmm(start),
                        "Shift_End": format_hhmm(end),
                        "Shift_Duration_Hours": "",
                        "Status": f"Ошибка: {exc}",
                        "Notes": "Непредвиденная ошибка при расчёте.",
                        **_empty_break_slots(),
                    },
                )
            )

    # Постобработка лимита слотов (блок со строки 408)
    apply_time_slot_limits(
        pending,
        max_per_slot=max_per_slot,
        slot_minutes=SLOT_MINUTES,
        max_first_window_slot=MAX_BREAKS_FIRST_WINDOW_SLOT,
    )

    rows_out: list[dict[str, Any]] = []
    for idx, item in sorted(ordered, key=lambda x: x[0]):
        if "_pending" in item:
            e = item["_pending"]
            rows_out.append(row_to_output(e["emp_id"], e["start"], e["end"], e["calc"]))
        else:
            rows_out.append(item)

    return pd.DataFrame(rows_out, columns=OUTPUT_COLUMNS)


def save_excel_safe(df: pd.DataFrame, target: Path) -> Path:
    """
    Сохраняет DataFrame в Excel.
    Если файл занят (открыт в Excel), пишет в резервную копию.
    """
    target = Path(target)
    candidates = [target] + [
        target.with_name(f"{target.stem}_{i}{target.suffix}") for i in range(1, 6)
    ]

    last_error: PermissionError | None = None
    for path in candidates:
        try:
            df.to_excel(path, index=False)
            if path != target:
                print(
                    f"Внимание: {target.name} занят (закройте файл в Excel). "
                    f"Результат сохранён в: {path}"
                )
            return path
        except PermissionError as exc:
            last_error = exc

    raise PermissionError(
        f"Не удалось записать результат: файл занят — {target}. "
        f"Закройте {target.name} в Excel и запустите скрипт снова."
    ) from last_error


def run_pipeline(
    input_path: Path,
    output_path: Path,
    create_dummy_if_missing: bool = True,
) -> dict[str, Any]:
    """
    Загружает Excel, считает перерывы, сохраняет результат.
    Возвращает словарь для CLI и GUI: success, message, saved_path, ok_count, total, error_count.
    """
    input_path = Path(input_path)
    output_path = Path(output_path)

    if not input_path.exists():
        if create_dummy_if_missing:
            generate_dummy_input(input_path)
        else:
            return {
                "success": False,
                "message": f"Файл не найден: {input_path}",
                "saved_path": None,
                "ok_count": 0,
                "total": 0,
                "error_count": 0,
            }

    try:
        df_in = pd.read_excel(input_path)
    except PermissionError:
        return {
            "success": False,
            "message": f"Не удалось прочитать {input_path.name}. Закройте файл в Excel.",
            "saved_path": None,
            "ok_count": 0,
            "total": 0,
            "error_count": 0,
        }

    df_out = process_dataframe(df_in)

    try:
        saved_path = save_excel_safe(df_out, output_path)
    except PermissionError as exc:
        return {
            "success": False,
            "message": str(exc),
            "saved_path": None,
            "ok_count": int((df_out["Status"] == STATUS_OK).sum()),
            "total": len(df_out),
            "error_count": int((df_out["Status"] != STATUS_OK).sum()),
        }

    ok_count = int((df_out["Status"] == STATUS_OK).sum())
    total = len(df_out)
    return {
        "success": True,
        "message": f"Готово: {ok_count} из {total} строк со статусом ОК.",
        "saved_path": saved_path,
        "ok_count": ok_count,
        "total": total,
        "error_count": total - ok_count,
    }


def main() -> int:
    """Точка входа: загрузка Excel, расчёт, сохранение результата."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    input_path = INPUT_PATH
    output_path = OUTPUT_PATH

    memo_path = DOCS_DIR / "ПАМЯТКА_для_сотрудников.md"
    if not memo_path.exists():
        memo_path = BASE_DIR / "ПАМЯТКА_для_сотрудников.md"
    if memo_path.exists():
        print(f"Инструкция для сотрудников: {memo_path}")

    result = run_pipeline(input_path, output_path, create_dummy_if_missing=True)
    if result["saved_path"]:
        print(
            f"Результат: {result['saved_path']} "
            f"({result['ok_count']}/{result['total']} со статусом ОК)."
        )
    print(result["message"])
    return 0 if result["success"] else 1


if __name__ == "__main__":
    sys.exit(main())
