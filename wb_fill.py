# wb_fill.py
from __future__ import annotations

import re
import json
import time
import random
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Callable, Set, Tuple

from openpyxl import load_workbook


# =========================
# Utils
# =========================
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").lower().strip())


def _cap(s: str) -> str:
    return s[:1].upper() + s[1:] if s else s


def _jaccard(a: str, b: str) -> float:
    wa = set(re.findall(r"[a-zа-я0-9]+", a.lower()))
    wb = set(re.findall(r"[a-zа-я0-9]+", b.lower()))
    return len(wa & wb) / max(len(wa | wb), 1)


def _first_sentences(text: str, n: int) -> str:
    parts = re.split(r"(?<=[.!?])\s+", text)
    return " ".join(parts[:n])


def _join_ru(items: List[str]) -> str:
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} и {items[1]}"
    return ", ".join(items[:-1]) + " и " + items[-1]


# =========================
# Params
# =========================
@dataclass
class FillParams:
    xlsx_path: str
    output_dir: str

    brand_lat: str
    brand_ru: str
    shape: str
    lenses: str
    collection: str

    holidays: str
    holiday_pos: str

    seo_level: str
    style: str
    wb_safe: bool
    wb_strict: bool

    brand_title_ratio: str
    rows_to_fill: int
    skip_first_rows: int
    batch_count: int

    progress_callback: Optional[Callable[[int], None]] = None


# =========================
# Content pools
# =========================
SLOGANS = [
    "Красивые", "Крутые", "Стильные", "Модные", "Молодёжные", "Трендовые",
    "Эффектные", "Лёгкие", "Актуальные", "Дизайнерские", "Удобные",
    "Свежие", "Яркие", "Универсальные", "Аккуратные", "Летние",
]

PRODUCTS = ["солнцезащитные очки", "солнечные очки"]

SCENES = [
    "город", "прогулки", "отпуск", "пляж", "поездки",
    "путешествия", "дорога", "выходные", "летние дни",
]

SEO_BASE = [
    "очки солнцезащитные",
    "солнечные очки",
    "модные очки",
    "очки для отпуска",
    "очки для города",
    "очки женские",
    "очки мужские",
    "очки унисекс",
    "брендовые очки",
    "очки uv400",
    "инста очки",
    "очки из tiktok",
]

RISK_WORDS = [
    r"\b100%\b", r"\bлучшие\b", r"\bгарантия\b", r"\bидеальные\b",
]


# =========================
# SEO pack
# =========================
def _seo_pack(rnd: random.Random, level: str, strict: bool) -> List[str]:
    keys = SEO_BASE[:]

    if strict:
        keys = [k for k in keys if "инста" not in k and "tiktok" not in k]

    rnd.shuffle(keys)

    if level == "low":
        return keys[:3]
    if level == "high":
        return keys[:6]
    return keys[:4]


# =========================
# Title
# =========================
def make_title(rnd, brand_ru, shape, lenses, ratio, used: Set[str]) -> str:
    for _ in range(20):
        slogan = rnd.choice(SLOGANS)
        prod = rnd.choice(PRODUCTS)

        include_brand = (
            ratio == "100/0"
            or (ratio == "50/50" and rnd.random() < 0.5)
        )

        parts = [slogan, prod]
        if include_brand and brand_ru:
            parts.append(brand_ru)

        title = " ".join(parts).strip()
        if len(title) > 60:
            title = title[:60].rstrip()

        sig = _norm(title)
        if sig not in used:
            used.add(sig)
            return title

    return title


# =========================
# Description
# =========================
def make_description(
    rnd, params: FillParams,
    used_descs: List[str],
    used_opens: Set[str],
) -> str:

    for _ in range(25):
        first = rnd.choice([
            "Очки отлично дополняют образ и подходят на каждый день.",
            "Эти очки легко вписываются в повседневный стиль.",
            "Аксессуар, который делает образ собранным и аккуратным.",
            "Очки смотрятся современно и подходят под разные образы.",
        ])

        blocks = [first]

        if params.brand_lat:
            blocks.append(
                f"Модель {params.brand_lat} выглядит аккуратно и подходит для города и отдыха."
            )

        if params.shape:
            blocks.append(
                f"Форма оправы подчёркивает черты лица и смотрится гармонично."
            )

        if params.lenses:
            blocks.append(
                f"Линзы обеспечивают комфорт при ярком солнце."
            )

        blocks.append(
            f"Подходят для таких ситуаций: {', '.join(rnd.sample(SCENES, 3))}."
        )

        if params.collection:
            blocks.append(
                f"Сезон {params.collection} — актуальный вариант на тёплое время года."
            )

        # Holidays
        if params.holidays:
            hs = [h.strip() for h in params.holidays.split("||") if h.strip()]
            if hs:
                blocks.append(
                    f"Часто выбирают в подарок к { _join_ru(hs) }."
                )

        # SEO — ВПЛЕТАЕМ
        keys = _seo_pack(rnd, params.seo_level, params.wb_strict)
        if keys:
            blocks.insert(
                2,
                f"Это {keys[0]}, которые удобно носить каждый день."
            )
        if len(keys) > 1:
            blocks.append(
                f"Хороший выбор, если нужны {keys[1]}."
            )

        text = " ".join(_cap(b).rstrip(".") + "." for b in blocks)

        if params.wb_safe:
            for r in RISK_WORDS:
                text = re.sub(r, "", text, flags=re.I)

        sig = _first_sentences(text, 2)
        if sig in used_opens:
            continue

        if any(_jaccard(text, d) > 0.45 for d in used_descs[-20:]):
            continue

        used_opens.add(sig)
        used_descs.append(text)
        return text

    return text


# =========================
# Excel fill
# =========================
def fill_wb_template(params: FillParams) -> Tuple[List[str], int, str]:
    rnd = random.Random(time.time())

    wb = load_workbook(params.xlsx_path)
    ws = wb.active

    header_row = 1
    name_col = desc_col = None

    for c in range(1, ws.max_column + 1):
        v = str(ws.cell(header_row, c).value or "").lower()
        if "наимен" in v:
            name_col = c
        if "описан" in v:
            desc_col = c

    if not name_col or not desc_col:
        raise ValueError("Не найдены колонки Наименование и/или Описание")

    start = header_row + 1 + params.skip_first_rows
    rows = list(range(start, start + params.rows_to_fill))

    used_titles = set()
    used_descs = []
    used_opens = set()

    out_files = []

    for idx in range(1, params.batch_count + 1):
        for r in rows:
            ws.cell(r, name_col).value = make_title(
                rnd, params.brand_ru,
                params.shape, params.lenses,
                params.brand_title_ratio,
                used_titles
            )
            ws.cell(r, desc_col).value = make_description(
                rnd, params,
                used_descs,
                used_opens
            )

        out = Path(params.output_dir) / f"result_{idx}.xlsx"
        wb.save(out)
        out_files.append(str(out))

        if params.progress_callback:
            params.progress_callback(int(idx * 100 / params.batch_count))

    return out_files, len(rows) * params.batch_count, "OK"
