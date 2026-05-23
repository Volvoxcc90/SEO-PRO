# wb_fill.py
from __future__ import annotations

import re
import json
import time
import random
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Callable, Set, Tuple, Dict

from openpyxl import load_workbook


# =========================
# Utils
# =========================
def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("&", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _cap(s: str) -> str:
    s = (s or "").strip()
    return s[:1].upper() + s[1:] if s else s


def _jaccard(a: str, b: str) -> float:
    wa = set(re.findall(r"[a-zа-я0-9]+", (a or "").lower()))
    wb = set(re.findall(r"[a-zа-я0-9]+", (b or "").lower()))
    return len(wa & wb) / max(len(wa | wb), 1)


def _first_sentences(text: str, n: int) -> str:
    parts = re.split(r"(?<=[.!?])\s+", (text or "").strip())
    parts = [p.strip() for p in parts if p.strip()]
    return " ".join(parts[:n]).strip()


def _join_ru(items: List[str]) -> str:
    items = [x.strip() for x in items if x and x.strip()]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} и {items[1]}"
    return ", ".join(items[:-1]) + " и " + items[-1]


def _safe_filename(name: str) -> str:
    name = (name or "").strip()
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:120] if name else "output"


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

    # WB extra fields (from Tom.xlsx)
    fill_wb_fields: bool
    overwrite_wb_fields_if_not_empty: bool  # False = fill only if empty
    kiz: bool
    adult18: bool
    gender: str
    color: str
    composition: str

    progress_callback: Optional[Callable[[int], None]] = None
    uniqueness: int = 92


# =========================
# Pools
# =========================
SLOGANS = [
    "Красивые", "Крутые", "Стильные", "Модные", "Молодёжные", "Трендовые",
    "Эффектные", "Лёгкие", "Актуальные", "Дизайнерские", "Удобные",
    "Свежие", "Яркие", "Универсальные", "Аккуратные", "Летние",
    "Выразительные", "Солидные", "Новые", "Топовые",
]
PRODUCTS = ["солнцезащитные очки", "солнечные очки"]

SCENES = [
    "город", "прогулки", "отпуск", "пляж", "поездки",
    "путешествия", "дорога", "выходные", "летние дни", "вождение",
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
    r"\b100%\b", r"\bлучшие\b", r"\bсамые лучшие\b",
    r"\bгарантированно\b", r"\bидеальные\b", r"\bабсолютно\b",
]

STOP_PHRASES_STRICT = [
    r"\bпо факту\b", r"\bпрям\b", r"\bреально\b", r"\bтоп\b",
]


# =========================
# SEO pack
# =========================
def _seo_pack(rnd: random.Random, level: str, strict: bool) -> List[str]:
    keys = SEO_BASE[:]
    if strict:
        keys = [k for k in keys if "инста" not in k and "tiktok" not in k]
    rnd.shuffle(keys)
    level = (level or "normal").strip().lower()
    if level == "low":
        return keys[:3]
    if level == "high":
        return keys[:6]
    return keys[:4]


# =========================
# Title
# =========================
def make_title(rnd: random.Random, brand_ru: str, shape: str, lenses: str, ratio: str, used: Set[str]) -> str:
    for _ in range(25):
        slogan = rnd.choice(SLOGANS)
        prod = rnd.choice(PRODUCTS)

        ratio = (ratio or "50/50").strip()
        if ratio == "100/0":
            include_brand = True
        elif ratio == "0/100":
            include_brand = False
        else:
            include_brand = rnd.random() < 0.5

        parts = [slogan, prod]
        if include_brand and brand_ru:
            parts.append(brand_ru)

        # иногда добавим линзы/форму без капса
        if lenses and rnd.random() < 0.60:
            parts.append(lenses.strip())
        if shape and rnd.random() < 0.45:
            parts.append(shape.strip().lower())

        title = " ".join([p for p in parts if p]).strip()

        # <= 60 без обрезки слов: если длинно — выкидываем хвост
        while len(title) > 60 and len(parts) > 2:
            parts.pop()
            title = " ".join([p for p in parts if p]).strip()
        if len(title) > 60:
            title = title[:60].rstrip()

        sig = _norm(title)
        if sig not in used:
            used.add(sig)
            return title

    return title


# =========================
# Description (народно + анти-монотонность)
# =========================
def _apply_safe(text: str) -> str:
    t = text
    for r in RISK_WORDS:
        t = re.sub(r, "", t, flags=re.I)
    t = re.sub(r"\s{2,}", " ", t).strip()
    return t


def _apply_strict(text: str) -> str:
    t = text
    for r in STOP_PHRASES_STRICT:
        t = re.sub(r, "", t, flags=re.I)
    t = re.sub(r"\s{2,}", " ", t).strip()
    return t


def make_description(rnd: random.Random, p: FillParams, used_descs: List[str], used_open2: Set[str], used_open3: Set[str]) -> str:
    # thresholds from uniqueness
    uni = max(0, min(100, int(p.uniqueness)))
    max_sim = 0.55 - (uni / 400.0)     # 92 => ~0.32
    max_sim = max(0.18, min(0.55, max_sim))

    first_pool = [
        "Очки отлично дополняют образ и подходят на каждый день.",
        "Эти очки легко вписываются в повседневный стиль и не перегружают лицо.",
        "Аксессуар, который делает образ собранным и аккуратным — без лишнего шума.",
        "Очки смотрятся современно и подходят под разные образы — и повседневные, и более нарядные.",
        "Лёгкая деталь, которая меняет образ: становится заметнее и аккуратнее.",
        "Очки выглядят аккуратно и “дорого”, при этом носить удобно каждый день.",
        "Хороший вариант на сезон: и стиль добавляет, и от солнца помогает.",
    ]

    for _ in range(30):
        rnd.shuffle(first_pool)
        first = first_pool[0]

        blocks: List[str] = [first]

        # бренд латиницей — естественно
        if p.brand_lat:
            blocks.append(f"Модель {p.brand_lat} смотрится уверенно и легко сочетается с одеждой на каждый день.")

        # форма — без ярлыков
        if p.shape:
            blocks.append("Оправа подчёркивает черты лица и добавляет образу выразительности — выглядит гармонично.")

        # линзы — по делу
        if p.lenses:
            lk = _norm(p.lenses)
            if "uv400" in lk:
                blocks.append("Линзы UV400 дают комфорт при ярком солнце — удобно в городе, в дороге и на отдыхе.")
            elif "поляр" in lk:
                blocks.append("Поляризационные линзы уменьшают блики — удобно за рулём и у воды.")
            elif "фотох" in lk or "хамелеон" in lk:
                blocks.append("Фотохромные линзы (хамелеон) подстраиваются под свет — комфортно, когда освещение меняется.")
            else:
                blocks.append("Линзы помогают чувствовать себя комфортно при ярком солнце.")

        # сценарии
        blocks.append(f"Подойдут для таких ситуаций: {', '.join(rnd.sample(SCENES, 4))}. Можно брать себе или на подарок.")

        # коллекция — без метки
        if p.collection and rnd.random() < 0.85:
            blocks.append(f"Сезон {p.collection}: хороший вариант, чтобы обновить аксессуары к тёплому времени года.")

        # праздники (мульти)
        if p.holidays:
            hs = [h.strip() for h in p.holidays.split("||") if h.strip()]
            if hs:
                holiday_line = f"Часто выбирают в подарок к {_join_ru(hs)} — практично и выглядит красиво."
                pos = (p.holiday_pos or "middle").lower()
                if pos == "start":
                    blocks.insert(1, holiday_line)
                elif pos == "end":
                    blocks.append(holiday_line)
                else:
                    blocks.insert(min(3, len(blocks)), holiday_line)

        # SEO — вплетаем, НЕ отдельной SEO-фразой
        keys = _seo_pack(rnd, p.seo_level, p.wb_strict)
        if keys:
            mid_templates = [
                f"По сути это {keys[0]} — удобный вариант на каждый день.",
                f"Если коротко: это {keys[0]}, который хорошо смотрится и в городе, и в поездках.",
                f"Можно сказать так: {keys[0]} для тех, кто любит аккуратный стиль.",
            ]
            blocks.insert(min(2, len(blocks)), rnd.choice(mid_templates))

        tail_keys = [k for k in keys[1:3] if k] if len(keys) > 1 else []
        if tail_keys:
            tail_templates = [
                f"Подойдёт тем, кто выбирает {', '.join(tail_keys)}.",
                f"Хороший выбор, если нужны {', '.join(tail_keys)} без лишнего пафоса.",
                f"Если ищешь {', '.join(tail_keys)}, этот вариант будет уместным.",
            ]
            blocks.append(rnd.choice(tail_templates))

        text = " ".join(_cap(b).rstrip(".") + "." for b in blocks)
        text = re.sub(r"\s{2,}", " ", text).strip()

        if p.wb_safe:
            text = _apply_safe(text)
        if p.wb_strict:
            text = _apply_strict(text)

        # анти одинаковых начал
        sig2 = _norm(_first_sentences(text, 2))
        sig3 = _norm(_first_sentences(text, 3))
        if sig2 in used_open2 or sig3 in used_open3:
            continue

        # анти похожести
        if any(_jaccard(text, prev) > max_sim for prev in used_descs[-25:]):
            continue

        used_open2.add(sig2)
        used_open3.add(sig3)
        used_descs.append(text)
        return text

    return text


# =========================
# Excel helpers
# =========================
def _detect_header_row(ws, scan: int = 30) -> int:
    for r in range(1, min(scan, ws.max_row) + 1):
        row = [ws.cell(r, c).value for c in range(1, min(ws.max_column, 80) + 1)]
        joined = " | ".join([str(x) for x in row if x is not None])
        j = _norm(joined)
        if "наимен" in j and "описан" in j:
            return r
    return 1


def _find_col(ws, header_row: int, name_substr: str) -> Optional[int]:
    target = _norm(name_substr)
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is None:
            continue
        if target in _norm(str(v)):
            return c
    return None


def _set_cell(ws, r: int, c: int, value, overwrite: bool):
    cur = ws.cell(r, c).value
    if (cur is None or str(cur).strip() == ""):
        ws.cell(r, c).value = value
        return
    if overwrite:
        ws.cell(r, c).value = value


# =========================
# Main fill
# =========================
def fill_wb_template(params: FillParams) -> Tuple[List[str], int, str]:
    in_path = Path(params.xlsx_path)
    out_dir = Path(params.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    used_titles: Set[str] = set()
    used_descs: List[str] = []
    used_open2: Set[str] = set()
    used_open3: Set[str] = set()

    outputs: List[str] = []
    total_filled = 0

    total_steps = max(1, int(params.batch_count))

    for file_idx in range(1, int(params.batch_count) + 1):
        rnd = random.Random()
        rnd.seed((time.time_ns() & 0xFFFFFFFFFFFF) ^ (file_idx * 99991) ^ (hash(params.brand_lat) & 0xFFFFFFFF))

        wb = load_workbook(in_path)
        ws = wb.active

        header_row = _detect_header_row(ws)
        col_name = _find_col(ws, header_row, "Наименование")
        col_desc = _find_col(ws, header_row, "Описание")
        if not col_name or not col_desc:
            raise ValueError("Не найдены колонки Наименование и/или Описание (проверь заголовки)")

        # WB extra columns (Tom.xlsx)
        col_kiz = _find_col(ws, header_row, "КИЗ")
        col_18 = _find_col(ws, header_row, "18+")
        col_gender = _find_col(ws, header_row, "Пол")
        col_color = _find_col(ws, header_row, "Цвет")
        col_comp = _find_col(ws, header_row, "Состав")

        # rows area
        start_row = header_row + 1
        skip_until = max(0, int(params.skip_first_rows))
        eligible = [r for r in range(start_row, ws.max_row + 1) if r > skip_until]
        eligible = eligible[: max(0, int(params.rows_to_fill))]

        for r in eligible:
            title = make_title(
                rnd=rnd,
                brand_ru=params.brand_ru,
                shape=params.shape,
                lenses=params.lenses,
                ratio=params.brand_title_ratio,
                used=used_titles,
            )
            desc = make_description(
                rnd=rnd,
                p=params,
                used_descs=used_descs,
                used_open2=used_open2,
                used_open3=used_open3,
            )

            ws.cell(r, col_name).value = title
            ws.cell(r, col_desc).value = desc

            # WB fields fill (optional)
            if params.fill_wb_fields:
                ow = bool(params.overwrite_wb_fields_if_not_empty)

                if col_kiz:
                    _set_cell(ws, r, col_kiz, "да" if params.kiz else "нет", ow)
                if col_18:
                    _set_cell(ws, r, col_18, "да" if params.adult18 else "нет", ow)
                if col_gender and params.gender.strip():
                    _set_cell(ws, r, col_gender, params.gender.strip(), ow)
                if col_color and params.color.strip():
                    _set_cell(ws, r, col_color, params.color.strip(), ow)
                if col_comp and params.composition.strip():
                    _set_cell(ws, r, col_comp, params.composition.strip(), ow)

        total_filled += len(eligible)

        base = _safe_filename(in_path.stem)
        out_name = f"{base}_{file_idx:02d}.xlsx" if int(params.batch_count) > 1 else f"{base}_out.xlsx"
        out_path = out_dir / out_name
        wb.save(out_path)
        outputs.append(str(out_path))

        if params.progress_callback:
            params.progress_callback(int(file_idx * 100 / total_steps))

    report = {
        "input": str(in_path),
        "outputs": outputs,
        "rows_total_filled": total_filled,
        "rows_per_file": int(params.rows_to_fill),
        "batch_count": int(params.batch_count),
        "fill_wb_fields": bool(params.fill_wb_fields),
    }

    return outputs, total_filled, json.dumps(report, ensure_ascii=False, indent=2)
