# main.py
from __future__ import annotations

import sys
import os
import json
import re
from pathlib import Path
from typing import List, Dict, Optional

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog, QLineEdit,
    QVBoxLayout, QHBoxLayout, QGridLayout, QComboBox, QMessageBox,
    QProgressBar, QGroupBox, QCheckBox, QSpinBox, QDialog, QScrollArea,
    QInputDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

from wb_fill import FillParams, fill_wb_template


APP_NAME = "Sunglasses SEO PRO"


def app_data_dir() -> Path:
    base = Path(os.getenv("APPDATA", str(Path.home())))
    p = base / APP_NAME / "data"
    p.mkdir(parents=True, exist_ok=True)
    return p


def settings_path() -> Path:
    base = Path(os.getenv("APPDATA", str(Path.home())))
    p = base / APP_NAME
    p.mkdir(parents=True, exist_ok=True)
    return p / "settings.json"


def load_settings() -> Dict:
    p = settings_path()
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_settings(d: Dict):
    settings_path().write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


def _norm_key(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("&", " ").replace("-", " ")
    return re.sub(r"\s+", " ", s).strip()


def list_file(path: Path, defaults: List[str]) -> List[str]:
    if not path.exists():
        path.write_text("\n".join(defaults) + "\n", encoding="utf-8")
        return defaults[:]

    items = [x.strip() for x in path.read_text(encoding="utf-8").splitlines() if x.strip()]
    for d in defaults:
        if d not in items:
            items.append(d)
    return items


def add_to_list_file(path: Path, value: str):
    value = (value or "").strip()
    if not value:
        return

    items = []
    if path.exists():
        items = [x.strip() for x in path.read_text(encoding="utf-8").splitlines() if x.strip()]

    if value not in items:
        items.append(value)

    path.write_text("\n".join(items) + "\n", encoding="utf-8")


def brands_ru_path() -> Path:
    return app_data_dir() / "brands_ru.json"


def load_brands_ru() -> Dict[str, str]:
    p = brands_ru_path()
    if not p.exists():
        p.write_text("{}", encoding="utf-8")
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_brands_ru(m: Dict[str, str]):
    brands_ru_path().write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")


def brand_to_ru(brand_lat: str, m: Dict[str, str]) -> str:
    return m.get(_norm_key(brand_lat), brand_lat)


THEMES = {
    "Graphite": {
        "bg": "#0b0f17", "card": "#111827", "card2": "#0f172a",
        "text": "#e5e7eb", "muted": "#9ca3af", "accent": "#3b82f6",
        "accent2": "#7c3aed", "border": "#1f2937", "input": "#0b1220"
    },
    "Midnight": {
        "bg": "#070a12", "card": "#0b1220", "card2": "#0a1020",
        "text": "#e5e7eb", "muted": "#a3a3a3", "accent": "#2563eb",
        "accent2": "#6d28d9", "border": "#111827", "input": "#0a1020"
    },
    "Sepia": {
        "bg": "#0f0d0b", "card": "#15120f", "card2": "#1b1713",
        "text": "#f5f5f4", "muted": "#d6d3d1", "accent": "#f59e0b",
        "accent2": "#a855f7", "border": "#292524", "input": "#1b1713"
    }
}


def make_stylesheet(theme_name: str) -> str:
    t = THEMES.get(theme_name, THEMES["Graphite"])
    return f"""
    QWidget {{
        background: {t["bg"]};
        color: {t["text"]};
        font-family: "Segoe UI";
        font-size: 11pt;
    }}
    QGroupBox {{
        border: 1px solid {t["border"]};
        border-radius: 14px;
        margin-top: 10px;
        padding: 10px;
        background: {t["card"]};
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 12px;
        padding: 2px 8px;
        color: {t["text"]};
        background: transparent;
        font-weight: 600;
    }}
    QLineEdit, QComboBox, QSpinBox {{
        background: {t["input"]};
        border: 1px solid {t["border"]};
        border-radius: 10px;
        padding: 8px 12px;
        min-height: 28px;
        selection-background-color: {t["accent"]};
    }}
    QComboBox::drop-down {{
        border: 0px;
        width: 28px;
        margin-right: 6px;
    }}
    QPushButton {{
        background: {t["card2"]};
        border: 1px solid {t["border"]};
        border-radius: 12px;
        padding: 10px 14px;
        font-weight: 600;
    }}
    QPushButton:hover {{
        border: 1px solid {t["accent"]};
    }}
    QPushButton#Primary {{
        background: {t["accent"]};
        border: 1px solid {t["accent"]};
        color: white;
    }}
    QPushButton#Primary:hover {{
        background: {t["accent2"]};
        border: 1px solid {t["accent2"]};
    }}
    QPushButton#Plus {{
        background: {t["accent"]};
        border: 1px solid {t["accent"]};
        color: white;
        min-width: 46px;
        max-width: 46px;
    }}
    QProgressBar {{
        border: 1px solid {t["border"]};
        border-radius: 12px;
        text-align: center;
        background: {t["card2"]};
        height: 20px;
    }}
    QProgressBar::chunk {{
        background: {t["accent2"]};
        border-radius: 12px;
    }}
    QLabel#Muted {{
        color: {t["muted"]};
    }}
    """


class Worker(QThread):
    progress = pyqtSignal(int)
    done = pyqtSignal(list, int, str)
    fail = pyqtSignal(str)

    def __init__(self, params: FillParams):
        super().__init__()
        self.params = params

    def run(self):
        try:
            def cb(p: int):
                self.progress.emit(int(p))

            self.params.progress_callback = cb
            outs, total, rep = fill_wb_template(self.params)
            self.done.emit(outs, total, rep)
        except Exception as e:
            self.fail.emit(str(e))


class HolidaysDialog(QDialog):
    def __init__(self, holidays: List[str], selected: List[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Выбрать праздники")
        self.resize(430, 520)
        self._picked: List[str] = []
        selected_set = set(selected or [])

        root = QVBoxLayout(self)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        body = QWidget()
        v = QVBoxLayout(body)

        self.checks: List[QCheckBox] = []
        for h in holidays:
            h = h.strip()
            if not h:
                continue
            cb = QCheckBox(h)
            cb.setChecked(h in selected_set)
            self.checks.append(cb)
            v.addWidget(cb)

        v.addStretch(1)
        scroll.setWidget(body)
        root.addWidget(scroll)

        row = QHBoxLayout()
        btn_cancel = QPushButton("Отмена")
        btn_ok = QPushButton("OK")
        btn_ok.setObjectName("Primary")

        btn_cancel.clicked.connect(self.reject)
        btn_ok.clicked.connect(self._ok)

        row.addWidget(btn_cancel)
        row.addWidget(btn_ok)
        root.addLayout(row)

    def _ok(self):
        self._picked = [c.text().strip() for c in self.checks if c.isChecked()]
        self.accept()

    def picked(self) -> List[str]:
        return self._picked


class App(QWidget):
    def __init__(self):
        super().__init__()

        self.data_dir = app_data_dir()
        self.settings = load_settings()

        self.file_counter = int(self.settings.get("file_counter", 1))
        if self.file_counter < 1:
            self.file_counter = 1

        self.brands_file = self.data_dir / "brands.txt"
        self.shapes_file = self.data_dir / "shapes.txt"
        self.lenses_file = self.data_dir / "lenses.txt"
        self.holidays_file = self.data_dir / "holidays.txt"
        self.gender_file = self.data_dir / "gender.txt"
        self.colors_file = self.data_dir / "colors.txt"
        self.comp_file = self.data_dir / "composition.txt"

        self.brands = list_file(self.brands_file, ["Dior", "Gucci", "Prada", "Cazal", "Ray-Ban", "Miu Miu"])
        self.shapes = list_file(self.shapes_file, ["Кошачий глаз", "Квадратные", "Овальные", "Круглые", "Прямоугольные", "Авиаторы"])
        self.lenses = list_file(self.lenses_file, ["UV400", "Поляризационные", "Фотохромные (хамелеон)", "Градиентные"])
        self.holidays = list_file(self.holidays_file, ["8 Марта", "14 Февраля", "Новый год", "23 Февраля", "День рождения", "Выпускной"])
        self.genders = list_file(self.gender_file, ["Женский", "Мужской", "Унисекс"])
        self.colors = list_file(self.colors_file, ["Черный", "Белый", "Коричневый", "Бежевый", "Серый", "Леопард"])
        self.comps = list_file(self.comp_file, ["пластик", "металл", "ацетат", "поликарбонат"])

        self.brand_map = load_brands_ru()
        self.selected_holidays: List[str] = []
        self.xlsx_path: Optional[str] = None
        self.worker: Optional[Worker] = None

        self._build_ui()
        self._restore_settings()

        self.setMinimumSize(1050, 760)

    def _build_ui(self):
        self.setWindowTitle(APP_NAME)

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(12)

        header = QGroupBox()
        hl = QVBoxLayout(header)

        title = QLabel("😎 Sunglasses SEO PRO")
        title.setStyleSheet("font-size: 20pt; font-weight: 800;")
        sub = QLabel("Human Seller Engine • WB-поля • Праздники • Нумерация файлов")
        sub.setObjectName("Muted")

        self.lb_counter = QLabel(f"Нумерация файлов продолжится с: {self.file_counter:04d}")
        self.lb_counter.setObjectName("Muted")

        hl.addWidget(title)
        hl.addWidget(sub)
        hl.addWidget(self.lb_counter)
        root.addWidget(header)

        top = QGroupBox("Файлы")
        tl = QGridLayout(top)

        tl.addWidget(QLabel("Тема"), 0, 0)
        self.cmb_theme = QComboBox()
        self.cmb_theme.addItems(list(THEMES.keys()))
        self.cmb_theme.currentTextChanged.connect(self._apply_theme)
        tl.addWidget(self.cmb_theme, 0, 1)

        tl.addWidget(QLabel("Data"), 0, 2)
        self.ed_data = QLineEdit(str(self.data_dir))
        self.ed_data.setReadOnly(True)
        tl.addWidget(self.ed_data, 0, 3)

        self.btn_data = QPushButton("Папка")
        self.btn_data.clicked.connect(self._open_data_folder)
        tl.addWidget(self.btn_data, 0, 4)

        tl.addWidget(QLabel("Папка вывода"), 1, 0)
        self.ed_out = QLineEdit("")
        self.ed_out.setPlaceholderText("Выбери папку, куда сохранять все файлы")
        tl.addWidget(self.ed_out, 1, 1, 1, 3)

        self.btn_out = QPushButton("Выбрать")
        self.btn_out.clicked.connect(self._pick_out_dir)
        tl.addWidget(self.btn_out, 1, 4)

        self.btn_xlsx = QPushButton("⬇️ Загрузить XLSX")
        self.btn_xlsx.setObjectName("Primary")
        self.btn_xlsx.clicked.connect(self._pick_xlsx)
        tl.addWidget(self.btn_xlsx, 2, 0)

        self.lb_file = QLabel("Файл не выбран")
        self.lb_file.setObjectName("Muted")
        tl.addWidget(self.lb_file, 2, 1, 1, 4)

        root.addWidget(top)

        gen = QGroupBox("Параметры генерации")
        gl = QGridLayout(gen)

        row = 0

        gl.addWidget(QLabel("Бренд (латиницей)"), row, 0)
        self.cmb_brand = QComboBox()
        self.cmb_brand.setEditable(True)
        self.cmb_brand.addItems(self.brands)
        gl.addWidget(self.cmb_brand, row, 1, 1, 3)

        self.btn_add_brand = QPushButton("+")
        self.btn_add_brand.setObjectName("Plus")
        self.btn_add_brand.clicked.connect(self._add_brand)
        gl.addWidget(self.btn_add_brand, row, 4)
        row += 1

        gl.addWidget(QLabel("Форма оправы"), row, 0)
        self.cmb_shape = QComboBox()
        self.cmb_shape.setEditable(True)
        self.cmb_shape.addItems(self.shapes)
        gl.addWidget(self.cmb_shape, row, 1, 1, 3)

        self.btn_add_shape = QPushButton("+")
        self.btn_add_shape.setObjectName("Plus")
        self.btn_add_shape.clicked.connect(lambda: self._add_to(self.cmb_shape, self.shapes_file))
        gl.addWidget(self.btn_add_shape, row, 4)
        row += 1

        gl.addWidget(QLabel("Линзы"), row, 0)
        self.cmb_lenses = QComboBox()
        self.cmb_lenses.setEditable(True)
        self.cmb_lenses.addItems(self.lenses)
        gl.addWidget(self.cmb_lenses, row, 1, 1, 3)

        self.btn_add_lenses = QPushButton("+")
        self.btn_add_lenses.setObjectName("Plus")
        self.btn_add_lenses.clicked.connect(lambda: self._add_to(self.cmb_lenses, self.lenses_file))
        gl.addWidget(self.btn_add_lenses, row, 4)
        row += 1

        gl.addWidget(QLabel("Коллекция"), row, 0)
        self.cmb_collection = QComboBox()
        self.cmb_collection.setEditable(True)
        self.cmb_collection.addItems(["Весна–Лето 2026", "Весна–Лето 2025–2026"])
        gl.addWidget(self.cmb_collection, row, 1, 1, 4)
        row += 1

        gl.addWidget(QLabel("Праздники"), row, 0)
        self.ed_holidays = QLineEdit("")
        self.ed_holidays.setReadOnly(True)
        self.ed_holidays.setPlaceholderText("Выбери праздники")
        gl.addWidget(self.ed_holidays, row, 1, 1, 2)

        self.btn_holidays = QPushButton("Выбрать")
        self.btn_holidays.clicked.connect(self._pick_holidays)
        gl.addWidget(self.btn_holidays, row, 3)

        self.cmb_holiday_pos = QComboBox()
        self.cmb_holiday_pos.addItems(["middle", "start", "end"])
        gl.addWidget(self.cmb_holiday_pos, row, 4)
        row += 1

        gl.addWidget(QLabel("SEO"), row, 0)
        self.cmb_seo = QComboBox()
        self.cmb_seo.addItems(["low", "normal", "high"])
        gl.addWidget(self.cmb_seo, row, 1)

        gl.addWidget(QLabel("Стиль"), row, 2)
        self.cmb_style = QComboBox()
        self.cmb_style.addItems(["neutral", "premium", "mass", "social"])
        gl.addWidget(self.cmb_style, row, 3)

        self.cmb_brand_ratio = QComboBox()
        self.cmb_brand_ratio.addItems(["50/50", "100/0", "0/100"])
        gl.addWidget(self.cmb_brand_ratio, row, 4)
        row += 1

        gl.addWidget(QLabel("Строк заполнять"), row, 0)
        self.spin_rows = QSpinBox()
        self.spin_rows.setRange(1, 1000)
        self.spin_rows.setValue(6)
        gl.addWidget(self.spin_rows, row, 1)

        gl.addWidget(QLabel("Excel файлов"), row, 2)
        self.spin_batch = QSpinBox()
        self.spin_batch.setRange(1, 100)
        self.spin_batch.setValue(1)
        gl.addWidget(self.spin_batch, row, 3)
        row += 1

        gl.addWidget(QLabel("Не трогать первые строк"), row, 0)
        self.spin_skip = QSpinBox()
        self.spin_skip.setRange(0, 100)
        self.spin_skip.setValue(4)
        gl.addWidget(self.spin_skip, row, 1)
        row += 1

        self.chk_safe = QCheckBox("WB Safe Mode")
        self.chk_safe.setChecked(True)

        self.chk_strict = QCheckBox("WB Strict")
        self.chk_strict.setChecked(True)

        gl.addWidget(self.chk_safe, row, 0, 1, 2)
        gl.addWidget(self.chk_strict, row, 2, 1, 3)

        root.addWidget(gen)

        wbbox = QGroupBox("WB-поля из Tom.xlsx")
        wl = QGridLayout(wbbox)

        wl.addWidget(QLabel("Заполнять WB-поля"), 0, 0)
        self.chk_fill_wb = QCheckBox("Да")
        self.chk_fill_wb.setChecked(True)
        wl.addWidget(self.chk_fill_wb, 0, 1)

        self.chk_overwrite_wb = QCheckBox("Перезаписывать заполненные")
        self.chk_overwrite_wb.setChecked(False)
        wl.addWidget(self.chk_overwrite_wb, 0, 2, 1, 3)

        wl.addWidget(QLabel("КИЗ"), 1, 0)
        self.chk_kiz = QCheckBox("да")
        wl.addWidget(self.chk_kiz, 1, 1)

        wl.addWidget(QLabel("18+"), 1, 2)
        self.chk_18 = QCheckBox("да")
        wl.addWidget(self.chk_18, 1, 3)

        wl.addWidget(QLabel("Пол"), 2, 0)
        self.cmb_gender = QComboBox()
        self.cmb_gender.setEditable(True)
        self.cmb_gender.addItems(self.genders)
        wl.addWidget(self.cmb_gender, 2, 1, 1, 3)

        self.btn_add_gender = QPushButton("+")
        self.btn_add_gender.setObjectName("Plus")
        self.btn_add_gender.clicked.connect(lambda: self._add_to(self.cmb_gender, self.gender_file))
        wl.addWidget(self.btn_add_gender, 2, 4)

        wl.addWidget(QLabel("Цвет"), 3, 0)
        self.cmb_color = QComboBox()
        self.cmb_color.setEditable(True)
        self.cmb_color.addItems(self.colors)
        wl.addWidget(self.cmb_color, 3, 1, 1, 3)

        self.btn_add_color = QPushButton("+")
        self.btn_add_color.setObjectName("Plus")
        self.btn_add_color.clicked.connect(lambda: self._add_to(self.cmb_color, self.colors_file))
        wl.addWidget(self.btn_add_color, 3, 4)

        wl.addWidget(QLabel("Состав"), 4, 0)
        self.cmb_comp = QComboBox()
        self.cmb_comp.setEditable(True)
        self.cmb_comp.addItems(self.comps)
        wl.addWidget(self.cmb_comp, 4, 1, 1, 3)

        self.btn_add_comp = QPushButton("+")
        self.btn_add_comp.setObjectName("Plus")
        self.btn_add_comp.clicked.connect(lambda: self._add_to(self.cmb_comp, self.comp_file))
        wl.addWidget(self.btn_add_comp, 4, 4)

        root.addWidget(wbbox)

        footer = QGroupBox()
        fl = QHBoxLayout(footer)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        fl.addWidget(self.progress, 1)

        self.btn_go = QPushButton("🚀 СГЕНЕРИРОВАТЬ")
        self.btn_go.setObjectName("Primary")
        self.btn_go.clicked.connect(self._run)
        fl.addWidget(self.btn_go)

        root.addWidget(footer)

        self._apply_theme(self.cmb_theme.currentText())

    def _apply_theme(self, name: str):
        self.setStyleSheet(make_stylesheet(name))
        self.settings["theme"] = name
        save_settings(self.settings)

    def _open_data_folder(self):
        try:
            os.startfile(str(self.data_dir))
        except Exception:
            QMessageBox.information(self, "Data", str(self.data_dir))

    def _pick_out_dir(self):
        p = QFileDialog.getExistingDirectory(self, "Выбери папку вывода", self.ed_out.text().strip() or str(Path.home()))
        if p:
            self.ed_out.setText(p)
            self.settings["out_dir"] = p
            save_settings(self.settings)

    def _pick_xlsx(self):
        p, _ = QFileDialog.getOpenFileName(self, "Выбери XLSX", str(Path.home()), "Excel (*.xlsx)")
        if p:
            self.xlsx_path = p
            self.lb_file.setText(Path(p).name)
            self.settings["last_xlsx"] = p
            save_settings(self.settings)

    def _add_to(self, cmb: QComboBox, file_path: Path):
        v = cmb.currentText().strip()
        if not v:
            return

        add_to_list_file(file_path, v)
        items = list_file(file_path, [])

        cmb.blockSignals(True)
        cmb.clear()
        cmb.addItems(items)
        cmb.setCurrentText(v)
        cmb.blockSignals(False)

    def _add_brand(self):
        v = self.cmb_brand.currentText().strip()
        if not v:
            return

        add_to_list_file(self.brands_file, v)
        self.brands = list_file(self.brands_file, self.brands)

        self.cmb_brand.clear()
        self.cmb_brand.addItems(self.brands)
        self.cmb_brand.setCurrentText(v)

        ru, ok = QInputDialog.getText(self, "Бренд на кириллице", f"{v} → например: Миу Миу")
        if ok and ru.strip():
            m = load_brands_ru()
            m[_norm_key(v)] = ru.strip()
            save_brands_ru(m)
            self.brand_map = m

    def _pick_holidays(self):
        dlg = HolidaysDialog(self.holidays, self.selected_holidays, self)
        if dlg.exec_() == QDialog.Accepted:
            self.selected_holidays = dlg.picked()
            self.ed_holidays.setText(", ".join(self.selected_holidays))
            self.settings["holidays_multi"] = self.selected_holidays
            save_settings(self.settings)

    def _run(self):
        if not self.xlsx_path or not Path(self.xlsx_path).exists():
            QMessageBox.warning(self, "XLSX", "Сначала выбери XLSX файл")
            return

        out_dir = self.ed_out.text().strip()
        if not out_dir:
            QMessageBox.warning(self, "Папка вывода", "Выбери папку вывода")
            return

        Path(out_dir).mkdir(parents=True, exist_ok=True)

        brand_lat = self.cmb_brand.currentText().strip()
        if not brand_lat:
            QMessageBox.warning(self, "Бренд", "Выбери или введи бренд")
            return

        self.brand_map = load_brands_ru()
        brand_ru = brand_to_ru(brand_lat, self.brand_map)

        start_index = int(self.settings.get("file_counter", self.file_counter))
        if start_index < 1:
            start_index = 1

        params = FillParams(
            xlsx_path=self.xlsx_path,
            output_dir=out_dir,
            file_prefix=Path(self.xlsx_path).stem,
            start_index=start_index,

            brand_lat=brand_lat,
            brand_ru=brand_ru,
            shape=self.cmb_shape.currentText().strip(),
            lenses=self.cmb_lenses.currentText().strip(),
            collection=self.cmb_collection.currentText().strip(),

            holidays="||".join([h.strip() for h in self.selected_holidays if h.strip()]),
            holiday_pos=self.cmb_holiday_pos.currentText().strip(),

            seo_level=self.cmb_seo.currentText().strip(),
            style=self.cmb_style.currentText().strip(),

            wb_safe=self.chk_safe.isChecked(),
            wb_strict=self.chk_strict.isChecked(),

            brand_title_ratio=self.cmb_brand_ratio.currentText().strip(),

            rows_to_fill=int(self.spin_rows.value()),
            skip_first_rows=int(self.spin_skip.value()),
            batch_count=int(self.spin_batch.value()),

            fill_wb_fields=self.chk_fill_wb.isChecked(),
            overwrite_wb_fields_if_not_empty=self.chk_overwrite_wb.isChecked(),

            kiz=self.chk_kiz.isChecked(),
            adult18=self.chk_18.isChecked(),

            gender=self.cmb_gender.currentText().strip(),
            color=self.cmb_color.currentText().strip(),
            composition=self.cmb_comp.currentText().strip(),
        )

        self._persist()

        self.btn_go.setEnabled(False)
        self.progress.setValue(0)

        self.worker = Worker(params)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.done.connect(self._on_done)
        self.worker.fail.connect(self._on_fail)
        self.worker.start()

    def _on_done(self, outs: list, total: int, report: str):
        self.btn_go.setEnabled(True)
        self.progress.setValue(100)

        batch = int(self.spin_batch.value())
        self.file_counter = int(self.settings.get("file_counter", self.file_counter))
        if self.file_counter < 1:
            self.file_counter = 1

        self.file_counter += batch
        self.settings["file_counter"] = self.file_counter
        save_settings(self.settings)

        self.lb_counter.setText(f"Нумерация файлов продолжится с: {self.file_counter:04d}")

        msg = f"Готово ✅\n\nФайлов: {len(outs)}\nСтрок заполнено: {total}\n\n" + "\n".join(outs[:10])
        QMessageBox.information(self, "Готово", msg)

    def _on_fail(self, err: str):
        self.btn_go.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", err)

    def _persist(self):
        self.settings["theme"] = self.cmb_theme.currentText()
        self.settings["out_dir"] = self.ed_out.text().strip()
        self.settings["brand"] = self.cmb_brand.currentText().strip()
        self.settings["shape"] = self.cmb_shape.currentText().strip()
        self.settings["lenses"] = self.cmb_lenses.currentText().strip()
        self.settings["collection"] = self.cmb_collection.currentText().strip()
        self.settings["holiday_pos"] = self.cmb_holiday_pos.currentText().strip()
        self.settings["seo"] = self.cmb_seo.currentText().strip()
        self.settings["style"] = self.cmb_style.currentText().strip()
        self.settings["brand_ratio"] = self.cmb_brand_ratio.currentText().strip()

        self.settings["rows"] = int(self.spin_rows.value())
        self.settings["batch"] = int(self.spin_batch.value())
        self.settings["skip"] = int(self.spin_skip.value())

        self.settings["safe"] = bool(self.chk_safe.isChecked())
        self.settings["strict"] = bool(self.chk_strict.isChecked())
        self.settings["holidays_multi"] = self.selected_holidays

        self.settings["fill_wb"] = bool(self.chk_fill_wb.isChecked())
        self.settings["ow_wb"] = bool(self.chk_overwrite_wb.isChecked())
        self.settings["kiz"] = bool(self.chk_kiz.isChecked())
        self.settings["adult18"] = bool(self.chk_18.isChecked())
        self.settings["gender"] = self.cmb_gender.currentText().strip()
        self.settings["color"] = self.cmb_color.currentText().strip()
        self.settings["comp"] = self.cmb_comp.currentText().strip()

        self.settings["file_counter"] = int(self.settings.get("file_counter", self.file_counter))

        save_settings(self.settings)

    def _restore_settings(self):
        theme = self.settings.get("theme", "Graphite")
        if theme in THEMES:
            self.cmb_theme.setCurrentText(theme)

        self._apply_theme(self.cmb_theme.currentText())

        self.ed_out.setText(self.settings.get("out_dir", ""))

        def set_combo(cmb: QComboBox, key: str):
            v = (self.settings.get(key) or "").strip()
            if v:
                cmb.setCurrentText(v)

        set_combo(self.cmb_brand, "brand")
        set_combo(self.cmb_shape, "shape")
        set_combo(self.cmb_lenses, "lenses")
        set_combo(self.cmb_collection, "collection")
        set_combo(self.cmb_holiday_pos, "holiday_pos")
        set_combo(self.cmb_seo, "seo")
        set_combo(self.cmb_style, "style")
        set_combo(self.cmb_brand_ratio, "brand_ratio")
        set_combo(self.cmb_gender, "gender")
        set_combo(self.cmb_color, "color")
        set_combo(self.cmb_comp, "comp")

        self.spin_rows.setValue(int(self.settings.get("rows", 6)))
        self.spin_batch.setValue(int(self.settings.get("batch", 1)))
        self.spin_skip.setValue(int(self.settings.get("skip", 4)))

        self.chk_safe.setChecked(bool(self.settings.get("safe", True)))
        self.chk_strict.setChecked(bool(self.settings.get("strict", True)))
        self.chk_fill_wb.setChecked(bool(self.settings.get("fill_wb", True)))
        self.chk_overwrite_wb.setChecked(bool(self.settings.get("ow_wb", False)))
        self.chk_kiz.setChecked(bool(self.settings.get("kiz", False)))
        self.chk_18.setChecked(bool(self.settings.get("adult18", False)))

        saved_h = self.settings.get("holidays_multi", [])
        if isinstance(saved_h, list):
            self.selected_holidays = [str(x) for x in saved_h if str(x).strip()]
        self.ed_holidays.setText(", ".join(self.selected_holidays))

        last_xlsx = self.settings.get("last_xlsx", "")
        if last_xlsx and Path(last_xlsx).exists():
            self.xlsx_path = last_xlsx
            self.lb_file.setText(Path(last_xlsx).name)

        self.file_counter = int(self.settings.get("file_counter", self.file_counter))
        if self.file_counter < 1:
            self.file_counter = 1

        self.lb_counter.setText(f"Нумерация файлов продолжится с: {self.file_counter:04d}")


def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    w = App()
    w.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
