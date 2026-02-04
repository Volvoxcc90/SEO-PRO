import json
import re
from pathlib import Path
import os
import logging

APP_NAME = "Sunglasses SEO PRO"

def app_data_dir() -> Path:
    base = Path(os.getenv("APPDATA", str(Path.home())))
    p = base / APP_NAME / "data"
    p.mkdir(parents=True, exist_ok=True)
    return p

def setup_logging():
    log_dir = app_data_dir() / "logs"
    log_dir.mkdir(exist_ok=True)
    logging.basicConfig(filename=log_dir / "app.log", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def normalize_brand_key(brand: str) -> str:
    b = (brand or "").strip().lower()
    b = b.replace("-", " ").replace("&", " ")
    b = re.sub(r"\s+", " ", b).strip()
    return b

TRANSLIT = [
    ("sch","ш"),("sh","ш"),("ch","ч"),("ya","я"),("yu","ю"),("yo","ё"),
    ("kh","х"),("ts","ц"),("ph","ф"),("th","т"),
    ("a","а"),("b","б"),("c","к"),("d","д"),("e","е"),("f","ф"),
    ("g","г"),("h","х"),("i","и"),("j","дж"),("k","к"),("l","л"),
    ("m","м"),("n","н"),("o","о"),("p","п"),("q","к"),("r","р"),
    ("s","с"),("t","т"),("u","у"),("v","в"),("w","в"),("x","кс"),
    ("y","и"),("z","з"),
]

def guess_ru(brand: str) -> str:
    if re.search(r"[А-Яа-яЁё]", brand):
        return brand
    key = normalize_brand_key(brand)
    out = []
    for w in key.split():
        ww = w
        for a,b in TRANSLIT:
            ww = ww.replace(a,b)
        out.append(ww)
    return " ".join(x.capitalize() for x in out)

def load_brands_ru() -> dict:
    p = app_data_dir() / "brands_ru.json"
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def save_brands_ru(m: dict):
    p = app_data_dir() / "brands_ru.json"
    p.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")

def auto_update_brand_map(brands: list):
    m = load_brands_ru()
    changed = False
    for b in brands:
        key = normalize_brand_key(b)
        if key and key not in m:
            m[key] = guess_ru(b)
            changed = True
    if changed:
        save_brands_ru(m)

def ensure_list(filename: str, defaults: list) -> Path:
    p = app_data_dir() / filename
    if not p.exists():
        p.write_text("\n".join(defaults), encoding="utf-8")
    return p

def load_list(filename: str, defaults: list) -> list:
    p = ensure_list(filename, defaults)
    return [x.strip() for x in p.read_text(encoding="utf-8").splitlines() if x.strip()]

def add_to_list(filename: str, value: str):
    value = value.strip()
    if not value:
        return
    p = app_data_dir() / filename
    items = load_list(filename, [])
    if value not in items:
        with p.open("a", encoding="utf-8") as f:
            f.write("\n" + value)
