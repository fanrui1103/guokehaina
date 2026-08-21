"""从标签图片里读出字段。按字段名取值，不写死某一张标签上的数字。"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from PIL import Image

from paths import bundle_dir

FIELD_MAP_PATH = bundle_dir() / "field_map.json"

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}

_ocr_engine = None
_field_map = None


def load_field_map() -> dict:
    global _field_map
    if _field_map is None:
        _field_map = json.loads(FIELD_MAP_PATH.read_text(encoding="utf-8"))
    return _field_map


def get_ocr():
    global _ocr_engine
    if _ocr_engine is None:
        import logging

        logging.getLogger("RapidOCR").setLevel(logging.ERROR)
        from rapidocr import RapidOCR

        _ocr_engine = RapidOCR()
    return _ocr_engine


def is_label_image(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTS


def decode_barcodes(image: Image.Image) -> list[str]:
    try:
        from pyzbar.pyzbar import decode
    except Exception:
        return []
    codes = []
    for item in decode(image):
        try:
            text = item.data.decode("utf-8", errors="replace").strip()
        except Exception:
            continue
        if text:
            codes.append(text)
    return codes


def _norm_key(text: str) -> str:
    text = text.strip()
    text = text.replace(" ", "").replace("　", "")
    text = text.replace(".", "").replace("·", "")
    return text.upper()


def _split_kv(text: str) -> tuple[str, str] | None:
    for sep in ("：", ":", "﹔", ";"):
        if sep in text:
            left, right = text.split(sep, 1)
            left, right = left.strip(), right.strip()
            if left:
                return left, right
    return None


def _parse_date(value: str):
    value = value.strip().replace("年", "/").replace("月", "/").replace("日", "")
    value = value.replace(".", "/").replace("-", "/")
    value = re.sub(r"\s+", "", value)
    try:
        return datetime.strptime(value, "%Y/%m/%d")
    except ValueError:
        pass
    m = re.fullmatch(r"(\d{4})/(\d{1,2})/(\d{1,2})", value)
    if m:
        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", value)
    if m:
        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def _to_number_if_possible(value: str):
    text = value.strip().replace(",", "")
    if re.fullmatch(r"\d+", text):
        try:
            return int(text)
        except ValueError:
            return text
    return value.strip()


def _alias_lookup(cfg: dict) -> dict[str, str]:
    table = {}
    for field_name, spec in cfg["fields"].items():
        for alias in spec.get("aliases", []):
            table[_norm_key(alias)] = field_name
    return table


def _pick_clid(barcodes: list[str], ocr_texts: list[str], cfg: dict) -> str:
    ignore = tuple(cfg.get("ignore_barcode_prefixes", []))
    pattern = re.compile(cfg.get("clid_pattern", r"[A-Z]{2,}[A-Z0-9]{10,}"))

    def ok(text: str) -> bool:
        if any(text.startswith(p) for p in ignore):
            return False
        if re.fullmatch(r"\d{6,14}", text):
            return False
        return bool(pattern.fullmatch(text.replace(" ", ""))) or (
            text.isalnum() and len(text) >= 15 and re.search(r"[A-Z]", text)
        )

    for code in barcodes:
        if ok(code):
            return code
    for text in ocr_texts:
        compact = text.replace(" ", "")
        m = pattern.search(compact)
        if m:
            return m.group(0)
        if ok(compact):
            return compact
    return ""


def parse_ocr_items(texts: list[str], barcodes: list[str] | None = None) -> dict:
    cfg = load_field_map()
    alias = _alias_lookup(cfg)
    found: dict[str, str] = {}
    barcodes = barcodes or []

    for raw in texts:
        text = raw.strip()
        if not text:
            continue
        kv = _split_kv(text)
        if not kv:
            continue
        key, value = kv
        if not value:
            continue
        field = alias.get(_norm_key(key))
        if field and field not in found:
            found[field] = value.strip()

    clid = _pick_clid(barcodes, texts, cfg)
    if clid:
        found["CLID"] = clid

    if "ManufactureDate" not in found and "DateCode" in found:
        found["ManufactureDate"] = found["DateCode"]
    if "DateCode" not in found and "ManufactureDate" in found:
        found["DateCode"] = found["ManufactureDate"]

    if "ManufacturerPN" not in found and "ArtesynPN" in found:
        found["ManufacturerPN"] = found["ArtesynPN"]
    if "ArtesynPN" not in found and "ManufacturerPN" in found:
        found["ArtesynPN"] = found["ManufacturerPN"]

    maker = found.get("Manufacturer", "")
    mapping = cfg.get("manufacturer_normalize", {})
    if maker:
        found["Manufacturer"] = mapping.get(str(maker).upper(), maker)

    for date_field in ("DateCode", "ManufactureDate", "ExpDate"):
        if date_field in found:
            dt = _parse_date(str(found[date_field]))
            if dt:
                found[date_field] = dt

    for num_field in ("UnitQty", "LotNo", "ArtesynPN", "ManufacturerPN"):
        if num_field in found and not isinstance(found[num_field], datetime):
            found[num_field] = _to_number_if_possible(str(found[num_field]))

    return found


def extract_label(path: str | Path) -> dict:
    path = Path(path)
    image = Image.open(path).convert("RGB")
    barcodes = decode_barcodes(image)

    result = get_ocr()(str(path))
    texts = [str(t).strip() for t in (result.txts or []) if str(t).strip()]
    fields = parse_ocr_items(texts, barcodes)

    return {
        "file": str(path),
        "name": path.name,
        "barcodes": barcodes,
        "ocr_texts": texts,
        "fields": fields,
        "ok": bool(fields.get("CLID") or fields.get("ArtesynPN")),
    }


def extract_many(paths: list[str | Path], progress=None) -> list[dict]:
    rows = []
    total = len(paths)
    for i, path in enumerate(paths, 1):
        if progress:
            progress(i, total, Path(path).name)
        try:
            rows.append(extract_label(path))
        except Exception as exc:
            rows.append(
                {
                    "file": str(path),
                    "name": Path(path).name,
                    "barcodes": [],
                    "ocr_texts": [],
                    "fields": {},
                    "ok": False,
                    "error": str(exc),
                }
            )
    return rows
