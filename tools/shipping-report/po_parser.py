"""解析采购订单 PDF，提取订单号与物料行。"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import fitz


@dataclass
class PoItem:
    part_no: str
    name: str
    qty: int


@dataclass
class PurchaseOrder:
    po_no: str
    items: list[PoItem]
    path: str


_PART_RE = re.compile(r"\b(\d{3}-\d{7,}[A-Z0-9]*)\b")
# 通用订单号：以字母开头，后接字母/数字（如 PAQ703802、PDQ800530、P9Q80157）
# 不含连字符，避免误认料号 221-xxxx / 277-xxxx
_PO_NO_RE = re.compile(r"\b([A-Za-z][A-Za-z0-9]{4,13})\b")


def _extract_po_no(text: str, path: Path) -> str:
    """从订单正文优先在「订单号」附近提取；否则看页眉区域；再退回文件名。"""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for i, ln in enumerate(lines):
        if "订单号" in ln or ln.upper() in ("PO", "PO.", "P.O.", "P.O"):
            window = " ".join(lines[i : i + 4])
            after = ln.split("订单号", 1)[-1] if "订单号" in ln else ln
            for chunk in (after, window):
                m = _PO_NO_RE.search(chunk)
                if m:
                    return m.group(1).upper()
    # 只扫正文前部（页眉/订单信息区），避免后面规格里的编码干扰
    head = text[:1800]
    m = _PO_NO_RE.search(head)
    if m:
        return m.group(1).upper()
    stem = re.sub(r"\(\d+\)$", "", path.stem)
    m2 = _PO_NO_RE.search(stem)
    return m2.group(1).upper() if m2 else stem.upper()


def parse_purchase_order(pdf_path: str | Path) -> PurchaseOrder:
    path = Path(pdf_path)
    doc = fitz.open(path)
    text = "\n".join(page.get_text() for page in doc)
    doc.close()

    po_no = _extract_po_no(text, path)

    items: list[PoItem] = []
    lines = [ln.strip() for ln in text.splitlines()]

    # 行内：料号 … 数量（EA 后常见 单价 数量 金额）
    i = 0
    while i < len(lines):
        line = lines[i]
        m = _PART_RE.search(line)
        if not m:
            i += 1
            continue
        part_no = m.group(1)
        # 品名：料号同行剩余，或后续若干行直到遇到 EA/数字单价
        name_parts: list[str] = []
        after = line[m.end() :].strip(" ,，")
        if after:
            name_parts.append(after)

        qty = None
        window = lines[i : i + 12]
        blob = "\n".join(window)
        # 优先：EA + 单价 + 数量
        ea_m = re.search(
            r"EA\s+(\d+(?:\.\d+)?)\s+(\d+)\s+(\d+(?:\.\d+)?)",
            blob,
            re.I,
        )
        if ea_m:
            qty = int(ea_m.group(2))
            # 收集 EA 之前的品名碎片
            ea_pos = blob.lower().find("ea")
            name_blob = blob[:ea_pos]
            name_blob = _PART_RE.sub("", name_blob, count=1)
            name_parts = [p.strip() for p in name_blob.splitlines() if p.strip()]
        else:
            # 退化：找料号后较大的整数当数量
            nums = [int(x) for x in re.findall(r"\b(\d{2,6})\b", blob)]
            # 排除日期片段
            nums = [n for n in nums if n not in (2026, 2025, 2024, 2027)]
            if nums:
                # 数量通常不是单价小数；取中等偏大且不像金额的
                qty = nums[0]

        name = "".join(name_parts)
        name = re.sub(r"\s+", "", name)
        if len(name) > 40:
            name = name[:40] + "…"

        if qty is not None and not any(it.part_no == part_no for it in items):
            items.append(PoItem(part_no=part_no, name=name or part_no, qty=qty))
        # 每次只前进一行，避免重复 +1 漏扫料号
        i += 1

    # 若上面解析失败，退化为全局找料号 + 邻近数量
    if not items:
        for m in _PART_RE.finditer(text):
            part_no = m.group(1)
            tail = text[m.end() : m.end() + 200]
            ea_m = re.search(r"EA\s+(\d+(?:\.\d+)?)\s+(\d+)\s+", tail, re.I)
            if ea_m and not any(it.part_no == part_no for it in items):
                items.append(PoItem(part_no=part_no, name=part_no, qty=int(ea_m.group(2))))

    return PurchaseOrder(po_no=po_no, items=items, path=str(path))


def find_purchase_orders(folder: str | Path) -> list[Path]:
    folder = Path(folder)
    results = []
    for p in folder.glob("*.pdf"):
        if p.parent.name == "生成结果":
            continue
        try:
            doc = fitz.open(p)
            t = doc[0].get_text()[:800]
            doc.close()
        except Exception:
            continue
        if "采购订单" in t or _PO_NO_RE.search(t):
            # 排除出货报告（含 Outgoing Quality）
            if "Outgoing Quality" in t or "出货检验" in t:
                continue
            results.append(p)
    return sorted(results)


def find_template_for_part(folder: str | Path, part_no: str) -> Path | None:
    """按料号找首份出货报告模板。"""
    folder = Path(folder)
    part_no = part_no.upper()
    candidates: list[Path] = []
    for p in folder.glob("*.pdf"):
        if "生成结果" in p.parts:
            continue
        stem = p.stem.upper()
        # 221-0122001RU(1) / 221-0122001RU
        base = re.sub(r"\(\d+\)$", "", stem)
        if base == part_no or stem.startswith(part_no):
            # 排除采购订单
            try:
                doc = fitz.open(p)
                t = doc[0].get_text()[:500]
                doc.close()
            except Exception:
                continue
            if "Outgoing Quality" in t or "出货检验" in t or "样本代字" in t:
                candidates.append(p)
    if not candidates:
        return None
    # 优先文件名带 (1) 的“首份”，否则取名字最短
    candidates.sort(key=lambda x: (0 if re.search(r"\(1\)$", x.stem) else 1, len(x.name)))
    return candidates[0]


def detect_part_no_from_template(pdf_path: str | Path) -> str | None:
    """从出货报告模板的文件名或正文识别料号。"""
    path = Path(pdf_path)
    stem = re.sub(r"\(\d+\)$", "", path.stem).upper()
    # 文件名可能带 tpl_ 等前缀，不能用 \b（下划线也算单词字符）
    m = re.search(r"(\d{3}-\d{7,}[A-Z0-9]*)", stem)
    if m:
        return m.group(1)

    try:
        doc = fitz.open(path)
        text = doc[0].get_text()[:2000]
        doc.close()
    except Exception:
        return None

    for m in _PART_RE.finditer(text):
        return m.group(1)
    m2 = re.search(r"(\d{3}-\d{7,}[A-Z0-9]*)", text.replace("\n", ""))
    if m2:
        return m2.group(1)
    return None


def match_templates(
    template_paths: list[Path],
    part_nos: list[str],
) -> dict[str, Path]:
    """
    将多选的模板匹配到料号。
    返回 {料号: 模板路径}；同一料号多个模板时优先文件名带 (1) 的。
    """
    part_set = {p.upper(): p for p in part_nos}  # upper -> original
    buckets: dict[str, list[Path]] = {p.upper(): [] for p in part_nos}

    unmatched_files: list[Path] = []
    for path in template_paths:
        detected = detect_part_no_from_template(path)
        if detected and detected.upper() in buckets:
            buckets[detected.upper()].append(path)
        else:
            # 再试文件名是否以某个料号开头
            stem = path.stem.upper()
            hit = None
            for up in buckets:
                if stem.startswith(up) or re.sub(r"\(\d+\)$", "", stem) == up:
                    hit = up
                    break
            if hit:
                buckets[hit].append(path)
            else:
                unmatched_files.append(path)

    result: dict[str, Path] = {}
    for up, paths in buckets.items():
        if not paths:
            continue
        paths.sort(key=lambda x: (0 if re.search(r"\(1\)$", x.stem) else 1, len(x.name)))
        result[part_set[up]] = paths[0]
    return result
