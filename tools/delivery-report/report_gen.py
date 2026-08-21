"""基于出货报告 PDF 模板，改写：日期、订单编号、订单/交货数量、尺寸样本数据。"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import fitz

from size_gen import generate_samples_smart, parse_all_tolerances, pick_tolerance_for_samples


@dataclass
class GenerateRequest:
    template_path: Path
    output_path: Path
    po_no: str
    order_qty: int
    deliver_qty: int
    ship_date: date | None = None


def _fmt_date(d: date) -> str:
    # 与现有报告一致：2026/8/10（月日不补零）
    return f"{d.year}/{d.month}/{d.day}"


def _span_fontsize(span: dict) -> float:
    return float(span.get("size") or 9)


def _iter_spans(page: fitz.Page):
    data = page.get_text("dict")
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text") or ""
                if not text.strip():
                    continue
                yield span, fitz.Rect(span["bbox"])


def _find_label_rect(page: fitz.Page, labels: list[str]) -> fitz.Rect | None:
    for span, rect in _iter_spans(page):
        t = span["text"].strip()
        for lab in labels:
            if lab in t:
                return rect
    return None


# ---------------------------------------------------------------------------
# 擦除 / 写入
#
# 残留根因（如「460 460…」「20000 CS」）：
# 1) PDF 文字 bbox 常比肉眼墨迹略窄，只按 bbox 擦会漏掉边缘笔画
# 2) 数量单位 PCS 常与数字分属两个 span；只擦数字会留下 CS
# 3) 仅靠 redact 有时擦不净，需再盖一层白块
#
# 做法：旧值矩形加宽加高边距 → redact → 同范围白块覆盖 → 再写入新字
# 右侧仍受 x1_limit 约束，避免擦到隔壁标签。
# ---------------------------------------------------------------------------

# 提取文字时用更贴合字形的高度，减少误伤上下行
try:
    fitz.TOOLS.set_small_glyph_heights(True)
except Exception:
    pass


def _clamp_erase(
    rect: fitz.Rect,
    *,
    x1_limit: float | None = None,
    x0_limit: float | None = None,
) -> fitz.Rect:
    r = fitz.Rect(rect)
    if x0_limit is not None:
        r.x0 = max(r.x0, x0_limit + 0.5)
    if x1_limit is not None:
        r.x1 = min(r.x1, x1_limit - 1.0)
    if r.x1 < r.x0 + 4:
        r.x1 = r.x0 + 4
    return r


def _erase_rect_for_old_value(
    old_rect: fitz.Rect,
    *,
    x1_limit: float | None = None,
    x0_limit: float | None = None,
    pad_left: float = 3.5,
    pad_right: float = 4.0,
) -> fitz.Rect:
    """
    专为「擦掉旧字」：水平多留边距盖住笔画外溢；高度略收，减少伤上下行。
    不按新字宽度扩展（新字更长时叠在空白上即可，绝不为新字去擦邻居）。
    pad_left 对页眉日期应更小，避免擦掉「Date：」的冒号。
    """
    x0 = old_rect.x0 - pad_left
    x1 = old_rect.x1 + pad_right
    h = old_rect.height
    cy = (old_rect.y0 + old_rect.y1) / 2
    half = max(h * 0.39, 4.5)
    y0 = cy - half
    y1 = cy + half
    return _clamp_erase(
        fitz.Rect(x0, y0, x1, y1),
        x1_limit=x1_limit,
        x0_limit=x0_limit,
    )


def _insert_text(page: fitz.Page, anchor: fitz.Rect, new_text: str, fontsize: float) -> None:
    baseline = fitz.Point(anchor.x0, anchor.y1 - 1.2)
    page.insert_text(
        baseline,
        new_text,
        fontsize=fontsize,
        fontname="helv",
        color=(0, 0, 0),
    )


def _batch_replace(
    page: fitz.Page,
    jobs: list[tuple[fitz.Rect, str, float, float | None, float, float]],
) -> None:
    """
    jobs: (旧值矩形, 新文字, 字号, 右侧上限, pad_left, pad_right)
    步骤：登记 redact → 一次 apply → 白块再盖一遍 → 写入新字
    """
    if not jobs:
        return

    erase_rects: list[fitz.Rect] = []
    for old_rect, _new_text, _fs, x1_limit, pad_left, pad_right in jobs:
        erase_rects.append(
            _erase_rect_for_old_value(
                old_rect,
                x1_limit=x1_limit,
                pad_left=pad_left,
                pad_right=pad_right,
            )
        )

    for r in erase_rects:
        page.add_redact_annot(r, fill=(1, 1, 1))
    page.apply_redactions(
        images=fitz.PDF_REDACT_IMAGE_NONE,
        graphics=fitz.PDF_REDACT_LINE_ART_NONE,
    )

    for r in erase_rects:
        shape = page.new_shape()
        shape.draw_rect(r)
        shape.finish(color=(1, 1, 1), fill=(1, 1, 1), width=0)
        shape.commit()

    for old_rect, new_text, fs, _lim, _pl, _pr in jobs:
        _insert_text(page, old_rect, new_text, fs)


_QTY_TEXT_RE = re.compile(
    r"^(?P<num>\d{1,10})\s*(?P<unit>pcs|PCS|Pcs|PC|CS|个|EA|ea)?\.?$",
)
_UNIT_TOKENS = ("pcs", "PCS", "Pcs", "PC", "CS", "S", "个", "EA", "ea")


def _parse_qty_span(text: str) -> tuple[str, str] | None:
    """返回 (数字, 原单位后缀)。"""
    t = (text or "").strip()
    m = _QTY_TEXT_RE.fullmatch(t)
    if not m:
        return None
    return m.group("num"), (m.group("unit") or "")


def _format_qty(qty: int, unit: str) -> str:
    return f"{qty}{unit}" if unit else str(qty)


def _find_nearby_unit(
    page: fitz.Page, rect: fitz.Rect
) -> tuple[fitz.Rect | None, str]:
    """在数量数字右侧找单位 span（含残留碎片 CS/S/PC）。"""
    best = None
    best_dx = 1e9
    best_unit = ""
    for span, r in _iter_spans(page):
        t = span["text"].strip()
        if t not in _UNIT_TOKENS and not re.fullmatch(r"[Pp]?[Cc][Ss]\.?", t):
            continue
        if abs(r.y0 - rect.y0) > 10:
            continue
        if r.x0 < rect.x1 - 2:
            continue
        dx = r.x0 - rect.x1
        if dx > 22:
            continue
        if dx < best_dx:
            best_dx = dx
            best = r
            if t == "Pcs":
                best_unit = "Pcs"
            elif t == "pcs":
                best_unit = "pcs"
            else:
                # PCS / PC / CS / S 等统一写回 PCS
                best_unit = "PCS"
    return best, best_unit


def _expand_qty_hit(
    page: fitz.Page, rect: fitz.Rect, raw: str, fs: float
) -> tuple[fitz.Rect, float, str, str]:
    """
    返回 (旧值总矩形, 字号, 原始文本, 写入用单位)。
    单位必须并入旧值矩形，否则会留下「CS」。
    """
    parsed = _parse_qty_span(raw)
    unit = parsed[1] if parsed else ""
    out = fitz.Rect(rect)

    if unit in ("PC", "CS", "S"):
        unit = "PCS"

    if not unit:
        ur, uu = _find_nearby_unit(page, rect)
        if ur is not None:
            out |= ur
            unit = uu or "PCS"

    if not unit:
        unit = "PCS"
        # 预留单位宽度，把可能存在的单位墨迹一并纳入擦除
        out = fitz.Rect(out.x0, out.y0, out.x1 + 32, out.y1)
    elif parsed and parsed[1] and parsed[1] not in raw:
        pass
    elif not (parsed and parsed[1]):
        # 数字 span 本身无单位但已并入右侧单位
        pass

    return out, fs, raw, unit


def _find_qty_near_label(
    page: fitz.Page,
    label: fitz.Rect,
    *,
    max_dy: float = 36,
    max_dx: float = 220,
    x_min: float | None = None,
    x_max: float | None = None,
) -> tuple[fitz.Rect, float, str, str] | None:
    best = None
    best_score = 1e18
    for span, rect in _iter_spans(page):
        raw = span["text"].strip()
        parsed = _parse_qty_span(raw)
        if parsed is None:
            continue
        # 纯单位碎片不算数量
        if not parsed[0]:
            continue
        if abs(rect.y0 - label.y0) > max_dy:
            continue
        if rect.x0 < label.x0 - 10:
            continue
        dx = rect.x0 - label.x0
        if dx > max_dx:
            continue
        if x_min is not None and rect.x0 < x_min:
            continue
        if x_max is not None and rect.x0 > x_max:
            continue
        score = abs(rect.y0 - label.y0) * 1000 + dx
        if score < best_score:
            best_score = score
            best = (rect, _span_fontsize(span), raw)
    if not best:
        return None
    rect, fs, raw = best
    return _expand_qty_hit(page, rect, raw, fs)


def _find_qty_pair(
    page: fitz.Page,
) -> tuple[tuple[fitz.Rect, float, str, str], tuple[fitz.Rect, float, str, str]] | None:
    order_label = _find_label_rect(page, ["订单数量"])
    if not order_label:
        order_label = _find_label_rect(page, ["Quantity:"])
    deliver_label = _find_label_rect(page, ["交货数量"])
    if not deliver_label:
        deliver_label = _find_label_rect(page, ["Delivered"])

    if not order_label or not deliver_label:
        return None

    order_hit = _find_qty_near_label(
        page,
        order_label,
        x_max=deliver_label.x0 - 4,
        max_dx=deliver_label.x0 - order_label.x0,
    )
    deliver_hit = _find_qty_near_label(
        page,
        deliver_label,
        x_min=deliver_label.x0 - 5,
        x_max=500,
        max_dx=220,
    )
    if order_hit and deliver_hit:
        if abs(order_hit[0].x0 - deliver_hit[0].x0) < 3:
            return None
        r2, fs2, raw2, u2 = deliver_hit
        if re.search(r"Pcs", raw2):
            u2 = "Pcs"
        elif re.search(r"pcs", raw2) and "PCS" not in raw2:
            u2 = "pcs"
        deliver_hit = (r2, fs2, raw2, u2)
        return order_hit, deliver_hit
    return None


_DATE_RE = re.compile(r"20\d{2}/\d{1,2}/\d{1,2}")


def _find_date_instances(page: fitz.Page) -> list[tuple[fitz.Rect, float, str, str]]:
    found: list[tuple[fitz.Rect, float, str, str]] = []
    data = page.get_text("dict")
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            parts: list[tuple[str, fitz.Rect, float]] = []
            for span in line.get("spans", []):
                text = span.get("text") or ""
                if not text:
                    continue
                parts.append((text, fitz.Rect(span["bbox"]), float(span.get("size") or 9)))
            if not parts:
                continue

            joined = "".join(t for t, _, _ in parts)
            owner: list[int] = []
            for i, (t, _, _) in enumerate(parts):
                owner.extend([i] * len(t))

            for m in _DATE_RE.finditer(joined):
                idxs = sorted(set(owner[m.start() : m.end()]))
                if not idxs:
                    continue
                rects = [parts[i][1] for i in idxs]
                union = rects[0]
                for r in rects[1:]:
                    union |= r
                fs = parts[idxs[0]][2]
                kind = "plain"
                prefix = joined[: m.start()]
                if "Date" in prefix:
                    kind = "header"
                found.append((union, fs, m.group(0), kind))
    return found


def _find_dates(page: fitz.Page) -> dict[str, tuple[fitz.Rect, float, str]]:
    found: dict[str, tuple[fitz.Rect, float, str]] = {}
    ship_label = _find_label_rect(page, ["出货日期"])
    if not ship_label:
        ship_label = _find_label_rect(page, ["Shipment"])
    test_label = _find_label_rect(page, ["检验日期"])
    if not test_label:
        test_label = _find_label_rect(page, ["QC Date", "Test dat"])

    instances = _find_date_instances(page)
    headers = [(r, fs, t) for r, fs, t, k in instances if k == "header"]
    plains = [(r, fs, t) for r, fs, t, k in instances if k == "plain"]

    if headers:
        headers.sort(key=lambda x: x[0].y0)
        found["header_date_full"] = headers[0]

    def nearest(label: fitz.Rect | None, pool: list) -> tuple | None:
        if not label or not pool:
            return None
        best = None
        best_score = 1e18
        for rect, fs, text in pool:
            if abs(rect.y0 - label.y0) > 40:
                continue
            if rect.x0 < label.x0 - 20:
                continue
            score = abs(rect.y0 - label.y0) * 1000 + abs(rect.x0 - label.x0)
            if score < best_score:
                best_score = score
                best = (rect, fs, text)
        return best

    pool = plains if plains else [(r, fs, t) for r, fs, t, _k in instances]
    # 页眉日期不参与出货/检验匹配
    header_rect = found["header_date_full"][0] if "header_date_full" in found else None

    def not_header(item: tuple) -> bool:
        if not header_rect:
            return True
        r = item[0]
        return abs(r.x0 - header_rect.x0) > 2 or abs(r.y0 - header_rect.y0) > 2

    pool = [p for p in pool if not_header(p)]

    ship = nearest(ship_label, pool)
    if ship:
        found["ship"] = ship
    test = nearest(test_label, pool)
    if test:
        found["test"] = test

    if "ship" not in found or "test" not in found:
        ordered = sorted(pool, key=lambda x: x[0].x0)
        mid = [d for d in ordered if 250 < d[0].x0 < 420]
        right = [d for d in ordered if d[0].x0 >= 450]
        if "ship" not in found and mid:
            found["ship"] = mid[0]
        if "test" not in found and right:
            found["test"] = right[0]
        if ("ship" not in found or "test" not in found) and len(ordered) >= 2:
            best_pair = None
            best_dy = 1e9
            for i in range(len(ordered)):
                for j in range(i + 1, len(ordered)):
                    a, b = ordered[i], ordered[j]
                    dy = abs(a[0].y0 - b[0].y0)
                    if dy < best_dy and dy < 20:
                        best_dy = dy
                        best_pair = (a, b) if a[0].x0 <= b[0].x0 else (b, a)
            if best_pair:
                if "ship" not in found:
                    found["ship"] = best_pair[0]
                if "test" not in found:
                    found["test"] = best_pair[1]
    return found


# 采购订单号：纯数字 6~14 位，或字母开头订单号。排除 13 位物料编码。
_PO_VALUE_RE = re.compile(r"([A-Za-z][A-Za-z0-9]{4,13}|\d{6,14})")


def _find_po_number(page: fitz.Page) -> tuple[fitz.Rect, float, str] | None:
    label = _find_label_rect(page, ["订单编号"])
    if not label:
        label = _find_label_rect(page, ["PO."])

    candidates: list[tuple[fitz.Rect, float, str, float]] = []
    data = page.get_text("dict")
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            parts: list[tuple[str, fitz.Rect, float]] = []
            for span in line.get("spans", []):
                text = span.get("text") or ""
                if not text:
                    continue
                parts.append((text, fitz.Rect(span["bbox"]), float(span.get("size") or 9)))
            if not parts:
                continue
            joined = "".join(t for t, _, _ in parts)
            owner: list[int] = []
            for i, (t, _, _) in enumerate(parts):
                owner.extend([i] * len(t))
            for m in _PO_VALUE_RE.finditer(joined):
                token = m.group(1)
                if "-" in token:
                    continue
                if re.fullmatch(r"20\d{2}", token):
                    continue
                if re.fullmatch(r"3\d{12}", token):
                    continue
                idxs = sorted(set(owner[m.start() : m.end()]))
                rects = [parts[i][1] for i in idxs]
                union = rects[0]
                for r in rects[1:]:
                    union |= r
                fs = parts[idxs[0]][2]
                if label:
                    if abs(union.y0 - label.y0) > 40:
                        continue
                    if union.x0 < label.x0 - 20:
                        continue
                    if union.x0 - label.x0 > 280:
                        continue
                    score = abs(union.y0 - label.y0) * 1000 + abs(union.x0 - label.x0)
                else:
                    if union.y0 > 200:
                        continue
                    score = union.y0 * 10 + union.x0
                candidates.append((union, fs, token, score))

    if not candidates:
        if label:
            rect = fitz.Rect(label.x1 + 8, label.y0 - 1, label.x1 + 100, label.y1 + 8)
            return rect, 9.5, ""
        return None

    candidates.sort(key=lambda x: x[3])
    rect, fs, text, _ = candidates[0]

    # 把同一行、紧邻的数字/空格碎片并入（清掉「460 460xxx」左侧残留）
    ship_label = _find_label_rect(page, ["出货日期"])
    x_right = ship_label.x0 - 2 if ship_label else rect.x1 + 80
    for span, r in _iter_spans(page):
        t = span["text"]
        if not t or not t.strip():
            continue
        if abs(r.y0 - rect.y0) > 10:
            continue
        if label is not None and r.x1 < label.x1:
            continue
        if r.x0 > x_right:
            continue
        # 仅合并数字、空格、短碎片
        if not re.fullmatch(r"[\d\s]{1,14}", t.strip()):
            continue
        # 与主矩形水平相邻或重叠
        if r.x0 > rect.x1 + 12:
            continue
        if r.x1 < rect.x0 - 12:
            continue
        rect |= r

    return rect, fs, text


# ---------------------------------------------------------------------------
# 尺寸表 Sample1~10：与旧「出货报告工具」同一套思路，并修复拆分小数识别
# ---------------------------------------------------------------------------


def _merge_split_decimals(
    value_spans: list[tuple[fitz.Rect, float, str]],
) -> list[tuple[fitz.Rect, float, str]]:
    """
    合并被拆开的小数：
    - 21.6 + 6 → 21.66
    - 14 + .18 → 14.18
    - 9 + . + 72 → 9.72
    """
    value_spans = sorted(value_spans, key=lambda x: x[0].x0)
    merged: list[tuple[fitz.Rect, float, str]] = []
    i = 0
    n = len(value_spans)
    while i < n:
        rect, fs, text = value_spans[i]
        t = (text or "").strip()

        # 14 + .18
        if (
            i + 1 < n
            and re.fullmatch(r"\d{1,4}", t)
            and re.fullmatch(r"\.\d+", (value_spans[i + 1][2] or "").strip())
            and value_spans[i + 1][0].x0 - rect.x1 < 10
        ):
            r2, fs2, t2 = value_spans[i + 1]
            union = fitz.Rect(rect.x0, min(rect.y0, r2.y0), r2.x1, max(rect.y1, r2.y1))
            merged.append((union, fs or fs2, t + t2.strip()))
            i += 2
            continue

        # 9 + . + 72
        if (
            i + 2 < n
            and re.fullmatch(r"\d{1,4}", t)
            and (value_spans[i + 1][2] or "").strip() == "."
            and re.fullmatch(r"\d{1,4}", (value_spans[i + 2][2] or "").strip())
            and value_spans[i + 1][0].x0 - rect.x1 < 8
            and value_spans[i + 2][0].x0 - value_spans[i + 1][0].x1 < 8
        ):
            r2, _, _ = value_spans[i + 1]
            r3, fs3, t3 = value_spans[i + 2]
            union = fitz.Rect(
                rect.x0,
                min(rect.y0, r2.y0, r3.y0),
                r3.x1,
                max(rect.y1, r2.y1, r3.y1),
            )
            merged.append((union, fs or fs3, t + "." + t3.strip()))
            i += 3
            continue

        # 21.6 + 6
        if (
            i + 1 < n
            and len(t) <= 6
            and "." in t
            and re.fullmatch(r"\d", (value_spans[i + 1][2] or "").strip())
            and value_spans[i + 1][0].x0 - rect.x1 < 8
        ):
            r2, _, t2 = value_spans[i + 1]
            union = fitz.Rect(rect.x0, min(rect.y0, r2.y0), r2.x1, max(rect.y1, r2.y1))
            merged.append((union, fs, t + t2.strip()))
            i += 2
            continue

        # 单独的 "." / ".18" 不应作为样本留下
        if t == "." or re.fullmatch(r"\.\d+", t):
            i += 1
            continue

        merged.append((rect, fs, t))
        i += 1
    return merged


def _cluster_spans_by_row(
    items: list[tuple[fitz.Rect, float, str]],
    dy: float = 7.0,
) -> list[list[tuple[fitz.Rect, float, str]]]:
    """把纵坐标接近的文字归为同一数据行。"""
    if not items:
        return []
    items = sorted(items, key=lambda x: (x[0].y0, x[0].x0))
    rows: list[list[tuple[fitz.Rect, float, str]]] = [[items[0]]]
    for item in items[1:]:
        ys = [r.y0 for r, _, _ in rows[-1]]
        ref_y = sum(ys) / len(ys)
        if abs(item[0].y0 - ref_y) <= dy:
            rows[-1].append(item)
        else:
            rows.append([item])
    return rows


def _strip_leading_row_numbers(
    row: list[tuple[fitz.Rect, float, str]],
) -> list[tuple[fitz.Rect, float, str]]:
    """去掉尺寸表最左「序号」列，避免拼进公差。"""
    body = sorted(row, key=lambda x: x[0].x0)
    while body:
        rect, _fs, text = body[0]
        raw = (text or "").strip()
        if not re.fullmatch(r"\d{1,2}", raw):
            break
        try:
            n = int(raw)
        except ValueError:
            break
        if n < 1 or n > 40:
            break

        if len(body) >= 2:
            nrect, _, ntext = body[1]
            ntext = (ntext or "").strip()
            gap = nrect.x0 - rect.x1
            if gap < 10 and re.match(r"^[.±+/]|mm", ntext, re.I):
                break
            if gap < 6 and re.match(r"^\d", ntext):
                break

        if rect.x0 >= 48:
            break
        body.pop(0)
    return body


def _spec_column_right(row: list[tuple[fitz.Rect, float, str]]) -> float:
    """尺寸/公差列右边界：样本数字必须在此右侧。"""
    right = 120.0
    for r, _, t in row:
        raw = (t or "").strip()
        if not raw:
            continue
        if any(ch in raw for ch in ("±", "+/", "＋")) or "mm" in raw.lower():
            right = max(right, r.x1)
        # 「14mm±0.3」整段
        if parse_all_tolerances(raw):
            right = max(right, r.x1)
    return right + 2


def _resolve_row_tolerance(
    row: list[tuple[fitz.Rect, float, str]],
    sample_values: list[float],
):
    """从一行文字里稳健解析公差，并用原样本校验。"""
    body = _strip_leading_row_numbers(row)
    if not body:
        return None
    candidates = []

    sample_xs = [
        r.x0
        for r, _, t in body
        if re.fullmatch(r"\d+(?:\.\d+)?", (t or "").strip()) and r.x0 > 100
    ]
    sample_x = min(sample_xs) if sample_xs else 130.0
    spec_right = max(120.0, sample_x - 5, _spec_column_right(body) - 2)

    spec_col = "".join(t for r, _, t in body if r.x0 < spec_right)
    candidates.extend(parse_all_tolerances(spec_col))

    full = "".join(t for _, _, t in body)
    candidates.extend(parse_all_tolerances(full))

    uniq = []
    seen = set()
    for t in candidates:
        key = (round(t.nominal, 6), round(t.low, 6), round(t.high, 6))
        if key in seen:
            continue
        seen.add(key)
        uniq.append(t)

    return pick_tolerance_for_samples(uniq, sample_values)


def _is_sample_header_row(merged: list[tuple[fitz.Rect, float, str]]) -> bool:
    """表头 Sample1~10 的纯序号行，不能当数据改。"""
    if len(merged) < 5:
        return False
    vals = []
    for _, _, raw in merged:
        if not re.fullmatch(r"\d{1,2}", raw):
            return False
        vals.append(int(raw))
    # 1,2,3... 连续
    return vals == list(range(vals[0], vals[0] + len(vals)))


def _replace_size_measurements(page: fitz.Page) -> int:
    """
    替换尺寸表 Sample1~10 数值。
    与旧出货报告工具相同思路：公差内小幅波动；并正确合并拆分小数。
    """
    y_lo, y_hi = 400, 660
    items: list[tuple[fitz.Rect, float, str]] = []
    for span, rect in _iter_spans(page):
        if rect.y0 < y_lo or rect.y0 > y_hi:
            continue
        text = span["text"]
        if text is None:
            continue
        items.append((rect, _span_fontsize(span), text))

    rows = _cluster_spans_by_row(items)
    replaced_rows = 0

    for row in rows:
        row = sorted(row, key=lambda x: x[0].x0)
        joined = "".join(t for _, _, t in row)
        if any(
            k in joined
            for k in ("Sample", "判定", "二次元", "性能", "指标", "Appearance", "Texture", "No.")
        ):
            continue

        spec_right = _spec_column_right(row)

        # 收集候选：完整小数、整数、以及 ".18" / "." 碎片（供合并）
        value_spans: list[tuple[fitz.Rect, float, str]] = []
        for rect, fs, text in row:
            raw = (text or "").strip()
            if not raw or raw.upper() == "OK":
                continue
            if rect.x0 < spec_right:
                continue
            if not (
                re.fullmatch(r"\d+(?:\.\d+)?", raw)
                or re.fullmatch(r"\.\d+", raw)
                or raw == "."
            ):
                continue
            value_spans.append((rect, fs, raw))

        merged = _merge_split_decimals(value_spans)
        # 合并后只保留真正的数值
        merged = [
            (r, fs, t)
            for r, fs, t in merged
            if re.fullmatch(r"\d+(?:\.\d+)?", t)
        ]
        if len(merged) < 5:
            continue
        if _is_sample_header_row(merged):
            continue
        merged = merged[:10]

        old_vals: list[float] = []
        old_decs: list[int] = []
        for _, _, raw in merged:
            try:
                old_vals.append(float(raw))
                if "." in raw:
                    old_decs.append(len(raw.split(".", 1)[1]))
                else:
                    old_decs.append(0)
            except ValueError:
                pass

        tol = _resolve_row_tolerance(row, old_vals)
        if not tol:
            continue

        if old_vals:
            med = sorted(old_vals)[len(old_vals) // 2]
            pad = max((tol.high - tol.low) * 3, abs(tol.nominal) * 0.15, 0.2)
            if not (tol.low - pad <= med <= tol.high + pad):
                continue

        # 小数位：优先跟原样本一致（如 0.571 三位），否则跟公差
        prefer_dec = max(old_decs) if old_decs else tol.decimals
        prefer_dec = max(prefer_dec, tol.decimals)

        new_vals = generate_samples_smart(tol, count=len(merged))
        safe_vals = []
        for idx, v in enumerate(new_vals):
            fv = float(v)
            fv = min(max(fv, tol.low), tol.high)
            # 该格原小数位优先
            if idx < len(old_decs) and old_decs[idx] > 0:
                dec = old_decs[idx]
            elif "." in v:
                dec = max(len(v.split(".", 1)[1]), prefer_dec)
            else:
                dec = max(prefer_dec, 2)
            safe_vals.append(f"{fv:.{dec}f}")

        # 批量擦除再写入，减少残影
        for rect, fs, _old in merged:
            pad_r = fitz.Rect(rect.x0 - 1.0, rect.y0 - 1.0, rect.x1 + 2.0, rect.y1 + 1.0)
            page.add_redact_annot(pad_r, fill=(1, 1, 1))
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        font_size = merged[0][1]
        for (rect, fs, _old), new_v in zip(merged, safe_vals):
            # 白块再盖，防拆分小数残影（如留下 18 / 21）
            cover = fitz.Rect(rect.x0 - 1.0, rect.y0 - 0.8, rect.x1 + 2.0, rect.y1 + 0.8)
            shape = page.new_shape()
            shape.draw_rect(cover)
            shape.finish(color=(1, 1, 1), fill=(1, 1, 1), width=0)
            shape.commit()
            baseline = fitz.Point(rect.x0, rect.y1 - 1.0)
            page.insert_text(
                baseline,
                new_v,
                fontsize=fs or font_size,
                fontname="helv",
                color=(0, 0, 0),
            )
        replaced_rows += 1

    return replaced_rows


def generate_report(req: GenerateRequest) -> dict:
    """改日期 / 订单编号 / 订单·交货数量 / 尺寸样本；其它照模板。"""
    ship = req.ship_date or date.today()
    ship_s = _fmt_date(ship)

    doc = fitz.open(req.template_path)
    try:
        page = doc[0]
        notes: list[str] = []
        failures: list[str] = []
        size_rows = 0
        jobs: list[tuple[fitz.Rect, str, float, float | None, float, float]] = []

        # 标签（用于擦除上限）
        deliver_label = _find_label_rect(page, ["交货数量"])
        batch_label = _find_label_rect(page, ["制造批号"])
        ship_label = _find_label_rect(page, ["出货日期"])
        test_label = _find_label_rect(page, ["检验日期"])
        if not test_label:
            test_label = _find_label_rect(page, ["QC Date"])

        dates = _find_dates(page)
        po_hit = _find_po_number(page)
        qty_pair = _find_qty_pair(page)

        # 1) 日期：左侧少扩，保护「Date：」冒号
        if "header_date_full" in dates:
            rect, fs, old = dates["header_date_full"]
            jobs.append((rect, ship_s, fs, None, 0.6, 3.0))
            notes.append(f"页眉日期 {old} → {ship_s}")
        if "ship" in dates:
            rect, fs, old = dates["ship"]
            lim = test_label.x0 if test_label else None
            jobs.append((rect, ship_s, fs, lim, 1.0, 3.0))
            notes.append(f"出货日期 {old} → {ship_s}")
        else:
            failures.append("未定位到出货日期")
        if "test" in dates:
            rect, fs, old = dates["test"]
            jobs.append((rect, ship_s, fs, None, 1.0, 3.0))
            notes.append(f"检验日期 {old} → {ship_s}")
        else:
            failures.append("未定位到检验日期")

        # 2) 订单编号（始终重写；左侧多扩清「460 」残留）
        if req.po_no:
            if po_hit:
                rect, fs, old = po_hit
                lim = ship_label.x0 if ship_label else None
                jobs.append((rect, str(req.po_no), fs, lim, 4.0, 5.0))
                notes.append(f"订单号 {old or '(空)'} → {req.po_no}")
            else:
                failures.append("未定位到订单号")
        else:
            notes.append("Excel 无采购订单号，跳过订单号改写")

        # 3) 订单数量 / 交货数量（右侧多扩清 PCS/CS 单位）
        if qty_pair:
            (r1, fs1, o1, u1), (r2, fs2, o2, u2) = qty_pair
            if u2 == "PCS" and re.search(r"Pcs|pcs", o2) and "PCS" not in o2:
                u2 = "Pcs"
            new1 = _format_qty(req.order_qty, u1)
            new2 = _format_qty(req.deliver_qty, u2)
            lim1 = deliver_label.x0 if deliver_label else None
            lim2 = batch_label.x0 if batch_label else None
            if lim1 is not None and r1.x1 > lim1 - 2:
                r1 = fitz.Rect(r1.x0, r1.y0, lim1 - 2, r1.y1)
            if lim2 is not None and r2.x1 > lim2 - 2:
                r2 = fitz.Rect(r2.x0, r2.y0, lim2 - 2, r2.y1)
            jobs.append((r1, new1, fs1, lim1, 2.5, 5.0))
            jobs.append((r2, new2, fs2, lim2, 2.5, 5.0))
            notes.append(f"订单数量 {o1} → {new1}")
            notes.append(f"交货数量 {o2} → {new2}")
        else:
            failures.append("未定位到订单/交货数量")

        if failures:
            raise RuntimeError("关键字段未能改写：" + "；".join(failures))

        label_guards = [
            ("订单数量", _find_label_rect(page, ["订单数量"])),
            ("交货数量", deliver_label),
            ("出货日期", ship_label),
            ("检验日期", test_label),
            ("订单编号", _find_label_rect(page, ["订单编号"])),
        ]
        for old_rect, new_text, fs, x1_limit, pad_left, pad_right in jobs:
            erase = _erase_rect_for_old_value(
                old_rect, x1_limit=x1_limit, pad_left=pad_left, pad_right=pad_right
            )
            for name, lab in label_guards:
                if not lab:
                    continue
                inter = erase & lab
                if inter.get_area() > 8:
                    raise RuntimeError(f"擦除范围会破坏标签「{name}」，已中止以免生成坏文件")

        _batch_replace(page, jobs)

        # 4) 尺寸 Sample1~10（与旧工具相同算法）
        size_rows = _replace_size_measurements(page)
        notes.append(f"尺寸样本已重新生成：{size_rows} 行")
        if size_rows < 1:
            notes.append("提示：未识别到尺寸样本行（若模板无尺寸表可忽略）")

        req.output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(req.output_path, garbage=4, deflate=True)
    finally:
        doc.close()

    return {
        "output": str(req.output_path),
        "ship_date": ship_s,
        "notes": notes,
        "size_rows": size_rows,
    }
