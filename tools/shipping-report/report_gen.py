"""基于首份出货报告模板，改写字段并生成新 PDF。"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path

import fitz

from aql import AqlResult, lookup_aql
from size_gen import (
    generate_samples_smart,
    parse_all_tolerances,
    parse_tolerance,
    pick_tolerance_for_samples,
)


@dataclass
class GenerateRequest:
    template_path: Path
    output_path: Path
    po_no: str
    order_qty: int
    deliver_qty: int
    ship_date: date | None = None


def _fmt_date(d: date) -> str:
    # 与现有报告一致：2026/7/29（月日不补零）
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


def _near(a: fitz.Rect, b: fitz.Rect, max_dx: float = 250, max_dy: float = 25) -> bool:
    return abs(a.y0 - b.y0) <= max_dy and 0 <= (b.x0 - a.x0) <= max_dx


def _find_label_rect(page: fitz.Page, labels: list[str]) -> fitz.Rect | None:
    for span, rect in _iter_spans(page):
        t = span["text"].strip()
        for lab in labels:
            if lab in t:
                return rect
    return None


def _replace_text(page: fitz.Page, rect: fitz.Rect, new_text: str, fontsize: float) -> None:
    # 擦除宽度取「原文字」与「新文字估算」的较大值，避免残留或遮挡
    est_w = max(rect.width, len(new_text) * fontsize * 0.58)
    pad = fitz.Rect(rect.x0 - 1, rect.y0 - 1, rect.x0 + est_w + 3, rect.y1 + 1)
    page.add_redact_annot(pad, fill=(1, 1, 1))
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    baseline = fitz.Point(rect.x0, rect.y1 - 1.2)
    page.insert_text(
        baseline,
        new_text,
        fontsize=fontsize,
        fontname="helv",
        color=(0, 0, 0),
    )


def _find_value_after_label(
    page: fitz.Page,
    labels: list[str],
    value_pattern: re.Pattern,
    *,
    prefer_right: bool = True,
    max_dx: float = 280,
    max_dy: float = 28,
) -> tuple[fitz.Rect, float, str] | None:
    label_rect = _find_label_rect(page, labels)
    if not label_rect:
        return None
    candidates: list[tuple[fitz.Rect, float, str, float]] = []
    for span, rect in _iter_spans(page):
        text = span["text"].strip()
        if not value_pattern.fullmatch(text.strip()):
            # 允许末尾空格
            if not value_pattern.fullmatch(text):
                continue
        if abs(rect.y0 - label_rect.y0) > max_dy:
            continue
        if prefer_right and rect.x0 < label_rect.x0 - 5:
            continue
        dx = rect.x0 - label_rect.x0
        if dx > max_dx:
            continue
        candidates.append((rect, _span_fontsize(span), text, dx))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[3])
    r, fs, text, _ = candidates[0]
    return r, fs, text


_QTY_TEXT_RE = re.compile(
    r"^(?P<num>\d{1,8})\s*(?:pcs|PCS|Pcs|个|EA|ea)?\.?$",
)


def _parse_qty_span(text: str) -> str | None:
    """识别数量文字：4000 / 60000pcs / 60000 pcs 等。"""
    t = (text or "").strip()
    m = _QTY_TEXT_RE.fullmatch(t)
    return m.group("num") if m else None


def _find_qty_near_label(
    page: fitz.Page,
    label: fitz.Rect,
    *,
    max_dy: float = 36,
    max_dx: float = 220,
    x_min: float | None = None,
    x_max: float | None = None,
) -> tuple[fitz.Rect, float, str] | None:
    """在标签右侧/附近找数量值（兼容有无表头黑框、带 pcs 单位）。"""
    best = None
    best_score = 1e18
    for span, rect in _iter_spans(page):
        raw = span["text"].strip()
        if _parse_qty_span(raw) is None:
            continue
        if abs(rect.y0 - label.y0) > max_dy:
            continue
        # 数量一般在标签右侧；允许略重叠
        if rect.x0 < label.x0 - 10:
            continue
        dx = rect.x0 - label.x0
        if dx > max_dx:
            continue
        if x_min is not None and rect.x0 < x_min:
            continue
        if x_max is not None and rect.x0 > x_max:
            continue
        # 优先更近、更靠右一点的
        score = abs(rect.y0 - label.y0) * 1000 + dx
        if score < best_score:
            best_score = score
            best = (rect, _span_fontsize(span), raw)
    return best


def _find_qty_pair(page: fitz.Page) -> tuple[tuple[fitz.Rect, float, str], tuple[fitz.Rect, float, str]] | None:
    """定位订单数量、交货数量（兼容纯数字或 60000pcs，以及版式略偏移）。"""
    order_label = _find_label_rect(page, ["订单数量", "Order quantit"])
    deliver_label = _find_label_rect(page, ["交货数量", "Delivered"])
    if not order_label or not deliver_label:
        return None

    # 各自在标签旁找，避免被批号区或其它数字干扰
    mid_x = (order_label.x0 + deliver_label.x0) / 2
    order_hit = _find_qty_near_label(
        page,
        order_label,
        x_max=max(mid_x + 40, order_label.x0 + 200),
    )
    deliver_hit = _find_qty_near_label(
        page,
        deliver_label,
        x_min=min(mid_x - 20, deliver_label.x0 - 30),
        x_max=500,
    )

    if order_hit and deliver_hit:
        return order_hit, deliver_hit

    # 退化：同行取最左两个数量（仍兼容 pcs）
    nums: list[tuple[fitz.Rect, float, str]] = []
    for span, rect in _iter_spans(page):
        raw = span["text"].strip()
        if _parse_qty_span(raw) is None:
            continue
        if abs(rect.y0 - order_label.y0) > 36 and abs(rect.y0 - deliver_label.y0) > 36:
            continue
        if rect.x0 > 500:
            continue
        nums.append((rect, _span_fontsize(span), raw))
    if len(nums) < 2:
        return None
    nums.sort(key=lambda x: x[0].x0)
    return nums[0], nums[1]


def _find_sample_letter(page: fitz.Page) -> tuple[fitz.Rect, float, str] | None:
    label = _find_label_rect(page, ["样本代字", "Sample substitute"])
    if not label:
        return None
    best = None
    best_dx = 1e9
    for span, rect in _iter_spans(page):
        text = span["text"].strip()
        if not re.fullmatch(r"[A-RT-Z]", text):  # 排除 S 等
            continue
        if abs(rect.y0 - label.y0) > 20:
            continue
        if rect.x0 <= label.x0:
            continue
        dx = rect.x0 - label.x0
        if dx < best_dx and dx < 200:
            best_dx = dx
            best = (rect, _span_fontsize(span), text)
    return best


def _find_sample_size(page: fitz.Page) -> tuple[fitz.Rect, float, str] | None:
    label = _find_label_rect(page, ["样本数", "sample numbe"])
    if not label:
        return None
    best = None
    best_dx = 1e9
    for span, rect in _iter_spans(page):
        text = span["text"].strip()
        if not re.fullmatch(r"\d{1,4}", text):
            continue
        if abs(rect.y0 - label.y0) > 22:
            continue
        if rect.x0 <= label.x0:
            continue
        dx = rect.x0 - label.x0
        if dx < best_dx and dx < 180:
            best_dx = dx
            best = (rect, _span_fontsize(span), text)
    return best


def _find_ac_re(page: fitz.Page) -> dict[str, tuple[fitz.Rect, float, str]]:
    """
    返回 maj_ac, maj_re, min_ac, min_re。
    注意：尺寸表里也可能出现 0.65 等数字，必须限定在 AQL 区域（页面上半）。
    """
    result: dict[str, tuple[fitz.Rect, float, str]] = {}

    # AQL 区大约在 y<400；尺寸实测更靠下
    aql_y_max = 400.0

    y_maj = None
    y_min = None
    for span, rect in _iter_spans(page):
        if rect.y0 > aql_y_max:
            continue
        t = span["text"].strip()
        if t == "0.65":
            # 取最靠上的（真正的 AQL 值）
            if y_maj is None or rect.y0 < y_maj:
                y_maj = rect.y0
        if t.rstrip() == "1.0":
            if y_min is None or rect.y0 < y_min:
                y_min = rect.y0

    if y_maj is None:
        maj_label = _find_label_rect(page, ["MAJ缺陷", "MAJ defect", "MAJ"])
        if maj_label and maj_label.y0 < aql_y_max:
            y_maj = maj_label.y0
    if y_min is None:
        min_label = _find_label_rect(page, ["MIN缺陷", "MIN defect", "MIN defec"])
        if min_label and min_label.y0 < aql_y_max:
            y_min = min_label.y0

    def collect_for_y(y: float) -> list[tuple[fitz.Rect, float, str]]:
        items = []
        for span, rect in _iter_spans(page):
            text = span["text"].strip()
            if not re.fullmatch(r"\d{1,2}", text):
                continue
            if abs(rect.y0 - y) > 14:
                continue
            # 允收约 x=380，拒收约 x=540（略放宽）
            if 340 <= rect.x0 <= 430 or 500 <= rect.x0 <= 590:
                items.append((rect, _span_fontsize(span), text))
        items.sort(key=lambda x: x[0].x0)
        return items

    def fill_from_labels(prefix: str, y_hint: float | None) -> None:
        """按同行「允收数/拒收数」标签右侧数字兜底。"""
        ac_key, re_key = f"{prefix}_ac", f"{prefix}_re"
        if ac_key in result and re_key in result:
            return
        ac_labs: list[fitz.Rect] = []
        re_labs: list[fitz.Rect] = []
        for span, rect in _iter_spans(page):
            if rect.y0 > aql_y_max:
                continue
            t = span["text"].strip()
            if t.startswith("允收数") or t.startswith("Acceptance"):
                if y_hint is None or abs(rect.y0 - y_hint) < 25:
                    ac_labs.append(rect)
            if t.startswith("拒收数") or t.startswith("Rejection"):
                if y_hint is None or abs(rect.y0 - y_hint) < 25:
                    re_labs.append(rect)

        def nearest_digit(lab: fitz.Rect) -> tuple[fitz.Rect, float, str] | None:
            best = None
            best_dx = 1e9
            for span, rect in _iter_spans(page):
                text = span["text"].strip()
                if not re.fullmatch(r"\d{1,2}", text):
                    continue
                if abs(rect.y0 - lab.y0) > 16:
                    continue
                if rect.x0 < lab.x0:
                    continue
                dx = rect.x0 - lab.x0
                if dx < best_dx and dx < 220:
                    best_dx = dx
                    best = (rect, _span_fontsize(span), text)
            return best

        # 同一 AQL 行上可能有多个「允收数」标签（MAJ/MIN 各一）；按 y 与 hint 最近的选
        if ac_key not in result and ac_labs:
            ac_labs.sort(key=lambda r: abs(r.y0 - (y_hint or r.y0)))
            hit = nearest_digit(ac_labs[0])
            if hit:
                result[ac_key] = hit
        if re_key not in result and re_labs:
            re_labs.sort(key=lambda r: abs(r.y0 - (y_hint or r.y0)))
            hit = nearest_digit(re_labs[0])
            if hit:
                result[re_key] = hit

    if y_maj is not None:
        maj_nums = collect_for_y(y_maj)
        if len(maj_nums) >= 2:
            result["maj_ac"] = maj_nums[0]
            result["maj_re"] = maj_nums[1]
        elif len(maj_nums) == 1:
            result["maj_ac"] = maj_nums[0]
        fill_from_labels("maj", y_maj)

    if y_min is not None:
        min_nums = collect_for_y(y_min)
        if len(min_nums) >= 2:
            result["min_ac"] = min_nums[0]
            result["min_re"] = min_nums[1]
        elif len(min_nums) == 1:
            result["min_ac"] = min_nums[0]
        fill_from_labels("min", y_min)

    return result


_DATE_RE = re.compile(r"20\d{2}/\d{1,2}/\d{1,2}")


def _find_date_instances(page: fitz.Page) -> list[tuple[fitz.Rect, float, str, str]]:
    """
    找出页面上所有日期。
    有的模板把 2026/7/27 拆成多个 span，需要按行拼接后再匹配。
    返回 (覆盖矩形, 字号, 日期文本, 种类 header|plain)
    """
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
            # 字符 -> span 下标
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
                    # 页眉 Date：xxxx 整段一起换
                    kind = "header"
                    # 从 Date 起算到日期结束
                    dpos = prefix.rfind("Date")
                    if dpos >= 0:
                        idxs2 = sorted(set(owner[dpos : m.end()]))
                        rects2 = [parts[i][1] for i in idxs2]
                        union = rects2[0]
                        for r in rects2[1:]:
                            union |= r
                        fs = parts[idxs2[0]][2]
                found.append((union, fs, m.group(0), kind))
    return found


def _find_dates(page: fitz.Page) -> dict[str, tuple[fitz.Rect, float, str]]:
    """出货日期、检验日期、页眉 Date。兼容日期被拆成多个文字块的模板。"""
    found: dict[str, tuple[fitz.Rect, float, str]] = {}
    ship_label = _find_label_rect(page, ["出货日期", "Shipment"])
    test_label = _find_label_rect(page, ["检验日期", "Test dat"])

    instances = _find_date_instances(page)
    headers = [(r, fs, t) for r, fs, t, k in instances if k == "header"]
    plains = [(r, fs, t) for r, fs, t, k in instances if k == "plain"]

    if headers:
        # 取最靠上的页眉日期
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
            # 日期通常在标签右侧；允许少量重叠
            if rect.x0 < label.x0 - 50:
                continue
            score = abs(rect.y0 - label.y0) * 1000 + abs(rect.x0 - label.x0)
            if score < best_score:
                best_score = score
                best = (rect, fs, text)
        return best

    # 优先用 plain；若没有，也可用全部实例（去掉已用作 header 的）
    pool = plains if plains else [(r, fs, t) for r, fs, t, _k in instances]
    ship = nearest(ship_label, pool)
    if ship:
        found["ship"] = ship
    test = nearest(test_label, pool)
    if test:
        found["test"] = test

    # 按水平位置兜底：中间=出货，右侧=检验
    if "ship" not in found or "test" not in found:
        ordered = sorted(pool, key=lambda x: x[0].x0)
        mid = [d for d in ordered if 250 < d[0].x0 < 420]
        right = [d for d in ordered if d[0].x0 >= 450]
        if "ship" not in found and mid:
            found["ship"] = mid[0]
        if "test" not in found and right:
            found["test"] = right[0]
        # 再不行：同行两个 plain 日期按左右分配
        if ("ship" not in found or "test" not in found) and len(ordered) >= 2:
            # 取 y 接近的一对
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


_PO_VALUE_RE = re.compile(r"([A-Za-z][A-Za-z0-9]{4,13})")


def _find_po_number(page: fitz.Page) -> tuple[fitz.Rect, float, str] | None:
    """定位报告上的订单号（PAQ/PDQ/P9Q 等任意字母数字组合；优先在订单编号旁）。"""
    label = _find_label_rect(page, ["订单编号", "PO."])
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
                # 跳过带连字符的料号碎片（一般不会进到这里）
                if "-" in token:
                    continue
                idxs = sorted(set(owner[m.start() : m.end()]))
                rects = [parts[i][1] for i in idxs]
                union = rects[0]
                for r in rects[1:]:
                    union |= r
                fs = parts[idxs[0]][2]
                score = 0.0
                if label:
                    if abs(union.y0 - label.y0) > 45:
                        continue
                    if union.x0 < label.x0 - 30:
                        continue
                    score = abs(union.y0 - label.y0) * 1000 + abs(union.x0 - label.x0)
                else:
                    if union.y0 > 250:
                        continue
                    score = union.y0 * 10 + union.x0
                candidates.append((union, fs, token.upper(), score))

    if candidates:
        candidates.sort(key=lambda x: x[3])
        rect, fs, text, _ = candidates[0]
        return rect, fs, text

    # 模板订单号为空或只有 / 时：在标签右侧开辟写入区
    if label:
        rect = fitz.Rect(label.x1 + 8, label.y0 - 1, label.x1 + 100, label.y1 + 10)
        return rect, 9.5, ""
    return None


_SPEC_LINE_RE = re.compile(
    r"(?P<spec>\d+(?:\.\d+)?\s*(?:mm)?\s*(?:±\s*\d+(?:\.\d+)?|\+\s*\d+(?:\.\d+)?\s*/\s*-\s*\d+(?:\.\d+)?))",
    re.I,
)


def _merge_split_decimals(
    value_spans: list[tuple[fitz.Rect, float, str]],
) -> list[tuple[fitz.Rect, float, str]]:
    """合并被拆开的小数，如 21.6 + 6 → 21.66。"""
    value_spans = sorted(value_spans, key=lambda x: x[0].x0)
    merged: list[tuple[fitz.Rect, float, str]] = []
    i = 0
    while i < len(value_spans):
        rect, fs, text = value_spans[i]
        if (
            i + 1 < len(value_spans)
            and len(text) <= 4
            and "." in text
            and re.fullmatch(r"\d", value_spans[i + 1][2])
            and value_spans[i + 1][0].x0 - rect.x1 < 8
        ):
            r2, _, t2 = value_spans[i + 1]
            union = fitz.Rect(rect.x0, min(rect.y0, r2.y0), r2.x1, max(rect.y1, r2.y1))
            merged.append((union, fs, text + t2))
            i += 2
        else:
            merged.append((rect, fs, text))
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
    """
    去掉尺寸表最左「序号」列（1、2、3…），避免拼进公差。
    典型序号在 x≈29；尺寸标称（如 4±0.3）更靠右，不能误删。
    """
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
            # 紧挨着 ± / + / 小数点 / mm → 这是尺寸标称，不是序号
            if gap < 10 and re.match(r"^[.±+/]|mm", ntext, re.I):
                break
            # 数字被拆成两段紧挨着（如 1 + 0.35），也不当序号剥
            if gap < 6 and re.match(r"^\d", ntext):
                break

        # 所有已知模板序号约在 x=28~30；再靠右的纯数字优先当尺寸
        if rect.x0 >= 48:
            break
        body.pop(0)
    return body


def _resolve_row_tolerance(
    row: list[tuple[fitz.Rect, float, str]],
    sample_values: list[float],
):
    """
    从一行文字里稳健解析公差：
    - 绝不把最左序号列拼进公差
    - 优先用「尺寸列」区域文字
    - 用现有样本中位数校验，避免 5+120→5120、1+0.35→10.35
    """
    body = _strip_leading_row_numbers(row)
    if not body:
        return None
    candidates = []

    # 尺寸列大致在样本列左侧；按相对位置取，不写死单一版式
    sample_xs = [
        r.x0
        for r, _, t in body
        if re.fullmatch(r"\d+(?:\.\d+)?", (t or "").strip()) and r.x0 > 100
    ]
    sample_x = min(sample_xs) if sample_xs else 130.0
    spec_right = max(120.0, sample_x - 5)

    # 候选1：尺寸列区域
    spec_col = "".join(t for r, _, t in body if r.x0 < spec_right)
    candidates.extend(parse_all_tolerances(spec_col))

    # 候选2：整行（已去序号）——公差被拆得很散时兜底
    full = "".join(t for _, _, t in body)
    candidates.extend(parse_all_tolerances(full))

    # 去重（按 nominal/low/high）
    uniq: list = []
    seen = set()
    for t in candidates:
        key = (round(t.nominal, 6), round(t.low, 6), round(t.high, 6))
        if key in seen:
            continue
        seen.add(key)
        uniq.append(t)

    return pick_tolerance_for_samples(uniq, sample_values)


def _replace_size_measurements(page: fitz.Page) -> int:
    """
    替换尺寸表 Sample1~10 数值。
    兼容公差拆段；并用原样本校验，防止序号拼进公差导致普遍错数。
    """
    y_lo, y_hi = 410, 640
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
        # 跳过表头 / 性能区
        if any(k in joined for k in ("Sample", "判定", "二次元", "性能", "指标", "Appearance", "Texture")):
            continue

        # 先收集右侧原样本值（用于校验公差是否读对）
        value_spans: list[tuple[fitz.Rect, float, str]] = []
        for rect, fs, text in row:
            raw = (text or "").strip()
            if not raw or raw.upper() == "OK":
                continue
            if rect.x0 < 115:
                continue
            if not re.fullmatch(r"\d+(?:\.\d+)?", raw):
                continue
            value_spans.append((rect, fs, raw))

        merged = _merge_split_decimals(value_spans)
        if len(merged) < 5:
            continue
        merged = merged[:10]

        old_vals: list[float] = []
        for _, _, raw in merged:
            try:
                old_vals.append(float(raw))
            except ValueError:
                pass

        tol = _resolve_row_tolerance(row, old_vals)
        if not tol:
            continue

        # 再保险：原样本中位数应大致靠近标称（防止仍读错）
        if old_vals:
            med = sorted(old_vals)[len(old_vals) // 2]
            pad = max((tol.high - tol.low) * 3, abs(tol.nominal) * 0.15, 0.2)
            if not (tol.low - pad <= med <= tol.high + pad):
                # 与标称差太大，本行跳过，避免写出离谱数据
                continue

        new_vals = generate_samples_smart(tol, count=len(merged))
        # 最终钳制，确保不越界
        safe_vals = []
        for v in new_vals:
            fv = float(v)
            fv = min(max(fv, tol.low), tol.high)
            # 保持原字符串小数位风格
            if "." in v:
                dec = len(v.split(".", 1)[1])
            else:
                dec = max(tol.decimals, 2)
            safe_vals.append(f"{fv:.{dec}f}")

        font_size = merged[0][1]
        for rect, fs, _old in merged:
            pad = fitz.Rect(rect.x0 - 0.8, rect.y0 - 0.8, rect.x1 + 1.5, rect.y1 + 0.8)
            page.add_redact_annot(pad, fill=(1, 1, 1))
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        for (rect, fs, _old), new_v in zip(merged, safe_vals):
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
    ship = req.ship_date or date.today()
    test = ship - timedelta(days=2)
    aql: AqlResult = lookup_aql(req.deliver_qty)

    doc = fitz.open(req.template_path)
    try:
        page = doc[0]
        notes: list[str] = []
        failures: list[str] = []

        # 1) 日期
        dates = _find_dates(page)
        ship_s = _fmt_date(ship)
        test_s = _fmt_date(test)
        if "header_date_full" in dates:
            rect, fs, old = dates["header_date_full"]
            _replace_text(page, rect, f"Date：{ship_s}", fs)
            notes.append(f"页眉日期 {old} → {ship_s}")
        if "ship" in dates:
            rect, fs, old = dates["ship"]
            _replace_text(page, rect, ship_s, fs)
            notes.append(f"出货日 {old} → {ship_s}")
        else:
            failures.append("未定位到出货日期")
        if "test" in dates:
            rect, fs, old = dates["test"]
            _replace_text(page, rect, test_s, fs)
            notes.append(f"检验日 {old} → {test_s}")
        else:
            failures.append("未定位到检验日期")

        # 2) 订单号（与当前采购订单对齐；支持 P9Q80157 等字母数字混合编号）
        po_hit = _find_po_number(page)
        if req.po_no:
            if po_hit:
                rect, fs, old = po_hit
                if (old or "").upper() != req.po_no.upper():
                    _replace_text(page, rect, req.po_no, fs)
                    notes.append(f"订单号 {old or '(空)'} → {req.po_no}")
                else:
                    notes.append(f"订单号保持 {old}")
            else:
                failures.append("未定位到订单号")

        # 3) 订单数量 / 交货数量
        qty_pair = _find_qty_pair(page)
        if qty_pair:
            (r1, fs1, o1), (r2, fs2, o2) = qty_pair
            _replace_text(page, r1, str(req.order_qty), fs1)
            _replace_text(page, r2, str(req.deliver_qty), fs2)
            notes.append(f"订单数量 {o1} → {req.order_qty}")
            notes.append(f"交货数量 {o2} → {req.deliver_qty}")
        else:
            failures.append("未定位到订单/交货数量")

        # 4) 样字 / 样本数
        letter_hit = _find_sample_letter(page)
        if letter_hit:
            rect, fs, old = letter_hit
            _replace_text(page, rect, aql.letter, fs)
            notes.append(f"样字 {old} → {aql.letter}")
        else:
            failures.append("未定位到样字")

        size_hit = _find_sample_size(page)
        if size_hit:
            rect, fs, old = size_hit
            wide = fitz.Rect(rect.x0, rect.y0, max(rect.x1, rect.x0 + 28), rect.y1)
            _replace_text(page, wide, str(aql.sample_size), fs)
            notes.append(f"样本数 {old} → {aql.sample_size}")
        else:
            failures.append("未定位到样本数")

        # 5) 允收 / 拒收
        ac_re = _find_ac_re(page)
        mapping = {
            "maj_ac": aql.maj_ac,
            "maj_re": aql.maj_re,
            "min_ac": aql.min_ac,
            "min_re": aql.min_re,
        }
        for key, new_v in mapping.items():
            if key in ac_re:
                rect, fs, old = ac_re[key]
                _replace_text(page, rect, str(new_v), fs)
                notes.append(f"{key} {old} → {new_v}")
            else:
                failures.append(f"未定位到{key}")

        # 6) 尺寸实测
        n_rows = _replace_size_measurements(page)
        notes.append(f"尺寸行已重新生成：{n_rows} 行")
        if n_rows < 1:
            failures.append("未改写尺寸实测数据")

        if failures:
            raise RuntimeError("关键字段未能改写：" + "；".join(failures))

        req.output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(req.output_path)
    finally:
        doc.close()

    return {
        "output": str(req.output_path),
        "aql": {
            "letter": aql.letter,
            "sample_size": aql.sample_size,
            "maj_ac": aql.maj_ac,
            "maj_re": aql.maj_re,
            "min_ac": aql.min_ac,
            "min_re": aql.min_re,
        },
        "ship_date": ship_s,
        "test_date": test_s,
        "notes": notes,
    }
