"""按物料编码匹配用户选择的出货报告 PDF 模板（只认文件名）。"""

from __future__ import annotations

import re
from pathlib import Path

# 送货单常见：纯数字料号（如 3228000410101）
_DIGIT_PART_RE = re.compile(r"(\d{10,20})")
# 旧式：221-0122001RU
_DASH_PART_RE = re.compile(r"(\d{3}-\d{7,}[A-Za-z0-9]*)")


def _clean_stem(stem: str) -> str:
    """去掉上传前缀 tpl_、末尾 (1)/(2) 等。"""
    s = re.sub(r"^tpl_(\d+_)?", "", stem, flags=re.I)
    return re.sub(r"\(\d+\)$", "", s).strip()


def detect_part_no_from_filename(
    pdf_path: str | Path,
    known_parts: list[str] | None = None,
) -> str | None:
    """
    只从文件名识别料号，不读 PDF 正文（避免正文旧客户货号误匹配）。
    文件名需等于或包含 Excel 里的物料编码，例如：
      3228000410101.pdf
      3228000410101(1).pdf
    """
    path = Path(pdf_path)
    stem = _clean_stem(path.stem)
    stem_up = stem.upper()
    known_upper = {p.upper(): p for p in (known_parts or [])}

    if known_upper:
        if stem_up in known_upper:
            return known_upper[stem_up]
        # 文件名包含已知料号时取最长匹配
        hits = [orig for up, orig in known_upper.items() if up in stem_up]
        if hits:
            hits.sort(key=len, reverse=True)
            return hits[0]
        return None

    m = _DASH_PART_RE.search(stem)
    if m:
        return m.group(1)
    m = _DIGIT_PART_RE.search(stem)
    if m:
        return m.group(1)
    return None


def match_templates(
    template_paths: list[Path],
    part_nos: list[str],
) -> dict[str, Path]:
    """返回 {料号: 模板路径}；同一料号多个模板时优先文件名带 (1) 的。"""
    part_set = {p.upper(): p for p in part_nos}
    buckets: dict[str, list[Path]] = {p.upper(): [] for p in part_nos}

    for path in template_paths:
        detected = detect_part_no_from_filename(path, part_nos)
        if detected and detected.upper() in buckets:
            buckets[detected.upper()].append(path)

    result: dict[str, Path] = {}
    for up, paths in buckets.items():
        if not paths:
            continue
        paths.sort(key=lambda x: (0 if re.search(r"\(1\)$", x.stem) else 1, len(x.name)))
        result[part_set[up]] = paths[0]
    return result
