"""从模板尺寸行解析公差，并生成波动较小的实测值。"""

from __future__ import annotations

import random
import re
from dataclasses import dataclass


@dataclass
class Tolerance:
    nominal: float
    low: float
    high: float
    decimals: int
    raw: str


# 21.5mm±0.3 / 146.2±0.3 / 29.0mm+0.1/-0.3 / 3.5mm+0.5/-0 / 1mm±0.2
_SPEC_RE = re.compile(
    r"(?P<nom>\d+(?:\.\d+)?)\s*(?:mm)?\s*"
    r"(?:"
    r"±\s*(?P<sym>\d+(?:\.\d+)?)"
    r"|"
    r"\+\s*(?P<up>\d+(?:\.\d+)?)\s*/\s*-\s*(?P<down>\d+(?:\.\d+)?)"
    r")",
    re.I,
)


def _decimals(text: str) -> int:
    if "." in text:
        return len(text.split(".", 1)[1])
    return 0


def parse_tolerance(spec: str) -> Tolerance | None:
    all_t = parse_all_tolerances(spec)
    return all_t[0] if all_t else None


def parse_all_tolerances(spec: str) -> list[Tolerance]:
    """找出文本中所有公差表达式（按出现顺序）。"""
    text = (spec or "").replace(" ", "")
    results: list[Tolerance] = []
    for m in _SPEC_RE.finditer(text):
        nom_s = m.group("nom")
        nominal = float(nom_s)
        dec = _decimals(nom_s)
        if m.group("sym") is not None:
            tol = float(m.group("sym"))
            dec = max(dec, _decimals(m.group("sym")))
            low, high = nominal - tol, nominal + tol
        else:
            up = float(m.group("up"))
            down = float(m.group("down"))
            dec = max(dec, _decimals(m.group("up")), _decimals(m.group("down")))
            low, high = nominal - down, nominal + up
        results.append(
            Tolerance(nominal=nominal, low=low, high=high, decimals=dec, raw=m.group(0))
        )
    return results


def pick_tolerance_for_samples(
    candidates: list[Tolerance],
    sample_values: list[float],
) -> Tolerance | None:
    """
    在多个候选公差里，选与现有样本最吻合的一个。
    用来避免「行号 5 + 120」被误读成 5120。
    """
    if not candidates:
        return None
    if not sample_values:
        return candidates[0]

    vals = sorted(sample_values)
    med = vals[len(vals) // 2]

    # 1) 样本中位数落在公差带内（含少量放宽）的优先
    fitting: list[Tolerance] = []
    for t in candidates:
        pad = max((t.high - t.low) * 0.5, abs(t.nominal) * 0.02, 0.05)
        if (t.low - pad) <= med <= (t.high + pad):
            fitting.append(t)
    if fitting:
        return min(fitting, key=lambda t: abs(t.nominal - med))

    # 2) 否则选标称最接近中位数的，但差距过大则判定失败（宁可不改这一行）
    best = min(candidates, key=lambda t: abs(t.nominal - med))
    limit = max(abs(med) * 0.35, (best.high - best.low) * 8, 1.0)
    if abs(best.nominal - med) > limit:
        return None
    return best


def _fmt(value: float, decimals: int) -> str:
    # 与常见报告一致：至少保留与标称相同的小数位
    q = round(value, decimals)
    return f"{q:.{decimals}f}"


def generate_samples(
    tol: Tolerance,
    count: int = 10,
    *,
    rng: random.Random | None = None,
) -> list[str]:
    """
    在公差内生成实测值：
    - 围绕标称附近小幅波动（默认不超过公差带宽的约 35%）
    - 绝不超出上下限
    """
    rng = rng or random.Random()
    span = tol.high - tol.low
    if span <= 0:
        return [_fmt(tol.nominal, tol.decimals)] * count

    # 相对带宽的“品控稳定”振幅
    amp = span * 0.35
    # 绝对下限：极窄公差（如 ±0.02）时仍允许细微变化
    min_step = 10 ** (-tol.decimals) if tol.decimals > 0 else 0.01
    amp = max(amp, min_step)

    values: list[float] = []
    for _ in range(count):
        # 正态采样，截断到 [nominal-amp, nominal+amp] 再夹到公差
        for _try in range(20):
            v = rng.gauss(tol.nominal, amp / 2.5)
            if abs(v - tol.nominal) <= amp:
                break
        v = min(max(v, tol.low), tol.high)
        # 再向标称轻微拉回，避免贴边过多
        v = tol.nominal + (v - tol.nominal) * 0.85
        v = min(max(v, tol.low), tol.high)
        values.append(v)

    # 若公差极窄且标称可整除到该精度，多数保持标称（如 0.25±0.02）
    if span <= 2 * (10 ** (-tol.decimals)):
        values = [tol.nominal for _ in range(count)]

    return [_fmt(v, max(tol.decimals, 2) if tol.decimals >= 2 else max(tol.decimals, 2)) for v in values]


def generate_samples_smart(tol: Tolerance, count: int = 10, rng: random.Random | None = None) -> list[str]:
    """按标称小数位格式化；窄公差时多数贴近标称。"""
    rng = rng or random.Random()
    span = tol.high - tol.low
    decimals = tol.decimals
    # 报告常见至少两位小数；更细的公差保留更多位
    out_decimals = max(decimals, 2)

    # 品控观感：波动控制在公差带宽约 22% 以内
    amp = max(span * 0.22, 10 ** (-out_decimals))
    # 极窄：几乎不变
    if span <= 0.04:
        base = [_fmt(tol.nominal, out_decimals) for _ in range(count)]
        for i in range(count):
            if rng.random() < 0.25:
                delta = (10 ** (-out_decimals)) * rng.choice([-1, 1])
                v = tol.nominal + delta
                if tol.low <= v <= tol.high:
                    base[i] = _fmt(v, out_decimals)
        return base

    result: list[str] = []
    for _ in range(count):
        v = rng.gauss(tol.nominal, amp / 2.8)
        v = tol.nominal + max(-amp, min(amp, v - tol.nominal))
        v = min(max(v, tol.low + 1e-12), tol.high - 1e-12)
        result.append(_fmt(v, out_decimals))
    return result
