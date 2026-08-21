"""MIL-STD-105E 正常检验（一般水平 II）查表。"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class AqlResult:
    letter: str
    sample_size: int
    maj_ac: int
    maj_re: int
    min_ac: int
    min_re: int


# 批量 → 样字 / 样本量
_LOT_TABLE: list[tuple[int, int, str, int]] = [
    (2, 8, "A", 2),
    (9, 15, "B", 3),
    (16, 25, "C", 5),
    (26, 50, "D", 8),
    (51, 90, "E", 13),
    (91, 150, "F", 20),
    (151, 280, "G", 32),
    (281, 500, "H", 50),
    (501, 1200, "J", 80),
    (1201, 3200, "K", 125),
    (3201, 10000, "L", 200),
    (10001, 35000, "M", 315),
    (35001, 150000, "N", 500),
    (150001, 500000, "P", 800),
    (500001, 10**12, "Q", 1250),
]

# J 起：MAJ=AQL0.65，MIN=AQL1.0（按确认书）
_AC_RE_FROM_J: dict[str, tuple[tuple[int, int], tuple[int, int]]] = {
    "J": ((1, 2), (2, 3)),
    "K": ((2, 3), (3, 4)),
    "L": ((3, 4), (5, 6)),
    "M": ((5, 6), (7, 8)),
    "N": ((7, 8), (10, 11)),
    "P": ((10, 11), (14, 15)),
    "Q": ((14, 15), (21, 22)),
    "R": ((21, 22), (30, 31)),
}

_AH_LETTERS = set("ABCDEFGH")


def lookup_aql(lot_qty: int) -> AqlResult:
    """根据本次发货数量查样字、样本数、允收/拒收。"""
    if lot_qty < 2:
        raise ValueError("发货数量至少为 2")

    letter = None
    sample_size = None
    for low, high, code, size in _LOT_TABLE:
        if low <= lot_qty <= high:
            letter, sample_size = code, size
            break
    if letter is None:
        raise ValueError(f"发货数量超出查表范围: {lot_qty}")

    # 样本数 ≥ 批量 → 全检（表注）
    if sample_size >= lot_qty:
        sample_size = lot_qty

    if letter in _AH_LETTERS:
        maj_ac, maj_re = 0, 1
        min_ac, min_re = 1, 2
    else:
        maj, minor = _AC_RE_FROM_J[letter]
        maj_ac, maj_re = maj
        min_ac, min_re = minor

    return AqlResult(
        letter=letter,
        sample_size=sample_size,
        maj_ac=maj_ac,
        maj_re=maj_re,
        min_ac=min_ac,
        min_re=min_re,
    )
