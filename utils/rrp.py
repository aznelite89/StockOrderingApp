"""RRP computation — Python mirror of the IFS formula in the sample workbook."""

from __future__ import annotations

import math
from typing import Union

from constants.invoice import CHECK_RRP_SENTINEL


def _ceiling(value: float, significance: float) -> float:
    """Excel CEILING(value, significance) — round up to next multiple."""
    if significance == 0:
        return 0.0
    return math.ceil(value / significance) * significance


def compute_rrp(total_after_discount: float) -> Union[float, str]:
    """Return the RRP for a given Total After Discount (Ex-GST).

    Mirrors the formula in cell J7 of the sample workbook:
        IFS(I<380,        CEILING(I*2.8,10)-1,
            I<=1400,      IF(I*2.5<300, CEILING(I*2.5,10)-1, CEILING(I*2.5,50)-1),
            I>1400,       "Check RRP")
    """
    i = float(total_after_discount)
    if i < 380:
        return _ceiling(i * 2.8, 10) - 1
    if i <= 1400:
        marked_up = i * 2.5
        if marked_up < 300:
            return _ceiling(marked_up, 10) - 1
        return _ceiling(marked_up, 50) - 1
    return CHECK_RRP_SENTINEL
