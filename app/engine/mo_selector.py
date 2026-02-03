"""Utilities for selecting the relevant material order for a case."""

from __future__ import annotations

from datetime import timedelta
from typing import Optional

from app.schemas.mo import MaterialOrder

_STATUS_PRIORITY = {
    "closed": 1,
    "pod": 2,
    "shipped": 3,
    "ordered": 4,
    "partially ordered": 5,
    "order pending": 6,
    "new": 7,
    "cancelled": 8,
    "unknown": 9,
}


def _priority(status: str) -> int:
    return _STATUS_PRIORITY.get(status.strip().lower(), _STATUS_PRIORITY["unknown"])


def get_latest_mo(
    case_id: str, mos: list[MaterialOrder]
) -> Optional[MaterialOrder]:
    """Return the latest material order within a 5-minute window by priority."""

    candidates = [mo for mo in mos if mo.case_id == case_id]
    if not candidates:
        return None

    latest_created = max(mo.created_on for mo in candidates)
    window_start = latest_created - timedelta(minutes=5)
    windowed = [mo for mo in candidates if mo.created_on >= window_start]

    def sort_key(order: MaterialOrder) -> tuple[int, float]:
        created_ts = order.created_on.timestamp()
        return (_priority(order.order_status), -created_ts)

    return min(windowed, key=sort_key)
