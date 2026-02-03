"""Utilities for calculating SBD outcomes."""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import Optional

from app.schemas.config import SBDConfig


def calculate_sbd(
    case_created_on: datetime,
    country: str,
    first_order_date: Optional[datetime],
    config: SBDConfig,
) -> str:
    """Calculate the SBD result for a case using cutoff and rollover rules."""

    if country not in config.cutoff_countries:
        return "NA"

    created_date = case_created_on.date()
    if created_date < config.period_start or created_date > config.period_end:
        return "NA"

    if first_order_date is None:
        return "Not Met"

    deadline = created_date + timedelta(days=1)
    if first_order_date.date() <= deadline:
        return "Met"

    return "Not Met"
