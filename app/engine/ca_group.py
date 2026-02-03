"""Utilities for computing case-age grouping buckets."""

from __future__ import annotations

from datetime import datetime


def get_ca_group(created_on: datetime) -> str:
    """Return the case-age group label based on the created timestamp."""

    now = datetime.now(tz=created_on.tzinfo)
    age_days = max((now - created_on).total_seconds() / 86_400, 0)

    if age_days <= 3:
        return "0–3 Days"
    if age_days <= 5:
        return "3–5 Days"
    if age_days <= 10:
        return "5–10 Days"
    if age_days <= 15:
        return "10–15 Days"
    if age_days <= 30:
        return "15–30 Days"
    if age_days <= 60:
        return "30–60 Days"
    if age_days <= 90:
        return "60–90 Days"
    return "> 90 Days"
