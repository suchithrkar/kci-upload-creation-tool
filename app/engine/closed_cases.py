"""Closed case reporting logic."""

from __future__ import annotations

from datetime import datetime

from app.schemas.closed_case import ClosedCase
from app.schemas.closed_case_report import ClosedCaseReport
from app.schemas.repair_case import RepairCase

_AUTOCLOSE_USER = "# CrmWebJobUser-Prod"
_OWNER_OVERRIDE_USERS = {
    "# MSFT-ServiceSystemAdmin",
    "# CrmEEGUser-Prod",
    "# MSFT-ServiceSystemAdminDev",
    "SYSTEM",
}


def _month_start_months_back(reference: datetime, months_back: int) -> datetime:
    year = reference.year
    month = reference.month - months_back
    while month <= 0:
        month += 12
        year -= 1
    return reference.replace(year=year, month=month, day=1, hour=0, minute=0, second=0, microsecond=0)


def _closed_by_value(modified_by: str, owner: str) -> str:
    if modified_by == _AUTOCLOSE_USER:
        return "CRM Auto Closed"
    if modified_by in _OWNER_OVERRIDE_USERS:
        return owner
    return modified_by


def build_closed_cases_report(
    closed_cases: list[ClosedCase],
    repair_cases: list[RepairCase],
    months_back: int = 6,
) -> list[ClosedCaseReport]:
    """Build closed case reports for cases within the requested month window."""

    cutoff = _month_start_months_back(datetime.now(), months_back)
    repair_lookup = {case.case_id: case for case in repair_cases}
    reports: list[ClosedCaseReport] = []

    for closed_case in closed_cases:
        if closed_case.closed_on < cutoff:
            continue

        repair_case = repair_lookup.get(closed_case.case_id)
        if repair_case is None:
            continue

        reports.append(
            ClosedCaseReport(
                case_id=closed_case.case_id,
                customer_name=closed_case.customer_name,
                created_on=closed_case.created_on,
                created_by=closed_case.created_by,
                modified_by=closed_case.modified_by,
                modified_on=closed_case.modified_on,
                closed_on=closed_case.closed_on,
                closed_by=_closed_by_value(
                    closed_case.modified_by, closed_case.owner
                ),
                country=closed_case.country,
                resolution_code=closed_case.resolution_code,
                case_owner=closed_case.case_owner,
                otc_code=closed_case.otc_code,
                tl=repair_case.tl,
                sbd=repair_case.sbd,
                market=repair_case.market,
            )
        )

    return reports
