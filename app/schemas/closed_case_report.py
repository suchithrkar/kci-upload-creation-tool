"""Schemas for closed case reporting."""

from datetime import datetime

from pydantic import BaseModel


class ClosedCaseReport(BaseModel):
    """Represents a closed case report record."""

    case_id: str
    customer_name: str
    created_on: datetime
    created_by: str
    modified_by: str
    modified_on: datetime
    closed_on: datetime
    closed_by: str
    country: str
    resolution_code: str
    case_owner: str
    otc_code: str
    tl: str
    sbd: str
    market: str
