"""Schemas for customer service order status data."""

from typing import Optional

from pydantic import BaseModel


class CSOStatus(BaseModel):
    """Represents the status of a customer service order."""

    case_id: str
    cso: str
    status: str
    tracking_number: Optional[str]
    repair_status: Optional[str]
