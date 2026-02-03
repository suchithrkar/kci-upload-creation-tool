"""Schemas for work order data."""

from datetime import datetime

from pydantic import BaseModel


class WorkOrder(BaseModel):
    """Represents a work order linked to a case."""

    case_id: str
    work_order_number: str
    created_on: datetime
    system_status: str
    workgroup: str
    country: str
    resolution_notes: str
