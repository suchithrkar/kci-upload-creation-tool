"""Schemas for service order data."""

from datetime import datetime

from pydantic import BaseModel


class ServiceOrder(BaseModel):
    """Represents a service order submission."""

    case_id: str
    submitted_on: datetime
    order_reference_id: str
