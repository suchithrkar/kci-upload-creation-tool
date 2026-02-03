"""Schemas for material order data."""

from datetime import datetime

from pydantic import BaseModel


class MaterialOrder(BaseModel):
    """Represents a material order tied to a case."""

    order_number: str
    case_id: str
    created_on: datetime
    order_status: str
