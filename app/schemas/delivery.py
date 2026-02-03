"""Schemas for delivery status data."""

from pydantic import BaseModel


class DeliveryStatus(BaseModel):
    """Represents delivery status for a case."""

    case_id: str
    current_status: str
