"""Schemas for material order line items."""

from typing import Optional

from pydantic import BaseModel


class MaterialOrderItem(BaseModel):
    """Represents a single line item within a material order."""

    order_number: str
    line_name: str
    part_number: str
    description: str
    tracking_url: Optional[str]
