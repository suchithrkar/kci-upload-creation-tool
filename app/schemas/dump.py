"""Schemas for dump case records."""

from datetime import datetime
from typing import Literal

from pydantic import BaseModel


class DumpCase(BaseModel):
    """Represents a raw case record from the dump export."""

    case_id: str
    customer_name: str
    created_on: datetime
    created_by: str
    country: str
    resolution_code: Literal[
        "parts shipped",
        "onsite solution",
        "offsite solution",
    ]
    case_owner: str
    otc_code: str
    serial_number: str
    product_name: str
    email_status: str
