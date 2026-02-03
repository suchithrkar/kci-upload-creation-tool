"""Schemas for closed case inputs."""

from datetime import datetime

from pydantic import BaseModel


class ClosedCase(BaseModel):
    """Represents a closed case record from the source system."""

    case_id: str
    customer_name: str
    created_on: datetime
    created_by: str
    modified_by: str
    modified_on: datetime
    closed_on: datetime
    closed_by: str
    owner: str
    country: str
    resolution_code: str
    case_owner: str
    otc_code: str
