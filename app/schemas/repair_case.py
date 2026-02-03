"""Schemas for consolidated repair case data."""

from datetime import datetime

from pydantic import BaseModel


class RepairCase(BaseModel):
    """Represents a final derived repair case record."""

    case_id: str
    customer_name: str
    created_on: datetime
    created_by: str
    country: str
    resolution_code: str
    case_owner: str
    otc_code: str
    ca_group: str
    tl: str
    sbd: str
    onsite_rfc: str
    csr_rfc: str
    bench_rfc: str
    market: str
    wo_closure_notes: str
    tracking_status: str
    part_number: str
    part_name: str
    serial_number: str
    product_name: str
    email_status: str
    dnap: str
