from datetime import datetime, timedelta

from app.engine.mo_selector import get_latest_mo
from app.schemas.mo import MaterialOrder


def test_get_latest_mo_filters_by_case() -> None:
    now = datetime.now()
    mos = [
        MaterialOrder(
            order_number="1",
            case_id="CASE-1",
            created_on=now,
            order_status="shipped",
        ),
        MaterialOrder(
            order_number="2",
            case_id="CASE-2",
            created_on=now,
            order_status="closed",
        ),
    ]

    result = get_latest_mo("CASE-1", mos)

    assert result is not None
    assert result.order_number == "1"


def test_get_latest_mo_uses_window_and_priority() -> None:
    now = datetime.now()
    mos = [
        MaterialOrder(
            order_number="1",
            case_id="CASE-1",
            created_on=now - timedelta(minutes=10),
            order_status="closed",
        ),
        MaterialOrder(
            order_number="2",
            case_id="CASE-1",
            created_on=now - timedelta(minutes=3),
            order_status="ordered",
        ),
        MaterialOrder(
            order_number="3",
            case_id="CASE-1",
            created_on=now - timedelta(minutes=2),
            order_status="shipped",
        ),
    ]

    result = get_latest_mo("CASE-1", mos)

    assert result is not None
    assert result.order_number == "3"
