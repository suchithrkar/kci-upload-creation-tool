from datetime import datetime, timedelta

from app.engine.copy_outputs import build_copy_so_orders, build_copy_tracking_urls
from app.schemas.cso_status import CSOStatus
from app.schemas.delivery import DeliveryStatus
from app.schemas.dump import DumpCase
from app.schemas.mo import MaterialOrder
from app.schemas.mo_items import MaterialOrderItem
from app.schemas.so import ServiceOrder


def test_build_copy_so_orders_filters_and_strips_suffix() -> None:
    cases = [
        DumpCase(
            case_id="CASE-1",
            customer_name="Customer",
            created_on=datetime(2024, 1, 1),
            created_by="user",
            country="US",
            resolution_code="offsite solution",
            case_owner="Owner",
            otc_code="OTC",
            serial_number="SN",
            product_name="Product",
            email_status="sent",
        ),
        DumpCase(
            case_id="CASE-2",
            customer_name="Customer",
            created_on=datetime(2024, 1, 1),
            created_by="user",
            country="US",
            resolution_code="parts shipped",
            case_owner="Owner",
            otc_code="OTC",
            serial_number="SN",
            product_name="Product",
            email_status="sent",
        ),
    ]
    service_orders = [
        ServiceOrder(
            case_id="CASE-1",
            submitted_on=datetime(2024, 1, 1, 10, 0),
            order_reference_id="ABC123-01",
        ),
        ServiceOrder(
            case_id="CASE-1",
            submitted_on=datetime(2024, 1, 2, 10, 0),
            order_reference_id="ABC123-02",
        ),
    ]
    cso_statuses = [
        CSOStatus(
            case_id="CASE-1",
            cso="CSO",
            status="Open",
            tracking_number=None,
            repair_status=None,
        )
    ]

    outputs = build_copy_so_orders(cases, service_orders, cso_statuses)

    assert outputs == ["CASE-1,ABC123"]


def test_build_copy_so_orders_excludes_closed_status() -> None:
    cases = [
        DumpCase(
            case_id="CASE-1",
            customer_name="Customer",
            created_on=datetime(2024, 1, 1),
            created_by="user",
            country="US",
            resolution_code="offsite solution",
            case_owner="Owner",
            otc_code="OTC",
            serial_number="SN",
            product_name="Product",
            email_status="sent",
        )
    ]
    service_orders = [
        ServiceOrder(
            case_id="CASE-1",
            submitted_on=datetime(2024, 1, 1, 10, 0),
            order_reference_id="ABC123-01",
        )
    ]
    cso_statuses = [
        CSOStatus(
            case_id="CASE-1",
            cso="CSO",
            status="Delivered",
            tracking_number=None,
            repair_status=None,
        )
    ]

    outputs = build_copy_so_orders(cases, service_orders, cso_statuses)

    assert outputs == []


def test_build_copy_tracking_urls_primary_logic() -> None:
    created_on = datetime.now() - timedelta(days=1)
    cases = [
        DumpCase(
            case_id="CASE-1",
            customer_name="Customer",
            created_on=created_on,
            created_by="user",
            country="US",
            resolution_code="parts shipped",
            case_owner="Owner",
            otc_code="OTC",
            serial_number="SN",
            product_name="Product",
            email_status="sent",
        )
    ]
    material_orders = [
        MaterialOrder(
            order_number="MO-1",
            case_id="CASE-1",
            created_on=created_on,
            order_status="closed",
        )
    ]
    mo_items = [
        MaterialOrderItem(
            order_number="MO-1",
            line_name="Line-1 - 1",
            part_number="PN",
            description="Part - Widget",
            tracking_url="http://tracking",
        )
    ]

    outputs = build_copy_tracking_urls(
        dump_cases=cases,
        material_orders=material_orders,
        mo_items=mo_items,
        cso_statuses=[],
        delivery_statuses=[],
    )

    assert outputs == ["CASE-1 | http://tracking"]


def test_build_copy_tracking_urls_fallback_and_exclusions() -> None:
    created_on = datetime.now() - timedelta(days=1)
    cases = [
        DumpCase(
            case_id="CASE-2",
            customer_name="Customer",
            created_on=created_on,
            created_by="user",
            country="US",
            resolution_code="parts shipped",
            case_owner="Owner",
            otc_code="OTC",
            serial_number="SN",
            product_name="Product",
            email_status="sent",
        ),
        DumpCase(
            case_id="CASE-3",
            customer_name="Customer",
            created_on=created_on,
            created_by="user",
            country="US",
            resolution_code="parts shipped",
            case_owner="Owner",
            otc_code="OTC",
            serial_number="SN",
            product_name="Product",
            email_status="sent",
        ),
    ]
    cso_statuses = [
        CSOStatus(
            case_id="CASE-2",
            cso="CSO",
            status="delivered",
            tracking_number="1Z999",
            repair_status=None,
        )
    ]
    delivery_statuses = [
        DeliveryStatus(case_id="CASE-3", current_status="In Transit")
    ]

    outputs = build_copy_tracking_urls(
        dump_cases=cases,
        material_orders=[],
        mo_items=[],
        cso_statuses=cso_statuses,
        delivery_statuses=delivery_statuses,
    )

    assert outputs == [
        "CASE-2 | http://wwwapps.ups.com/WebTracking/processInputRequest?TypeOfInquiryNumber=T&InquiryNumber1=1Z999"
    ]
