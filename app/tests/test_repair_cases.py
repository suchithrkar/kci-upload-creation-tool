from datetime import date, datetime, timedelta

from app.engine.repair_cases import build_repair_cases
from app.schemas.config import MarketMap, SBDConfig, TLMap
from app.schemas.cso_status import CSOStatus
from app.schemas.delivery import DeliveryStatus
from app.schemas.dump import DumpCase
from app.schemas.mo import MaterialOrder
from app.schemas.mo_items import MaterialOrderItem
from app.schemas.so import ServiceOrder
from app.schemas.wo import WorkOrder


def test_build_repair_cases_filters_resolution_code() -> None:
    cases = [
        DumpCase(
            case_id="1",
            customer_name="A",
            created_on=datetime(2024, 1, 1),
            created_by="user",
            country="US",
            resolution_code="parts shipped",
            case_owner="Agent",
            otc_code="OTC",
            serial_number="SN",
            product_name="Product",
            email_status="sent",
        ),
        DumpCase(
            case_id="2",
            customer_name="B",
            created_on=datetime(2024, 1, 1),
            created_by="user",
            country="US",
            resolution_code="invalid",
            case_owner="Agent",
            otc_code="OTC",
            serial_number="SN",
            product_name="Product",
            email_status="sent",
        ),
    ]

    repair_cases = build_repair_cases(
        dump_cases=cases,
        work_orders=[],
        material_orders=[],
        mo_items=[],
        service_orders=[],
        cso_statuses=[],
        delivery_statuses=[],
        tl_map=[],
        market_map=[],
        sbd_config=SBDConfig(
            period_start=date(2024, 1, 1),
            period_end=date(2024, 12, 31),
            cutoff_countries=["US"],
        ),
    )

    assert len(repair_cases) == 1
    assert repair_cases[0].case_id == "1"


def test_build_repair_cases_populates_parts_and_tracking() -> None:
    created_on = datetime.now() - timedelta(days=2)
    cases = [
        DumpCase(
            case_id="CASE-1",
            customer_name="Acme",
            created_on=created_on,
            created_by="user",
            country="US",
            resolution_code="parts shipped",
            case_owner="Agent One",
            otc_code="OTC",
            serial_number="SN",
            product_name="Widget",
            email_status="sent",
        )
    ]
    material_orders = [
        MaterialOrder(
            order_number="MO-1",
            case_id="CASE-1",
            created_on=created_on,
            order_status="shipped",
        )
    ]
    mo_items = [
        MaterialOrderItem(
            order_number="MO-1",
            line_name="Line-1 - 1",
            part_number="PN-1",
            description="Part - Widget",
            tracking_url=None,
        )
    ]
    delivery_statuses = [
        DeliveryStatus(case_id="CASE-1", current_status="Delivered")
    ]

    repair_cases = build_repair_cases(
        dump_cases=cases,
        work_orders=[],
        material_orders=material_orders,
        mo_items=mo_items,
        service_orders=[],
        cso_statuses=[],
        delivery_statuses=delivery_statuses,
        tl_map=[TLMap(name="TL1", agents=["agent one"])],
        market_map=[MarketMap(name="NA", countries=["us"])],
        sbd_config=SBDConfig(
            period_start=created_on.date(),
            period_end=created_on.date(),
            cutoff_countries=["US"],
        ),
    )

    assert repair_cases[0].part_number == "PN-1"
    assert repair_cases[0].part_name == "Widget"
    assert repair_cases[0].tracking_status == "Delivered"
    assert repair_cases[0].csr_rfc == "shipped"
    assert repair_cases[0].tl == "TL1"
    assert repair_cases[0].market == "NA"


def test_build_repair_cases_offsite_dnap() -> None:
    created_on = datetime(2024, 6, 1, 10, 0)
    cases = [
        DumpCase(
            case_id="CASE-2",
            customer_name="Acme",
            created_on=created_on,
            created_by="user",
            country="US",
            resolution_code="offsite solution",
            case_owner="Agent Two",
            otc_code="OTC",
            serial_number="SN",
            product_name="Widget",
            email_status="sent",
        )
    ]
    cso_statuses = [
        CSOStatus(
            case_id="CASE-2",
            cso="CSO-1",
            status="open",
            tracking_number=None,
            repair_status="Product Returned Unrepaired to Customer",
        )
    ]
    work_orders = [
        WorkOrder(
            case_id="CASE-2",
            work_order_number="WO-1",
            created_on=created_on,
            system_status="closed",
            workgroup="WG",
            country="US",
            resolution_notes="done",
        )
    ]
    service_orders = [
        ServiceOrder(
            case_id="CASE-2",
            submitted_on=created_on,
            order_reference_id="SO-1",
        )
    ]

    repair_cases = build_repair_cases(
        dump_cases=cases,
        work_orders=work_orders,
        material_orders=[],
        mo_items=[],
        service_orders=service_orders,
        cso_statuses=cso_statuses,
        delivery_statuses=[],
        tl_map=[TLMap(name="TL2", agents=["agent two"])],
        market_map=[MarketMap(name="NA", countries=["US"])],
        sbd_config=SBDConfig(
            period_start=date(2024, 1, 1),
            period_end=date(2024, 12, 31),
            cutoff_countries=["US"],
        ),
    )

    assert repair_cases[0].bench_rfc == "open"
    assert repair_cases[0].dnap == "True"
