from datetime import datetime, timedelta

from app.engine.closed_cases import build_closed_cases_report
from app.schemas.closed_case import ClosedCase
from app.schemas.repair_case import RepairCase


def test_build_closed_cases_report_filters_months_back() -> None:
    now = datetime.now()
    old_closed = ClosedCase(
        case_id="CASE-1",
        customer_name="Customer",
        created_on=now - timedelta(days=300),
        created_by="creator",
        modified_by="user",
        modified_on=now - timedelta(days=200),
        closed_on=now - timedelta(days=200),
        closed_by="user",
        owner="Owner",
        country="US",
        resolution_code="parts shipped",
        case_owner="Agent",
        otc_code="OTC",
    )

    repair_case = RepairCase(
        case_id="CASE-1",
        customer_name="Customer",
        created_on=now - timedelta(days=300),
        created_by="creator",
        country="US",
        resolution_code="parts shipped",
        case_owner="Agent",
        otc_code="OTC",
        ca_group="0–3 Days",
        tl="TL",
        sbd="Met",
        onsite_rfc="Not Found",
        csr_rfc="shipped",
        bench_rfc="Not Found",
        market="NA",
        wo_closure_notes="Not Found",
        tracking_status="Not Found",
        part_number="Not Found",
        part_name="Not Found",
        serial_number="SN",
        product_name="Product",
        email_status="sent",
        dnap="False",
    )

    report = build_closed_cases_report([old_closed], [repair_case], months_back=1)

    assert report == []


def test_build_closed_cases_report_closed_by_logic() -> None:
    now = datetime.now()
    closed_case = ClosedCase(
        case_id="CASE-2",
        customer_name="Customer",
        created_on=now - timedelta(days=10),
        created_by="creator",
        modified_by="# CrmWebJobUser-Prod",
        modified_on=now - timedelta(days=5),
        closed_on=now - timedelta(days=5),
        closed_by="user",
        owner="Owner",
        country="US",
        resolution_code="parts shipped",
        case_owner="Agent",
        otc_code="OTC",
    )

    repair_case = RepairCase(
        case_id="CASE-2",
        customer_name="Customer",
        created_on=now - timedelta(days=10),
        created_by="creator",
        country="US",
        resolution_code="parts shipped",
        case_owner="Agent",
        otc_code="OTC",
        ca_group="0–3 Days",
        tl="TL",
        sbd="Met",
        onsite_rfc="Not Found",
        csr_rfc="shipped",
        bench_rfc="Not Found",
        market="NA",
        wo_closure_notes="Not Found",
        tracking_status="Not Found",
        part_number="Not Found",
        part_name="Not Found",
        serial_number="SN",
        product_name="Product",
        email_status="sent",
        dnap="False",
    )

    reports = build_closed_cases_report([closed_case], [repair_case], months_back=6)

    assert reports[0].closed_by == "CRM Auto Closed"
    assert reports[0].tl == "TL"
    assert reports[0].sbd == "Met"
    assert reports[0].market == "NA"


def test_build_closed_cases_report_owner_override() -> None:
    now = datetime.now()
    closed_case = ClosedCase(
        case_id="CASE-3",
        customer_name="Customer",
        created_on=now - timedelta(days=10),
        created_by="creator",
        modified_by="SYSTEM",
        modified_on=now - timedelta(days=5),
        closed_on=now - timedelta(days=5),
        closed_by="user",
        owner="Owner Value",
        country="US",
        resolution_code="parts shipped",
        case_owner="Agent",
        otc_code="OTC",
    )

    repair_case = RepairCase(
        case_id="CASE-3",
        customer_name="Customer",
        created_on=now - timedelta(days=10),
        created_by="creator",
        country="US",
        resolution_code="parts shipped",
        case_owner="Agent",
        otc_code="OTC",
        ca_group="0–3 Days",
        tl="TL",
        sbd="Met",
        onsite_rfc="Not Found",
        csr_rfc="shipped",
        bench_rfc="Not Found",
        market="NA",
        wo_closure_notes="Not Found",
        tracking_status="Not Found",
        part_number="Not Found",
        part_name="Not Found",
        serial_number="SN",
        product_name="Product",
        email_status="sent",
        dnap="False",
    )

    reports = build_closed_cases_report([closed_case], [repair_case], months_back=6)

    assert reports[0].closed_by == "Owner Value"
