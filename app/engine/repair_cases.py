"""Repair case generation logic."""

from __future__ import annotations

from datetime import datetime
from typing import Optional

from app.engine.ca_group import get_ca_group
from app.engine.mo_selector import get_latest_mo
from app.engine.sbd import calculate_sbd
from app.schemas.config import MarketMap, SBDConfig, TLMap
from app.schemas.cso_status import CSOStatus
from app.schemas.delivery import DeliveryStatus
from app.schemas.dump import DumpCase
from app.schemas.mo import MaterialOrder
from app.schemas.mo_items import MaterialOrderItem
from app.schemas.repair_case import RepairCase
from app.schemas.so import ServiceOrder
from app.schemas.wo import WorkOrder

_VALID_RESOLUTIONS = {
    "parts shipped",
    "onsite solution",
    "offsite solution",
}


def _find_tl(case_owner: str, tl_map: list[TLMap]) -> str:
    owner = case_owner.casefold()
    for mapping in tl_map:
        if any(owner == agent.casefold() for agent in mapping.agents):
            return mapping.name
    return "Not Found"


def _find_market(country: str, market_map: list[MarketMap]) -> str:
    country_value = country.casefold()
    for mapping in market_map:
        if any(country_value == entry.casefold() for entry in mapping.countries):
            return mapping.name
    return "Not Found"


def _latest_work_order(work_orders: list[WorkOrder]) -> Optional[WorkOrder]:
    if not work_orders:
        return None
    return max(work_orders, key=lambda wo: wo.created_on)


def _first_order_date(
    work_orders: list[WorkOrder],
    material_orders: list[MaterialOrder],
    service_orders: list[ServiceOrder],
) -> Optional[datetime]:
    timestamps: list[datetime] = [
        *[wo.created_on for wo in work_orders],
        *[mo.created_on for mo in material_orders],
        *[so.submitted_on for so in service_orders],
    ]
    if not timestamps:
        return None
    return min(timestamps)


def _find_cso_status(case_id: str, cso_statuses: list[CSOStatus]) -> Optional[CSOStatus]:
    for status in cso_statuses:
        if status.case_id == case_id:
            return status
    return None


def _find_delivery_status(
    case_id: str, delivery_statuses: list[DeliveryStatus]
) -> Optional[DeliveryStatus]:
    for status in delivery_statuses:
        if status.case_id == case_id:
            return status
    return None


def _part_details(
    material_order: MaterialOrder,
    mo_items: list[MaterialOrderItem],
) -> tuple[str, str]:
    candidates = [
        item
        for item in mo_items
        if item.order_number == material_order.order_number
        and item.line_name.endswith("- 1")
    ]
    if not candidates:
        return "Not Found", "Not Found"

    item = candidates[0]
    description = item.description
    if "-" in description:
        part_name = description.split("-", 1)[1].strip()
    else:
        part_name = description.strip()
    return item.part_number, part_name


def build_repair_cases(
    dump_cases: list[DumpCase],
    work_orders: list[WorkOrder],
    material_orders: list[MaterialOrder],
    mo_items: list[MaterialOrderItem],
    service_orders: list[ServiceOrder],
    cso_statuses: list[CSOStatus],
    delivery_statuses: list[DeliveryStatus],
    tl_map: list[TLMap],
    market_map: list[MarketMap],
    sbd_config: SBDConfig,
) -> list[RepairCase]:
    """Build consolidated repair case records from source data."""

    repair_cases: list[RepairCase] = []

    for case in dump_cases:
        if case.resolution_code not in _VALID_RESOLUTIONS:
            continue

        case_work_orders = [wo for wo in work_orders if wo.case_id == case.case_id]
        case_material_orders = [
            mo for mo in material_orders if mo.case_id == case.case_id
        ]
        case_service_orders = [
            so for so in service_orders if so.case_id == case.case_id
        ]

        latest_work_order = _latest_work_order(case_work_orders)
        latest_material_order = get_latest_mo(case.case_id, case_material_orders)
        cso_status = _find_cso_status(case.case_id, cso_statuses)
        delivery_status = _find_delivery_status(case.case_id, delivery_statuses)

        ca_group = get_ca_group(case.created_on)
        tl = _find_tl(case.case_owner, tl_map)
        market = _find_market(case.country, market_map)
        first_order_date = _first_order_date(
            case_work_orders, case_material_orders, case_service_orders
        )
        sbd = calculate_sbd(
            case_created_on=case.created_on,
            country=case.country,
            first_order_date=first_order_date,
            config=sbd_config,
        )

        onsite_rfc = "Not Found"
        csr_rfc = "Not Found"
        bench_rfc = "Not Found"
        wo_closure_notes = "Not Found"
        part_number = "Not Found"
        part_name = "Not Found"
        tracking_status = "Not Found"

        if case.resolution_code == "onsite solution":
            if latest_work_order:
                onsite_rfc = latest_work_order.system_status
                wo_closure_notes = latest_work_order.resolution_notes
        elif case.resolution_code == "parts shipped":
            if latest_material_order:
                csr_rfc = latest_material_order.order_status
                part_number, part_name = _part_details(
                    latest_material_order, mo_items
                )
        elif case.resolution_code == "offsite solution":
            if cso_status:
                bench_rfc = cso_status.status

        if delivery_status:
            tracking_status = delivery_status.current_status

        dnap = "False"
        if case.resolution_code == "offsite solution" and cso_status:
            if cso_status.repair_status and "product returned unrepaired to customer" in (
                cso_status.repair_status.casefold()
            ):
                dnap = "True"

        repair_cases.append(
            RepairCase(
                case_id=case.case_id,
                customer_name=case.customer_name,
                created_on=case.created_on,
                created_by=case.created_by,
                country=case.country,
                resolution_code=case.resolution_code,
                case_owner=case.case_owner,
                otc_code=case.otc_code,
                ca_group=ca_group,
                tl=tl,
                sbd=sbd,
                onsite_rfc=onsite_rfc,
                csr_rfc=csr_rfc,
                bench_rfc=bench_rfc,
                market=market,
                wo_closure_notes=wo_closure_notes,
                tracking_status=tracking_status,
                part_number=part_number,
                part_name=part_name,
                serial_number=case.serial_number,
                product_name=case.product_name,
                email_status=case.email_status,
                dnap=dnap,
            )
        )

    return repair_cases
