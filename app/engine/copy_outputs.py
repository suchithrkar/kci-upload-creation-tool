"""Helpers for building copy-ready output lists."""

from __future__ import annotations

from typing import Optional

from app.engine.mo_selector import get_latest_mo
from app.schemas.cso_status import CSOStatus
from app.schemas.delivery import DeliveryStatus
from app.schemas.dump import DumpCase
from app.schemas.mo import MaterialOrder
from app.schemas.mo_items import MaterialOrderItem
from app.schemas.so import ServiceOrder

_EXCLUDED_CSO_STATUSES = {
    "delivered",
    "order cancelled, not to be reopened",
}

_ALLOWED_MO_STATUSES = {"closed", "pod"}

_UPS_TRACKING_URL = (
    "http://wwwapps.ups.com/WebTracking/processInputRequest?"
    "TypeOfInquiryNumber=T&InquiryNumber1={tracking_number}"
)


def _latest_service_order(
    service_orders: list[ServiceOrder],
) -> Optional[ServiceOrder]:
    if not service_orders:
        return None
    return max(service_orders, key=lambda order: order.submitted_on)


def _strip_order_suffix(order_reference_id: str) -> str:
    return order_reference_id.split("-", 1)[0]


def _find_cso_status(case_id: str, cso_statuses: list[CSOStatus]) -> Optional[CSOStatus]:
    for status in cso_statuses:
        if status.case_id == case_id:
            return status
    return None


def _delivery_exclusions(delivery_statuses: list[DeliveryStatus]) -> set[str]:
    exclusions = set()
    for status in delivery_statuses:
        current = status.current_status.strip()
        if current and current.casefold() != "no status found":
            exclusions.add(status.case_id)
    return exclusions


def build_copy_so_orders(
    dump_cases: list[DumpCase],
    service_orders: list[ServiceOrder],
    cso_statuses: list[CSOStatus],
) -> list[str]:
    """Build copy output lines for service orders tied to offsite cases."""

    outputs: list[str] = []
    for case in dump_cases:
        if case.resolution_code != "offsite solution":
            continue

        status = _find_cso_status(case.case_id, cso_statuses)
        if status and status.status.casefold() in _EXCLUDED_CSO_STATUSES:
            continue

        case_service_orders = [
            order for order in service_orders if order.case_id == case.case_id
        ]
        latest_order = _latest_service_order(case_service_orders)
        if not latest_order:
            continue

        order_id = _strip_order_suffix(latest_order.order_reference_id)
        outputs.append(f"{case.case_id},{order_id}")

    return outputs


def build_copy_tracking_urls(
    dump_cases: list[DumpCase],
    material_orders: list[MaterialOrder],
    mo_items: list[MaterialOrderItem],
    cso_statuses: list[CSOStatus],
    delivery_statuses: list[DeliveryStatus],
) -> list[str]:
    """Build copy output lines for tracking URLs based on case rules."""

    outputs: list[str] = []
    excluded_cases = _delivery_exclusions(delivery_statuses)

    for case in dump_cases:
        if case.case_id in excluded_cases:
            continue

        url: Optional[str] = None

        if case.resolution_code == "parts shipped":
            case_material_orders = [
                order for order in material_orders if order.case_id == case.case_id
            ]
            latest_order = get_latest_mo(case.case_id, case_material_orders)
            if latest_order and latest_order.order_status.casefold() in _ALLOWED_MO_STATUSES:
                for item in mo_items:
                    if (
                        item.order_number == latest_order.order_number
                        and item.line_name.endswith("- 1")
                        and item.tracking_url
                    ):
                        url = item.tracking_url
                        break

        if url is None:
            status = _find_cso_status(case.case_id, cso_statuses)
            if status and status.status.casefold() == "delivered" and status.tracking_number:
                url = _UPS_TRACKING_URL.format(tracking_number=status.tracking_number)

        if url:
            outputs.append(f"{case.case_id} | {url}")

    return outputs
