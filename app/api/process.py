"""Endpoints for processing loaded data into outputs."""

from __future__ import annotations

from datetime import date
from typing import Any

from fastapi import APIRouter, HTTPException
from fastapi.responses import PlainTextResponse

from app.core.state import app_state
from app.engine.closed_cases import build_closed_cases_report
from app.engine.copy_outputs import build_copy_so_orders, build_copy_tracking_urls
from app.engine.repair_cases import build_repair_cases
from app.schemas.config import SBDConfig

router = APIRouter(prefix="/process", tags=["process"])


def _require_data(name: str, data: list) -> None:
    if not data:
        raise HTTPException(status_code=400, detail=f"Missing required data: {name}.")


def _has_closed_case_fields(item: Any) -> bool:
    return all(
        hasattr(item, field)
        for field in (
            "case_id",
            "modified_by",
            "modified_on",
            "closed_on",
            "owner",
        )
    )


@router.post("/repair-cases")
def process_repair_cases() -> list:
    _require_data("dump", app_state.dump)
    _require_data("wo", app_state.wo)
    _require_data("mo", app_state.mo)
    _require_data("mo_items", app_state.mo_items)
    _require_data("so", app_state.so)

    repair_cases = build_repair_cases(
        dump_cases=app_state.dump,
        work_orders=app_state.wo,
        material_orders=app_state.mo,
        mo_items=app_state.mo_items,
        service_orders=app_state.so,
        cso_statuses=app_state.cso,
        delivery_statuses=app_state.delivery,
        tl_map=[],
        market_map=[],
        sbd_config=SBDConfig(
            period_start=date.min,
            period_end=date.max,
            cutoff_countries=[],
        ),
    )
    return [_model_dump(case) for case in repair_cases]


@router.post("/closed-cases")
def process_closed_cases() -> list:
    _require_data("closed_cases", app_state.dump)
    if not all(_has_closed_case_fields(item) for item in app_state.dump):
        raise HTTPException(
            status_code=400,
            detail="Closed cases data not loaded.",
        )
    repair_cases = build_repair_cases(
        dump_cases=app_state.dump,
        work_orders=app_state.wo,
        material_orders=app_state.mo,
        mo_items=app_state.mo_items,
        service_orders=app_state.so,
        cso_statuses=app_state.cso,
        delivery_statuses=app_state.delivery,
        tl_map=[],
        market_map=[],
        sbd_config=SBDConfig(
            period_start=date.min,
            period_end=date.max,
            cutoff_countries=[],
        ),
    )
    if not repair_cases:
        raise HTTPException(status_code=400, detail="Missing required data: repair_cases.")
    closed_reports = build_closed_cases_report(app_state.dump, repair_cases)
    return [_model_dump(report) for report in closed_reports]


@router.get("/copy-so", response_class=PlainTextResponse)
def process_copy_so() -> str:
    _require_data("dump", app_state.dump)
    _require_data("so", app_state.so)
    lines = build_copy_so_orders(
        dump_cases=app_state.dump,
        service_orders=app_state.so,
        cso_statuses=app_state.cso,
    )
    return "\n".join(lines)


@router.get("/copy-tracking", response_class=PlainTextResponse)
def process_copy_tracking() -> str:
    _require_data("dump", app_state.dump)
    _require_data("mo", app_state.mo)
    _require_data("mo_items", app_state.mo_items)
    lines = build_copy_tracking_urls(
        dump_cases=app_state.dump,
        material_orders=app_state.mo,
        mo_items=app_state.mo_items,
        cso_statuses=app_state.cso,
        delivery_statuses=app_state.delivery,
    )
    return "\n".join(lines)


def _model_dump(model: Any) -> dict:
    dump = getattr(model, "model_dump", None)
    if callable(dump):
        return dump()
    return model.dict()
