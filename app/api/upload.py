from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Callable

import pandas as pd
from fastapi import APIRouter, File, HTTPException, UploadFile

from app.core.loaders.cso_loader import load_cso
from app.core.loaders.delivery_loader import load_delivery
from app.core.loaders.dump_loader import load_dump
from app.core.loaders.mo_items_loader import load_mo_items
from app.core.loaders.mo_loader import load_mo
from app.core.loaders.so_loader import load_so
from app.core.loaders.wo_loader import load_wo
from app.core.state import app_state

router = APIRouter(prefix="/upload", tags=["upload"])

_KCI_SHEETS = {
    "dump": "Dump",
    "wo": "WO",
    "mo": "MO",
    "mo_items": "MO Items",
    "so": "SO",
}


def _save_upload(upload: UploadFile, suffix: str) -> Path:
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
        temp_file.write(upload.file.read())
        return Path(temp_file.name)


def _load_sheet_via_loader(
    file_path: Path, sheet_name: str, loader: Callable[[str], list]
) -> list:
    try:
        frame = pd.read_excel(file_path, sheet_name=sheet_name, usecols="D:")
    except ValueError:
        return []
    with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as temp_csv:
        frame.to_csv(temp_csv.name, index=False)
        return loader(temp_csv.name)


@router.post("/kci-excel")
def upload_kci_excel(file: UploadFile = File(...)) -> dict[str, int]:
    if not file.filename:
        raise HTTPException(status_code=400, detail="No file provided.")
    if not file.filename.lower().endswith((".xls", ".xlsx")):
        raise HTTPException(status_code=400, detail="Expected an Excel file.")

    temp_path = _save_upload(file, suffix=Path(file.filename).suffix)
    try:
        dump = _load_sheet_via_loader(temp_path, _KCI_SHEETS["dump"], load_dump)
        wo = _load_sheet_via_loader(temp_path, _KCI_SHEETS["wo"], load_wo)
        mo = _load_sheet_via_loader(temp_path, _KCI_SHEETS["mo"], load_mo)
        mo_items = _load_sheet_via_loader(
            temp_path, _KCI_SHEETS["mo_items"], load_mo_items
        )
        so = _load_sheet_via_loader(temp_path, _KCI_SHEETS["so"], load_so)
    finally:
        temp_path.unlink(missing_ok=True)

    app_state.dump = dump
    app_state.wo = wo
    app_state.mo = mo
    app_state.mo_items = mo_items
    app_state.so = so

    return {
        "dump": len(dump),
        "wo": len(wo),
        "mo": len(mo),
        "mo_items": len(mo_items),
        "so": len(so),
    }


@router.post("/cso-csv")
def upload_cso_csv(file: UploadFile = File(...)) -> dict[str, int]:
    if not file.filename:
        raise HTTPException(status_code=400, detail="No file provided.")
    if not file.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="Expected a CSV file.")

    temp_path = _save_upload(file, suffix=".csv")
    try:
        cso = load_cso(str(temp_path))
    finally:
        temp_path.unlink(missing_ok=True)

    app_state.cso = cso
    return {"cso": len(cso)}


@router.post("/tracking-csv")
def upload_tracking_csv(file: UploadFile = File(...)) -> dict[str, int]:
    if not file.filename:
        raise HTTPException(status_code=400, detail="No file provided.")
    if not file.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="Expected a CSV file.")

    temp_path = _save_upload(file, suffix=".csv")
    try:
        delivery = load_delivery(str(temp_path))
    finally:
        temp_path.unlink(missing_ok=True)

    app_state.delivery = delivery
    return {"delivery": len(delivery)}
