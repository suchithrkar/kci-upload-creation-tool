"""Loader for material order item records."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.schemas.mo_items import MaterialOrderItem


def _schema_fields() -> list[str]:
    fields = getattr(MaterialOrderItem, "model_fields", None)
    if fields is None:
        fields = MaterialOrderItem.__fields__
    return list(fields.keys())


def _read_table(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    if path.suffix.lower() in {".xls", ".xlsx"}:
        return pd.read_excel(path, usecols="D:")
    return pd.read_csv(path)


def load_mo_items(file_path: str) -> list[MaterialOrderItem]:
    """Load material order items from a CSV or Excel file into schema objects."""

    df = _read_table(file_path).fillna("")
    fields = _schema_fields()
    records: list[MaterialOrderItem] = []

    for _, row in df.iterrows():
        values = {field: row.get(field, "") for field in fields}
        if "case_id" in values and not str(values.get("case_id", "")).strip():
            continue

        values = {
            key: value.strip() if isinstance(value, str) else value
            for key, value in values.items()
        }

        records.append(MaterialOrderItem(**values))

    return records
