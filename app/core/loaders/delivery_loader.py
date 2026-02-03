"""Loader for delivery status records."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.schemas.delivery import DeliveryStatus


def _schema_fields() -> list[str]:
    fields = getattr(DeliveryStatus, "model_fields", None)
    if fields is None:
        fields = DeliveryStatus.__fields__
    return list(fields.keys())


def _read_table(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    if path.suffix.lower() in {".xls", ".xlsx"}:
        return pd.read_excel(path, usecols="D:")
    return pd.read_csv(path)


def load_delivery(file_path: str) -> list[DeliveryStatus]:
    """Load delivery statuses from a CSV or Excel file into schema objects."""

    df = _read_table(file_path).fillna("")
    fields = _schema_fields()
    records: list[DeliveryStatus] = []

    for _, row in df.iterrows():
        values = {field: row.get(field, "") for field in fields}
        if not str(values.get("case_id", "")).strip():
            continue

        values = {
            key: value.strip() if isinstance(value, str) else value
            for key, value in values.items()
        }

        records.append(DeliveryStatus(**values))

    return records
