"""Loader for dump case records."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from app.schemas.dump import DumpCase


_DATE_COLUMNS = {"created_on"}


def _schema_fields() -> list[str]:
    fields = getattr(DumpCase, "model_fields", None)
    if fields is None:
        fields = DumpCase.__fields__
    return list(fields.keys())


def _parse_datetime(value: Any) -> Any:
    if value == "" or value is None:
        return ""
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)):
        result = pd.to_datetime(value, unit="d", origin="1899-12-30")
    else:
        result = pd.to_datetime(value, errors="coerce")
    if pd.isna(result):
        return ""
    return result


def _read_table(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    if path.suffix.lower() in {".xls", ".xlsx"}:
        return pd.read_excel(path, usecols="D:")
    return pd.read_csv(path)


def load_dump(file_path: str) -> list[DumpCase]:
    """Load dump cases from a CSV or Excel file into schema objects."""

    df = _read_table(file_path).fillna("")
    fields = _schema_fields()
    records: list[DumpCase] = []

    for _, row in df.iterrows():
        values = {field: row.get(field, "") for field in fields}
        if not str(values.get("case_id", "")).strip():
            continue

        for column in _DATE_COLUMNS:
            values[column] = _parse_datetime(values.get(column, ""))

        values = {
            key: value.strip() if isinstance(value, str) else value
            for key, value in values.items()
        }

        records.append(DumpCase(**values))

    return records
