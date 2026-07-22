"""Shared date formatting for rate card builders."""

from __future__ import annotations

from datetime import datetime

import pandas as pd

_EXCEL_SERIAL_MIN = 20_000
_EXCEL_SERIAL_MAX = 80_000


def format_dd_mm_yyyy(value: object) -> str | None:
    if pd.isna(value):
        return None

    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")

    if isinstance(value, (int, float)):
        number = float(value)
        if _EXCEL_SERIAL_MIN <= number <= _EXCEL_SERIAL_MAX:
            parsed = pd.to_datetime(number, unit="D", origin="1899-12-30", errors="coerce")
            if pd.notna(parsed):
                return parsed.strftime("%d.%m.%Y")

    text = str(value).strip()
    if not text:
        return None

    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if pd.notna(parsed):
        return parsed.strftime("%d.%m.%Y")

    return text
