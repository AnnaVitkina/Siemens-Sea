"""Lookup helpers for FCL_THC fee values."""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from config import FCL_THC_BASE_TAB
from extractor import extract_sheet_to_dataframe, list_excel_files, load_processing_context

THC_SIZE_PATTERN = re.compile(r"THC\s+(\d+)'", re.IGNORECASE)
THC_CURRENCY_PATTERN = re.compile(r"(EUR|USD)", re.IGNORECASE)


def _normalize_country_code(value: object) -> str | None:
    if pd.isna(value):
        return None
    code = str(value).strip().upper().rstrip("*")
    return code or None


def _extract_container_size(container_type_code: str) -> int:
    parts = str(container_type_code).split("_")
    if parts and parts[0].isdigit():
        size = int(parts[0])
        if size >= 40:
            return 40
        return 20
    return 20


def _parse_thc_columns(columns: list[object]) -> dict[tuple[int, str], str]:
    mapping: dict[tuple[int, str], str] = {}
    for column in columns:
        column_name = str(column)
        size_match = THC_SIZE_PATTERN.search(column_name)
        currency_match = THC_CURRENCY_PATTERN.search(column_name)
        if not size_match or not currency_match:
            continue
        size = int(size_match.group(1))
        currency = currency_match.group(1).upper()
        mapping[(size, currency)] = column_name
    return mapping


def _row_has_thc_values(row: pd.Series, thc_columns: dict[tuple[int, str], str]) -> bool:
    for column_name in thc_columns.values():
        if pd.notna(row.get(column_name)):
            return True
    return False


def _extract_row_values(
    row: pd.Series,
    thc_columns: dict[tuple[int, str], str],
) -> dict[int, dict[str, float]]:
    values: dict[int, dict[str, float]] = {}
    for (size, currency), column_name in thc_columns.items():
        raw_value = row.get(column_name)
        if pd.notna(raw_value):
            values.setdefault(size, {})[currency] = float(raw_value)
    return values


class FclThcLookup:
    def __init__(self, country_values: dict[str, dict[int, dict[str, float]]]):
        self.country_values = country_values

    @classmethod
    def from_dataframe(cls, df: pd.DataFrame) -> FclThcLookup:
        country_column = "2-letter code Country"
        if country_column not in df.columns:
            raise KeyError(f"FCL_THC is missing required column: {country_column}")

        thc_columns = _parse_thc_columns(list(df.columns))
        if not thc_columns:
            raise KeyError("Could not identify THC value columns in FCL_THC.")

        working = df.copy()
        working["Region Group"] = working["Region"].replace("", pd.NA).ffill()

        country_values: dict[str, dict[int, dict[str, float]]] = {}
        for _, region_rows in working.groupby("Region Group", sort=False):
            template_values: dict[int, dict[str, float]] | None = None
            for _, row in region_rows.iterrows():
                country = _normalize_country_code(row[country_column])
                if not country:
                    continue

                row_values = _extract_row_values(row, thc_columns)
                if row_values:
                    country_values[country] = row_values
                    template_values = row_values
                elif template_values is not None:
                    country_values[country] = template_values

        return cls(country_values)

    def lookup(
        self,
        country: object,
        container_type_code: str,
        currency: object,
    ) -> float | None:
        country_code = _normalize_country_code(country)
        currency_code = str(currency).strip().upper() if pd.notna(currency) else None
        if not country_code or currency_code not in {"EUR", "USD"}:
            return None

        size = _extract_container_size(container_type_code)
        country_rates = self.country_values.get(country_code)
        if not country_rates:
            return None

        size_rates = country_rates.get(size)
        if not size_rates:
            return None

        value = size_rates.get(currency_code)
        if value is None:
            return None
        return value


def load_fcl_thc_lookup(
    processing_path: Path | None = None,
    source_file: Path | None = None,
) -> FclThcLookup:
    if processing_path and processing_path.exists():
        sheet_names = pd.ExcelFile(processing_path).sheet_names
        if FCL_THC_BASE_TAB in sheet_names:
            df = pd.read_excel(processing_path, sheet_name=FCL_THC_BASE_TAB)
            return FclThcLookup.from_dataframe(df)

    if source_file and source_file.exists():
        df = extract_sheet_to_dataframe(source_file, FCL_THC_BASE_TAB)
        return FclThcLookup.from_dataframe(df)

    context = load_processing_context()
    if context:
        context_path = Path(context["output_path"])
        if context_path.exists() and FCL_THC_BASE_TAB in pd.ExcelFile(context_path).sheet_names:
            df = pd.read_excel(context_path, sheet_name=FCL_THC_BASE_TAB)
            return FclThcLookup.from_dataframe(df)

    thc_files = list_excel_files("THC fee")
    for file_path in thc_files:
        sheet_names = pd.ExcelFile(file_path).sheet_names
        if FCL_THC_BASE_TAB in sheet_names:
            df = extract_sheet_to_dataframe(file_path, FCL_THC_BASE_TAB)
            return FclThcLookup.from_dataframe(df)

    raise FileNotFoundError(
        "Could not locate FCL_THC data. Run main.py extraction first "
        "or place a THC fee file in the input folder."
    )
