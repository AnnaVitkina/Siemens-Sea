"""Lookup helpers for IMO, EU ETS, Emergency Bunker, and War Risk surcharges."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from config import RATES_BASE_TAB, RATES_REEFER_TAB
from extractor import extract_sheet_to_dataframe, list_excel_files, load_processing_context
from thc_lookup import _extract_container_size


def _is_rf_container(container_type_code: str) -> bool:
    code = str(container_type_code).upper()
    return "_RF" in code or code.endswith("RF")


def is_line_id_column(column: object) -> bool:
    return _normalize_column_name(column).lower() == "line id"


def _format_numeric_line_id(value: int) -> str:
    """Recover leading zeros when Excel already coerced Line ID to a number."""
    if value < 0:
        return str(value)
    digits = str(value)
    if value < 1000:
        return digits.zfill(4)
    if value < 100000:
        return digits.zfill(5)
    return digits


def format_line_id(value: object) -> str | None:
    """Preserve Line ID text from source (4 or 5 digits, with leading zeros)."""
    if pd.isna(value):
        return None
    if isinstance(value, bool):
        return str(value)

    if isinstance(value, int):
        return _format_numeric_line_id(value)
    if isinstance(value, float):
        if float(value).is_integer():
            return _format_numeric_line_id(int(value))
        return str(value).strip()

    text = str(value).strip()
    if not text:
        return None
    if text.endswith(".0") and text[:-2].isdigit():
        return _format_numeric_line_id(int(text[:-2]))
    return text


def line_id_read_converters(columns: list[object]) -> dict[object, type]:
    return {column: str for column in columns if is_line_id_column(column)}


def read_excel_preserving_line_id(
    file_path: Path,
    sheet_name: str,
    *,
    header: int | None = 0,
    engine: str | None = None,
) -> pd.DataFrame:
    preview = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=header,
        nrows=0,
        engine=engine,
    )
    converters = line_id_read_converters(list(preview.columns))
    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=header,
        engine=engine,
        converters=converters or None,
    )
    return apply_line_id_formatting(df)


def apply_line_id_formatting(
    df: pd.DataFrame,
    columns: list[str] | tuple[str, ...] | None = None,
) -> pd.DataFrame:
    if df.empty:
        return df
    result = df.copy()
    target_columns = list(columns) if columns is not None else list(result.columns)
    for column in target_columns:
        if column not in result.columns:
            continue
        if _normalize_column_name(column).lower() != "line id":
            continue
        result[column] = result[column].apply(format_line_id)
    return result


def _normalize_line_id(value: object) -> str | None:
    return format_line_id(value)


def _normalize_column_name(column: object) -> str:
    return " ".join(str(column).split())


def _war_risk_column_key(container_type_code: str, *, reefer_sheet: bool) -> str:
    code = str(container_type_code).upper()
    size = _extract_container_size(container_type_code)

    if reefer_sheet:
        return "wrs_20_rf" if size == 20 else "wrs_40_rf"

    if code == "20_GP":
        return "wrs_20_dry"
    if code == "40_GP":
        return "wrs_40_dry"
    return "wrs_20_special" if size == 20 else "wrs_40_special"


def _find_surcharge_columns(
    columns: list[object],
    *,
    reefer: bool = False,
    strict: bool = True,
) -> dict[str, str]:
    normalized = {_normalize_column_name(column): column for column in columns}
    mapping: dict[str, str] = {}

    for key, column in normalized.items():
        key_upper = key.upper()
        if "IMO SURCHARGE" in key_upper and "20" in key_upper and "TYPES" in key_upper:
            mapping["imo_20"] = column
        elif "IMO SURCHARGE" in key_upper and "40" in key_upper and "TYPES" in key_upper:
            mapping["imo_40"] = column
        elif "EU ETS" in key_upper and "20" in key_upper:
            mapping["ets_20"] = column
        elif "EU ETS" in key_upper and "40" in key_upper:
            mapping["ets_40"] = column
        elif "EMERGENCY BUNKER SURCHARGE - AMOUNT" in key_upper and "20" in key_upper:
            mapping["ebs_20"] = column
        elif "EMERGENCY BUNKER SURCHARGE - AMOUNT" in key_upper and "40" in key_upper:
            mapping["ebs_40"] = column
        elif "IMO SURCHARGE" in key_upper and "CURRENCY" in key_upper:
            mapping["imo_currency"] = column
        elif "EU ETS" in key_upper and "CURRENCY" in key_upper:
            mapping["ets_currency"] = column
        elif "EMERGENCY BUNKER SURCHARGE - CURRENCY" in key_upper:
            mapping["ebs_currency"] = column
        elif "WAR RISK" in key_upper and "CURRENCY" in key_upper:
            mapping["wrs_currency"] = column
        elif "FREIGHT RATE CURRENCY" in key_upper:
            mapping["freight_currency"] = column
        elif not reefer and "WAR RISK" in key_upper and "20" in key_upper and "DRY" in key_upper:
            mapping["wrs_20_dry"] = column
        elif not reefer and "WAR RISK" in key_upper and "40" in key_upper and "DRY" in key_upper:
            mapping["wrs_40_dry"] = column
        elif not reefer and "WAR RISK" in key_upper and "20" in key_upper and "SPECIAL" in key_upper:
            mapping["wrs_20_special"] = column
        elif not reefer and "WAR RISK" in key_upper and "40" in key_upper and "SPECIAL" in key_upper:
            mapping["wrs_40_special"] = column
        elif reefer and "WAR RISK" in key_upper and "20" in key_upper and "RF" in key_upper:
            mapping["wrs_20_rf"] = column
        elif reefer and "WAR RISK" in key_upper and "40" in key_upper and "RF" in key_upper:
            mapping["wrs_40_rf"] = column

    required = {
        "imo_20",
        "imo_40",
        "ets_20",
        "ets_40",
        "ebs_20",
        "ebs_40",
        "imo_currency",
        "ets_currency",
        "ebs_currency",
        "wrs_currency",
        "freight_currency",
    }
    if reefer:
        required |= {"wrs_20_rf", "wrs_40_rf"}
    else:
        required |= {"wrs_20_dry", "wrs_40_dry", "wrs_20_special", "wrs_40_special"}

    missing = required - set(mapping)
    if strict and missing:
        sheet_label = "Rates_Reefer_Containers" if reefer else "Rates"
        raise KeyError(
            f"Could not find surcharge columns in {sheet_label}: {', '.join(sorted(missing))}"
        )

    return mapping


def _extract_imo_ets_values(
    row: pd.Series,
    surcharge_columns: dict[str, str],
    container_type_code: str,
) -> tuple[float | None, float | None]:
    size = _extract_container_size(container_type_code)
    size_key = "20" if size == 20 else "40"

    imo_column = surcharge_columns.get(f"imo_{size_key}")
    ets_column = surcharge_columns.get(f"ets_{size_key}")
    imo_value = row.get(imo_column) if imo_column else None
    ets_value = row.get(ets_column) if ets_column else None

    imo = float(imo_value) if pd.notna(imo_value) else None
    ets = float(ets_value) if pd.notna(ets_value) else None
    return imo, ets


def _extract_ebs_value(
    row: pd.Series,
    surcharge_columns: dict[str, str],
    container_type_code: str,
) -> float | None:
    size = _extract_container_size(container_type_code)
    size_key = "20" if size == 20 else "40"
    ebs_column = surcharge_columns.get(f"ebs_{size_key}")
    if not ebs_column:
        return None
    ebs_value = row.get(ebs_column)
    if pd.notna(ebs_value):
        return float(ebs_value)
    return None


def _extract_wrs_value(
    row: pd.Series,
    surcharge_columns: dict[str, str],
    container_type_code: str,
    *,
    reefer_sheet: bool,
) -> float | None:
    column_key = _war_risk_column_key(container_type_code, reefer_sheet=reefer_sheet)
    wrs_column = surcharge_columns.get(column_key)
    if not wrs_column:
        return None
    wrs_value = row.get(wrs_column)
    if pd.notna(wrs_value):
        return float(wrs_value)
    return None


def _lookup_row_currency(
    row: pd.Series,
    surcharge_columns: dict[str, str],
    currency_key: str,
) -> str | None:
    currency_column = surcharge_columns.get(currency_key)
    currency_value = row.get(currency_column) if currency_column else None
    if pd.notna(currency_value):
        return str(currency_value).strip().upper()

    freight_column = surcharge_columns.get("freight_currency")
    freight_currency = row.get(freight_column) if freight_column else None
    if pd.notna(freight_currency):
        return str(freight_currency).strip().upper()
    return None


class RatesSurchargeLookup:
    def __init__(
        self,
        rates_rows: dict[str, pd.Series],
        reefer_rows: dict[str, pd.Series],
        rates_columns: dict[str, str],
        reefer_columns: dict[str, str],
    ):
        self.rates_rows = rates_rows
        self.reefer_rows = reefer_rows
        self.rates_columns = rates_columns
        self.reefer_columns = reefer_columns

    @classmethod
    def from_dataframes(
        cls,
        rates_df: pd.DataFrame,
        reefer_df: pd.DataFrame,
        *,
        strict: bool = True,
    ) -> RatesSurchargeLookup:
        rates_columns = _find_surcharge_columns(list(rates_df.columns), reefer=False, strict=strict)
        reefer_columns = _find_surcharge_columns(list(reefer_df.columns), reefer=True, strict=strict)

        rates_rows: dict[str, pd.Series] = {}
        for _, row in rates_df.iterrows():
            line_id = _normalize_line_id(row.get("Line ID"))
            if line_id:
                rates_rows[line_id] = row

        reefer_rows: dict[str, pd.Series] = {}
        for _, row in reefer_df.iterrows():
            line_id = _normalize_line_id(row.get("Line ID"))
            if line_id:
                reefer_rows[line_id] = row

        return cls(rates_rows, reefer_rows, rates_columns, reefer_columns)

    def _resolve_row(
        self,
        line_id: object,
        container_type_code: str,
    ) -> tuple[pd.Series | None, dict[str, str], bool]:
        normalized_line_id = _normalize_line_id(line_id)
        if not normalized_line_id:
            return None, {}, False

        rf_container = _is_rf_container(container_type_code)
        if normalized_line_id.startswith("R"):
            if not rf_container:
                return None, {}, False
            return self.reefer_rows.get(normalized_line_id), self.reefer_columns, True

        if rf_container:
            return self.reefer_rows.get(normalized_line_id), self.reefer_columns, True
        return self.rates_rows.get(normalized_line_id), self.rates_columns, False

    def _resolve_imo_ets_row(
        self,
        line_id: object,
        container_type_code: str,
    ) -> tuple[pd.Series | None, dict[str, str], bool]:
        return self._resolve_row(line_id, container_type_code)

    def lookup_imo(self, line_id: object, container_type_code: str) -> float | None:
        row, columns, _reefer_sheet = self._resolve_imo_ets_row(line_id, container_type_code)
        if row is None:
            return None
        imo, _ = _extract_imo_ets_values(row, columns, container_type_code)
        return imo

    def lookup_ets(self, line_id: object, container_type_code: str) -> float | None:
        row, columns, _reefer_sheet = self._resolve_imo_ets_row(line_id, container_type_code)
        if row is None:
            return None
        _, ets = _extract_imo_ets_values(row, columns, container_type_code)
        return ets

    def lookup_ebs(self, line_id: object, container_type_code: str) -> float | None:
        row, columns, _reefer_sheet = self._resolve_row(line_id, container_type_code)
        if row is None:
            return None
        return _extract_ebs_value(row, columns, container_type_code)

    def lookup_wrs(self, line_id: object, container_type_code: str) -> float | None:
        row, columns, reefer_sheet = self._resolve_row(line_id, container_type_code)
        if row is None:
            return None
        return _extract_wrs_value(
            row,
            columns,
            container_type_code,
            reefer_sheet=reefer_sheet,
        )

    def lookup_imo_currency(self, line_id: object, container_type_code: str) -> str | None:
        row, columns, _reefer_sheet = self._resolve_row(line_id, container_type_code)
        if row is None:
            return None
        return _lookup_row_currency(row, columns, "imo_currency")

    def lookup_ets_currency(self, line_id: object, container_type_code: str) -> str | None:
        row, columns, _reefer_sheet = self._resolve_row(line_id, container_type_code)
        if row is None:
            return None
        return _lookup_row_currency(row, columns, "ets_currency")

    def lookup_ebs_currency(self, line_id: object, container_type_code: str) -> str | None:
        row, columns, _reefer_sheet = self._resolve_row(line_id, container_type_code)
        if row is None:
            return None
        return _lookup_row_currency(row, columns, "ebs_currency")

    def lookup_wrs_currency(self, line_id: object, container_type_code: str) -> str | None:
        row, columns, _reefer_sheet = self._resolve_row(line_id, container_type_code)
        if row is None:
            return None
        return _lookup_row_currency(row, columns, "wrs_currency")


def load_rates_surcharge_lookup(
    processing_path: Path | None = None,
    source_file: Path | None = None,
    *,
    strict: bool = True,
) -> RatesSurchargeLookup:
    if processing_path and processing_path.exists():
        rates_df = read_excel_preserving_line_id(
            processing_path,
            RATES_BASE_TAB,
            header=0,
        )
        reefer_df = read_excel_preserving_line_id(
            processing_path,
            RATES_REEFER_TAB,
            header=0,
        )
        return RatesSurchargeLookup.from_dataframes(rates_df, reefer_df, strict=strict)

    if source_file and source_file.exists():
        rates_df = extract_sheet_to_dataframe(source_file, RATES_BASE_TAB)
        reefer_df = extract_sheet_to_dataframe(source_file, RATES_REEFER_TAB)
        return RatesSurchargeLookup.from_dataframes(rates_df, reefer_df, strict=strict)

    context = load_processing_context()
    if context:
        context_path = Path(context["output_path"])
        if context_path.exists():
            sheet_names = pd.ExcelFile(context_path).sheet_names
            if RATES_BASE_TAB in sheet_names and RATES_REEFER_TAB in sheet_names:
                rates_df = read_excel_preserving_line_id(
                    context_path,
                    RATES_BASE_TAB,
                    header=0,
                )
                reefer_df = read_excel_preserving_line_id(
                    context_path,
                    RATES_REEFER_TAB,
                    header=0,
                )
                return RatesSurchargeLookup.from_dataframes(rates_df, reefer_df, strict=strict)

    main_rate_files = list_excel_files("main rates")
    for file_path in main_rate_files:
        sheet_names = pd.ExcelFile(file_path).sheet_names
        if RATES_BASE_TAB in sheet_names and RATES_REEFER_TAB in sheet_names:
            rates_df = extract_sheet_to_dataframe(file_path, RATES_BASE_TAB)
            reefer_df = extract_sheet_to_dataframe(file_path, RATES_REEFER_TAB)
            return RatesSurchargeLookup.from_dataframes(rates_df, reefer_df, strict=strict)

    raise FileNotFoundError(
        "Could not locate Rates / Rates_Reefer_Containers data. Run main.py extraction first "
        "or place a main rates file in the input folder."
    )
