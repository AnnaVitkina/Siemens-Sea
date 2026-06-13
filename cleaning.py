"""Data cleaning rules applied after sheet extraction."""

from __future__ import annotations

import pandas as pd


def _is_missing_value(value: object) -> bool:
    if pd.isna(value):
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def _find_cost_column(df: pd.DataFrame) -> str:
    columns = list(df.columns)
    if "Cost" in columns:
        return "Cost"

    for index, column in enumerate(columns):
        if str(column).strip() == "Currency Freight Rate" and index + 1 < len(columns):
            return columns[index + 1]

    raise KeyError(
        "Could not find 'Cost' column next to 'Currency Freight Rate' in DIGI_FCL_Rates."
    )


def clean_digi_fcl_rates(df: pd.DataFrame) -> pd.DataFrame:
    cost_column = _find_cost_column(df)
    mask = df[cost_column].apply(_is_missing_value)
    return df.loc[~mask].reset_index(drop=True)


SHEET_CLEANERS = {
    "DIGI_FCL_Rates": clean_digi_fcl_rates,
}


def clean_sheet_dataframe(sheet_name: str, df: pd.DataFrame) -> pd.DataFrame:
    cleaner = SHEET_CLEANERS.get(sheet_name)
    if cleaner is None:
        return df
    return cleaner(df)
