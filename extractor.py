"""Rate card data extraction utilities."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from carrier_lookup import detect_carrier_key
from cleaning import clean_sheet_dataframe
from config import (
    EXCEL_EXTENSIONS,
    FLOW_OPTIONAL_TABS,
    FLOW_TAB_PRESETS,
    ADDON_SMF_TAB,
    GLOSSARY_TAB,
    INDIVIDUAL_RATE_SUBFOLDER,
    INPUT_DIR,
    INPUT_SUBFOLDERS,
    PROCESSING_CONTEXT_FILE,
    PROCESSING_DIR,
    SHEET_HEADER_ROWS,
)


@dataclass
class SubfolderSelection:
    subfolder: str
    file_path: Path
    tabs: list[str]


@dataclass
class ProcessingContext:
    shipper: str
    flow: str
    underflow: str | None
    selections: list[SubfolderSelection]
    output_path: Path
    created_at: str


def list_input_subfolders() -> list[str]:
    if not INPUT_DIR.exists():
        return []
    return [
        name
        for name in INPUT_SUBFOLDERS
        if (INPUT_DIR / name).is_dir()
    ]


def list_excel_files(subfolder: str) -> list[Path]:
    folder = INPUT_DIR / subfolder
    if not folder.is_dir():
        return []
    files = [
        path
        for path in sorted(folder.iterdir())
        if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS
    ]
    return files


def read_workbook_sheet_names(file_path: Path) -> list[str]:
    suffix = file_path.suffix.lower()
    if suffix == ".xlsb":
        from pyxlsb import open_workbook

        with open_workbook(file_path) as workbook:
            return list(workbook.sheets)
    return pd.ExcelFile(file_path).sheet_names


def resolve_sheet_header_row(sheet_name: str) -> int | None:
    if sheet_name in SHEET_HEADER_ROWS:
        return SHEET_HEADER_ROWS[sheet_name]

    from lcl_rate_card_builder import is_lcl_rates_tab

    if is_lcl_rates_tab(sheet_name):
        return 0
    return None


def resolve_default_tabs(flow: str, subfolder: str, file_path: Path) -> list[str]:
    if flow == "LCL" and subfolder == INDIVIDUAL_RATE_SUBFOLDER:
        from lcl_rate_card_builder import resolve_lcl_tabs

        return resolve_lcl_tabs(file_path)

    presets = FLOW_TAB_PRESETS.get(flow, {}).get(subfolder, [])
    available = set(read_workbook_sheet_names(file_path))
    tabs = [tab for tab in presets if tab in available]

    for tab in FLOW_OPTIONAL_TABS.get(flow, {}).get(subfolder, []):
        if tab in available and tab not in tabs:
            tabs.append(tab)

    return tabs


def _read_excel_engine(file_path: Path) -> str | None:
    if file_path.suffix.lower() == ".xlsb":
        return "pyxlsb"
    return None


def extract_sheet_to_dataframe(file_path: Path, sheet_name: str) -> pd.DataFrame:
    from lcl_rate_card_builder import is_lcl_rates_tab, read_lcl_rates_tab_dataframe

    if is_lcl_rates_tab(sheet_name):
        return read_lcl_rates_tab_dataframe(file_path, sheet_name)

    header_row = resolve_sheet_header_row(sheet_name)
    engine = _read_excel_engine(file_path)

    if header_row is None:
        return pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            header=None,
            engine=engine,
        )

    return pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=header_row,
        engine=engine,
    )


def sanitize_excel_sheet_name(name: str, used_names: set[str]) -> str:
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", name).strip()
    if not cleaned:
        cleaned = "Sheet"

    base = cleaned[:31]
    candidate = base
    counter = 1
    while candidate in used_names:
        suffix = f"_{counter}"
        candidate = f"{base[: 31 - len(suffix)]}{suffix}"
        counter += 1

    used_names.add(candidate)
    return candidate


def slugify(value: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9]+", "_", value).strip("_")
    return slug or "unknown"


def build_output_path(flow: str, shipper: str) -> Path:
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    flow_slug = slugify(flow)
    shipper_slug = slugify(shipper)
    return PROCESSING_DIR / f"{flow_slug}_{shipper_slug}_extracted_{timestamp}.xlsx"


def resolve_output_tab_name(
    tab: str,
    selection: SubfolderSelection,
    shipper: str,
    flow: str,
) -> str:
    if selection.subfolder != INDIVIDUAL_RATE_SUBFOLDER:
        return tab

    carrier_key = detect_carrier_key(
        selection.file_path.name,
        shipper=shipper,
        flow=flow,
    )
    if not carrier_key:
        return tab

    if tab == GLOSSARY_TAB:
        return f"Glossary_{carrier_key}"
    if tab == ADDON_SMF_TAB:
        return f"Add-on SMF (FCL)_{carrier_key}"

    from lcl_rate_card_builder import is_lcl_rates_tab

    if flow == "LCL" and is_lcl_rates_tab(tab):
        if carrier_key:
            return f"LCL_{carrier_key}_{tab}"
        return f"LCL_{tab}"
    return tab


def _selection_to_dict(selection: SubfolderSelection) -> dict:
    return {
        "subfolder": selection.subfolder,
        "file_path": str(selection.file_path),
        "file_name": selection.file_path.name,
        "tabs": selection.tabs,
    }


def save_processing_context(context: ProcessingContext) -> Path:
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "shipper": context.shipper,
        "flow": context.flow,
        "underflow": context.underflow,
        "output_path": str(context.output_path),
        "created_at": context.created_at,
        "selections": [_selection_to_dict(selection) for selection in context.selections],
    }
    PROCESSING_CONTEXT_FILE.write_text(
        json.dumps(payload, indent=2),
        encoding="utf-8",
    )
    return PROCESSING_CONTEXT_FILE


def load_processing_context() -> dict | None:
    if not PROCESSING_CONTEXT_FILE.exists():
        return None
    return json.loads(PROCESSING_CONTEXT_FILE.read_text(encoding="utf-8"))


def save_selections_to_excel(
    flow: str,
    shipper: str,
    selections: list[SubfolderSelection],
    underflow: str | None = None,
    output_path: Path | None = None,
) -> tuple[Path, ProcessingContext]:
    created_at = datetime.now().isoformat(timespec="seconds")
    if output_path is None:
        output_path = build_output_path(flow, shipper)

    used_sheet_names: set[str] = set()

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for selection in selections:
            for tab in selection.tabs:
                df = extract_sheet_to_dataframe(selection.file_path, tab)
                df = clean_sheet_dataframe(tab, df)
                output_tab = sanitize_excel_sheet_name(
                    resolve_output_tab_name(tab, selection, shipper, flow),
                    used_sheet_names,
                )
                df.to_excel(writer, sheet_name=output_tab, index=False)

    context = ProcessingContext(
        shipper=shipper,
        flow=flow,
        underflow=underflow,
        selections=selections,
        output_path=output_path,
        created_at=created_at,
    )
    save_processing_context(context)
    return output_path, context
