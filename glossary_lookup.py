"""Parse TMP and Financing fees from the individual-rate Glossary tab."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from carrier_lookup import (
    build_carrier_apply_if,
    carrier_code_from_filename,
    detect_carrier_key,
)
from config import GLOSSARY_TAB, INDIVIDUAL_RATE_SUBFOLDER
from extractor import SubfolderSelection, extract_sheet_to_dataframe, load_processing_context
from thc_lookup import _extract_container_size

TMP_FEE_CODE = "TMP-Fee-ALL_IN"
FINANCING_FEE_CODE = "Financing Fee"
TMP_COST_NAME = "TMP Fee (TMP-Fee-ALL_IN)"
TMP_RATE_BY = "Quantity/Container"

# Glossary TMP section selection is shipper-specific. Currently tuned for
# Siemens Divisions; other shippers need their own markers / parsing rules.
SHIPPER_SECTION_MARKERS: dict[str, str] = {
    "Siemens Divisions": r"Siemens AG incl\. DI, SI, Mobility and SHS",
    "Siemens Healthineers": r"Siemens Healthineers DX Business Line",
    "Siemens Healthineers LATAM": r"Siemens Healthineers DX Business Line",
    "Innomotics": r"Siemens AG incl\. DI, SI, Mobility and SHS",
}

@dataclass(frozen=True)
class TmpFeeEntry:
    key: str
    currency: str
    unit_rate: float
    max_cap: float | None = None
    apply_if_suffix: str = ""


@dataclass(frozen=True)
class FinancingFeeRates:
    currency: str
    rate_20: float
    rate_40: float


@dataclass(frozen=True)
class GlossaryFees:
    carrier_name: str
    tmp_fees: tuple[TmpFeeEntry, ...]
    financing_fee: FinancingFeeRates | None


def _parse_amount(value: str) -> float:
    cleaned = value.strip()
    cleaned = cleaned.replace("€", "").replace("\u20ac", "")
    cleaned = re.sub(r"[^\d,\.]", "", cleaned)
    if not cleaned:
        raise ValueError(f"Could not parse amount from {value!r}")

    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")
    return float(cleaned)


def _parse_currency(*candidates: str | None) -> str:
    for candidate in candidates:
        if not candidate:
            continue
        match = re.search(r"(EUR|USD)", candidate.upper())
        if match:
            return match.group(1).upper()
    return "EUR"


def _truncate_before_lcl(text: str) -> str:
    match = re.search(r"\bLCL\b", text, flags=re.IGNORECASE)
    if not match:
        return text
    return text[: match.start()]


def _extract_lcl_section(text: str) -> str:
    match = re.search(r"\bLCL\b", text, flags=re.IGNORECASE)
    if not match:
        return ""

    section = text[match.start() :]
    next_section = re.search(
        r"\n(?:TMP CONSOL|FCL\b)",
        section[len(match.group(0)) :],
        flags=re.IGNORECASE,
    )
    if next_section:
        section = section[: len(match.group(0)) + next_section.start()]
    return section


def parse_lcl_tmp_fee(text: str | None) -> TmpFeeEntry | None:
    if not text:
        return None

    section = _extract_lcl_section(text)
    if not section:
        return None

    unit_match = re.search(
        r"LCL\s+TMP-fee-all-in[^:\n]*:\s*([\d,\.]+)\s*(?:€|EUR)?",
        section,
        flags=re.IGNORECASE,
    )
    if not unit_match:
        return None

    cap_match = re.search(
        r"CAP:\s*([\d,\.]+)\s*(?:€|EUR)?",
        section,
        flags=re.IGNORECASE,
    )
    return TmpFeeEntry(
        key="lcl_import_export",
        currency="EUR",
        unit_rate=_parse_amount(unit_match.group(1)),
        max_cap=_parse_amount(cap_match.group(1)) if cap_match else None,
    )


def _extract_shipper_section(text: str, shipper: str) -> str:
    marker_pattern = SHIPPER_SECTION_MARKERS.get(shipper)
    if not marker_pattern:
        return _truncate_before_lcl(text)

    marker_match = re.search(marker_pattern, text, flags=re.IGNORECASE)
    if not marker_match:
        return _truncate_before_lcl(text)

    section = text[marker_match.start() :]
    next_section = re.search(
        r"\nSiemens (?:AG incl\.|Healthineers)",
        section[len(marker_match.group(0)) :],
        flags=re.IGNORECASE,
    )
    if next_section:
        section = section[: len(marker_match.group(0)) + next_section.start()]
    return _truncate_before_lcl(section)


def _parse_payment_term_tmp_fees(section: str) -> list[TmpFeeEntry]:
    entries: list[TmpFeeEntry] = []

    blocks = [
        (
            "120_days",
            r"120\s*days[\s\S]*?FCL TMP-fee-all-in[^:\n]*:\s*([\d,\.]+)\s*(?:€|EUR)?[\s\S]*?"
            r"CAP:\s*([\d,\.]+)",
            "Payment term equals 120 days",
        ),
        (
            "less_120_days",
            r"deviating Payment term[^:\n]*:[\s\S]*?FCL TMP-fee-all-in[^:\n]*:\s*([\d,\.]+)\s*(?:€|EUR)?[\s\S]*?"
            r"CAP:\s*([\d,\.]+)",
            "Payment term less than 120 days",
        ),
    ]

    for key, pattern, apply_if_suffix in blocks:
        match = re.search(pattern, section, flags=re.IGNORECASE)
        if not match:
            continue
        unit_rate = _parse_amount(match.group(1))
        max_cap = _parse_amount(match.group(2))
        entries.append(
            TmpFeeEntry(
                key=key,
                currency="EUR",
                unit_rate=unit_rate,
                max_cap=max_cap,
                apply_if_suffix=apply_if_suffix,
            )
        )
    return entries


def _parse_worldwide_tmp_fee(section: str) -> TmpFeeEntry | None:
    match = re.search(
        r"TMP Fee worldwide[\s\S]*?"
        r"(EUR|USD)\s*([\d,\.]+)\s*per Container with a Maximum \(CAP\)\s*(EUR|USD)?\s*([\d,\.]+)",
        section,
        flags=re.IGNORECASE,
    )
    if not match:
        return None

    return TmpFeeEntry(
        key="worldwide",
        currency=_parse_currency(match.group(1), match.group(3)),
        unit_rate=_parse_amount(match.group(2)),
        max_cap=_parse_amount(match.group(4)),
    )


def _parse_region_tmp_fees(section: str) -> list[TmpFeeEntry]:
    entries: list[TmpFeeEntry] = []
    patterns = [
        (
            "usa",
            r"USA:\s*TMP Fee\s*(EUR|USD)\s*([\d,\.]+)\s*[–-]\s*CAP\s*(EUR|USD)?\s*([\d,\.]+)",
            "Destination country equals USA",
        ),
        (
            "row",
            r"RoW:\s*TMP Fee\s*(EUR|USD)\s*([\d,\.]+)\s*[–-]\s*CAP\s*(EUR|USD)?\s*([\d,\.]+)",
            "Destination country not equals USA",
        ),
    ]
    for key, pattern, apply_if_suffix in patterns:
        match = re.search(pattern, section, flags=re.IGNORECASE)
        if not match:
            continue
        entries.append(
            TmpFeeEntry(
                key=key,
                currency=_parse_currency(match.group(1), match.group(3)),
                unit_rate=_parse_amount(match.group(2)),
                max_cap=_parse_amount(match.group(4)),
                apply_if_suffix=apply_if_suffix,
            )
        )
    return entries


def parse_tmp_consol_fee(text: str) -> TmpFeeEntry | None:
    match = re.search(
        r"FCL TMP Consol Fee:\s*([\d,\.]+)\s*(?:€|EUR)?",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        return None
    return TmpFeeEntry(
        key="consol",
        currency="EUR",
        unit_rate=_parse_amount(match.group(1)),
        max_cap=None,
    )


def parse_tmp_fees(text: str, shipper: str) -> list[TmpFeeEntry]:
    section = _extract_shipper_section(text, shipper)

    payment_term_entries = _parse_payment_term_tmp_fees(section)
    if payment_term_entries:
        return payment_term_entries

    worldwide = _parse_worldwide_tmp_fee(section)
    if worldwide:
        return [worldwide]

    region_entries = _parse_region_tmp_fees(section)
    if region_entries:
        return region_entries

    return []


def parse_financing_fee(text: str) -> FinancingFeeRates | None:
    explicit_20 = re.search(
        r"20.?[\s'\"]*Container.*?([\d,\.]+)\s*(USD|EUR)",
        text,
        flags=re.IGNORECASE,
    )
    explicit_40 = re.search(
        r"40.?[\s'\"/]*45.?[\s'\"]*Container.*?([\d,\.]+)\s*(USD|EUR)",
        text,
        flags=re.IGNORECASE,
    )
    if explicit_20 and explicit_40:
        return FinancingFeeRates(
            currency=explicit_20.group(2).upper(),
            rate_20=_parse_amount(explicit_20.group(1)),
            rate_40=_parse_amount(explicit_40.group(1)),
        )

    teu_match = re.search(
        r"RoW:\s*(USD|EUR)\s*([\d,\.]+)\s*per TEU",
        text,
        flags=re.IGNORECASE,
    )
    if teu_match:
        currency = teu_match.group(1).upper()
        rate_20 = _parse_amount(teu_match.group(2))
        return FinancingFeeRates(
            currency=currency,
            rate_20=rate_20,
            rate_40=round(rate_20 * 2, 2),
        )

    return None


def _glossary_fee_text(glossary_df: pd.DataFrame, fee_code: str) -> str | None:
    code_column = glossary_df.iloc[:, 1].astype(str).str.strip()
    matches = glossary_df.loc[code_column.str.casefold() == fee_code.casefold()]
    if matches.empty:
        return None

    row = matches.iloc[0]
    description_parts = [str(value) for value in row.iloc[2:] if pd.notna(value)]
    return "\n".join(description_parts)


class GlossaryFeeLookup:
    def __init__(self, fees: GlossaryFees, carrier_key: str):
        self.fees = fees
        self.carrier_key = carrier_key.lower()

    @property
    def carrier_name(self) -> str:
        return self.fees.carrier_name

    def tmp_fee_column_groups(self) -> list[dict[str, object]]:
        groups: list[dict[str, object]] = []
        for entry in self.fees.tmp_fees:
            groups.append(
                {
                    "cost_name": TMP_COST_NAME,
                    "apply_if": build_carrier_apply_if(
                        self.fees.carrier_name,
                        entry.apply_if_suffix,
                    ),
                    "rate_by": TMP_RATE_BY,
                    "column_count": 3,
                    "rate_types": ["p/unit", "Flat"],
                    "row3_flat_label": "MAX",
                }
            )
        return groups

    def financing_fee_column_groups(self, container_types: list[str]) -> list[dict[str, str]]:
        if self.fees.financing_fee is None:
            return []

        apply_if = build_carrier_apply_if(self.fees.carrier_name)
        groups: list[dict[str, str]] = []
        for container_type in container_types:
            groups.append(
                {
                    "container_type": container_type,
                    "cost_name": f"Financing Fee ({container_type})",
                    "apply_if": apply_if,
                    "rate_by": f"Container/{container_type}",
                    "rate_type": "p/unit",
                    "column_count": 2,
                }
            )
        return groups

    def financing_fee_rate(self, container_type_code: str) -> float | None:
        if self.fees.financing_fee is None:
            return None
        size = _extract_container_size(container_type_code)
        if size == 20:
            return self.fees.financing_fee.rate_20
        return self.fees.financing_fee.rate_40

    def financing_fee_currency(self) -> str | None:
        if self.fees.financing_fee is None:
            return None
        return self.fees.financing_fee.currency


def parse_glossary_fees(
    glossary_df: pd.DataFrame,
    *,
    shipper: str,
    carrier_name: str,
) -> GlossaryFees:
    tmp_text = _glossary_fee_text(glossary_df, TMP_FEE_CODE)
    financing_text = _glossary_fee_text(glossary_df, FINANCING_FEE_CODE)

    tmp_fees = parse_tmp_fees(tmp_text, shipper) if tmp_text else []
    financing_fee = parse_financing_fee(financing_text) if financing_text else None

    return GlossaryFees(
        carrier_name=carrier_name,
        tmp_fees=tuple(tmp_fees),
        financing_fee=financing_fee,
    )


def load_glossary_fee_lookup(
    shipper: str,
    source_file: Path | None = None,
    flow: str = "FCL",
) -> GlossaryFeeLookup | None:
    if source_file is None or not source_file.exists():
        return None

    carrier_name = carrier_code_from_filename(
        source_file.name,
        shipper=shipper,
        flow=flow,
    )
    if not carrier_name:
        return None

    carrier_key = detect_carrier_key(source_file.name, shipper=shipper, flow=flow)
    if not carrier_key:
        return None

    glossary_df = extract_sheet_to_dataframe(source_file, GLOSSARY_TAB)
    fees = parse_glossary_fees(glossary_df, shipper=shipper, carrier_name=carrier_name)
    return GlossaryFeeLookup(fees, carrier_key=carrier_key)


def load_glossary_fee_lookups(
    shipper: str,
    individual_selections: list[SubfolderSelection],
    flow: str = "FCL",
) -> list[GlossaryFeeLookup]:
    lookups: list[GlossaryFeeLookup] = []
    for selection in individual_selections:
        if GLOSSARY_TAB not in selection.tabs:
            continue
        lookup = load_glossary_fee_lookup(
            shipper,
            source_file=selection.file_path,
            flow=flow,
        )
        if lookup is not None:
            lookups.append(lookup)
    return lookups


def individual_rate_selections_from_context(
    context: dict | None = None,
) -> list[SubfolderSelection]:
    processing_context = context or load_processing_context()
    if not processing_context:
        return []

    selections: list[SubfolderSelection] = []
    for selection in processing_context.get("selections", []):
        if selection.get("subfolder") != INDIVIDUAL_RATE_SUBFOLDER:
            continue
        selections.append(
            SubfolderSelection(
                subfolder=selection["subfolder"],
                file_path=Path(selection["file_path"]),
                tabs=selection["tabs"],
            )
        )
    return selections
