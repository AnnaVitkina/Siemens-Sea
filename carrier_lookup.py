"""Carrier detection from individual rate card file names."""

from __future__ import annotations

import re
from typing import Any

from config import (
    IMPLEMENTED_FLOW,
    IMPLEMENTED_SHIPPER,
    SHIPPER_CARRIER_CONFIG_BY_FLOW,
)

CARRIER_FILE_PATTERNS = {
    "FCL": re.compile(r"FCL_([A-Z]+)_", re.IGNORECASE),
    "BCN": re.compile(r"BCN_([A-Z]+)_", re.IGNORECASE),
}


def resolve_shipper_name(shipper: str | None = None) -> str:
    if shipper:
        return shipper
    from extractor import load_processing_context

    context = load_processing_context()
    if context and context.get("shipper"):
        return context["shipper"]
    return IMPLEMENTED_SHIPPER


def resolve_flow_name(flow: str | None = None) -> str:
    if flow:
        return flow
    from extractor import load_processing_context

    context = load_processing_context()
    if context and context.get("flow"):
        return context["flow"]
    return IMPLEMENTED_FLOW


def get_shipper_carrier_config(
    shipper: str | None = None,
    flow: str | None = None,
) -> dict[str, Any]:
    resolved_shipper = resolve_shipper_name(shipper)
    resolved_flow = resolve_flow_name(flow)
    flow_config = SHIPPER_CARRIER_CONFIG_BY_FLOW.get(
        resolved_flow,
        SHIPPER_CARRIER_CONFIG_BY_FLOW[IMPLEMENTED_FLOW],
    )
    if resolved_shipper in flow_config:
        return flow_config[resolved_shipper]
    return flow_config[IMPLEMENTED_SHIPPER]


def get_carrier_file_keys(
    shipper: str | None = None,
    flow: str | None = None,
) -> tuple[str, ...]:
    return get_shipper_carrier_config(shipper, flow=flow)["file_keys"]


def get_carrier_codes(
    shipper: str | None = None,
    flow: str | None = None,
) -> dict[str, str]:
    return get_shipper_carrier_config(shipper, flow=flow)["carrier_codes"]


def get_in_scope_carrier_codes(
    shipper: str | None = None,
    flow: str | None = None,
) -> tuple[str, ...]:
    config = get_shipper_carrier_config(shipper, flow=flow)
    if "in_scope_carrier_codes" in config:
        return tuple(config["in_scope_carrier_codes"])
    return tuple(config["carrier_codes"].values())


def get_ebs_carrier_variants(
    shipper: str | None = None,
    flow: str | None = None,
) -> tuple[dict[str, str], ...]:
    return get_shipper_carrier_config(shipper, flow=flow)["ebs_variants"]


def get_biofuel_cost_names(
    shipper: str | None = None,
    flow: str | None = None,
) -> dict[str, str]:
    return get_shipper_carrier_config(shipper, flow=flow).get("biofuel_cost_names", {})


def _filename_matches_carrier_key(
    file_name: str,
    carrier_key: str,
    flow: str | None = None,
) -> bool:
    upper_name = file_name.upper()
    key = carrier_key.upper()
    resolved_flow = resolve_flow_name(flow).upper()
    return (
        f"_{key}_" in upper_name
        or f"{resolved_flow}_{key}_" in upper_name
        or key in upper_name
    )


def detect_carrier_key(
    file_name: str,
    shipper: str | None = None,
    flow: str | None = None,
) -> str | None:
    upper_name = file_name.upper()
    carrier_codes = get_carrier_codes(shipper, flow=flow)
    resolved_flow = resolve_flow_name(flow)

    for carrier_key in get_carrier_file_keys(shipper, flow=flow):
        if _filename_matches_carrier_key(file_name, carrier_key, flow=resolved_flow):
            return carrier_key

    pattern = CARRIER_FILE_PATTERNS.get(resolved_flow, CARRIER_FILE_PATTERNS[IMPLEMENTED_FLOW])
    match = pattern.search(upper_name)
    if match:
        candidate = match.group(1).upper()
        if candidate in carrier_codes:
            return candidate
    return None


def carrier_code_from_key(
    carrier_key: str,
    shipper: str | None = None,
    flow: str | None = None,
) -> str | None:
    return get_carrier_codes(shipper, flow=flow).get(carrier_key.upper())


def carrier_code_from_filename(
    file_name: str,
    shipper: str | None = None,
    flow: str | None = None,
) -> str | None:
    carrier_key = detect_carrier_key(file_name, shipper=shipper, flow=flow)
    if not carrier_key:
        return None
    return carrier_code_from_key(carrier_key, shipper=shipper, flow=flow)


def build_carrier_apply_if(carrier_code: str, suffix: str = "") -> str:
    base = f"Carrier name equals {carrier_code}"
    if suffix:
        return f"{base}; {suffix}"
    return base


def build_ebs_carrier_apply_if(carrier_code: str) -> str:
    return f"Carrier Name equals {carrier_code}"


def build_carrier_name_apply_if(carrier_code: str, suffix: str = "") -> str:
    base = f"Carrier Name equals {carrier_code}"
    if suffix:
        return f"{base}; {suffix}"
    return base


def build_all_carriers_apply_if(
    shipper: str | None = None,
    flow: str | None = None,
) -> str:
    parts = [
        f"Carrier name equals {carrier_code}"
        for carrier_code in get_in_scope_carrier_codes(shipper, flow=flow)
    ]
    return "; ".join(parts)
