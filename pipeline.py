"""End-to-end rate card pipeline: extract input files and build the rate card."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from carrier_lookup import carrier_code_from_filename, detect_carrier_key
from config import GLOSSARY_TAB, IMPLEMENTED_FLOWS, INDIVIDUAL_RATE_SUBFOLDER, RATE_CARD_REQUIRED_TABS_BY_FLOW
from extractor import ProcessingContext, SubfolderSelection, save_selections_to_excel
from extractor import build_flow_result_output_path, primary_source_file_for_output
from glossary_lookup import GlossaryFeeLookup, load_glossary_fee_lookups
from lcl_rate_card_builder import (
    build_output_rate_card_path as build_lcl_output_rate_card_path,
    is_lcl_data_tab,
    is_lcl_rates_tab,
    resolve_lcl_carrier_slug,
    save_lcl_rate_card,
)
from preon_carriage_builder import save_preon_generic_rate_card, save_preon_per_carrier_rate_card
from preon_carriage_builder import (
    PREON_DIESELFLOATER_TAB,
    PREON_GENERIC_ADD_SERVICES_TABS,
    build_output_rate_card_path as build_preon_output_rate_card_path,
    resolve_preon_carrier_slug,
)
from rate_card_builder import (
    build_output_rate_card_path,
    load_digi_fcl_rates_dataframe,
    save_rate_card,
)
from rates_surcharge_lookup import RatesSurchargeLookup, load_rates_surcharge_lookup
from thc_lookup import FclThcLookup, load_fcl_thc_lookup

MAIN_RATES_OPTIONAL_SUPPLEMENT_SHIPPERS = frozenset({
    "Siemens Divisions",
    "Siemens Healthineers",
    "Siemens Healthineers LATAM",
})
MAIN_RATES_OPTIONAL_SUPPLEMENT_FLOWS = frozenset({"FCL", "BCN"})


def _main_rates_optional_supplements(shipper: str, flow: str) -> bool:
    return shipper in MAIN_RATES_OPTIONAL_SUPPLEMENT_SHIPPERS and flow in MAIN_RATES_OPTIONAL_SUPPLEMENT_FLOWS


def _load_thc_lookup_if_available(processing_path: Path) -> FclThcLookup:
    try:
        return load_fcl_thc_lookup(processing_path=processing_path)
    except (FileNotFoundError, KeyError, ValueError):
        return FclThcLookup({})


@dataclass
class CarrierSummary:
    carrier_key: str
    carrier_code: str
    individual_rate_file: str
    tmp_fee_blocks: int = 0
    financing_fee: str | None = None


@dataclass
class PipelineResult:
    shipper: str
    flow: str
    processing_path: Path
    rate_card_path: Path
    source_rows: int
    rate_card_rows: int
    rate_card_columns: int
    carriers: list[CarrierSummary] = field(default_factory=list)


def _selected_tabs(selections: list[SubfolderSelection]) -> set[str]:
    tabs: set[str] = set()
    for selection in selections:
        tabs.update(selection.tabs)
    return tabs


def split_selections(
    selections: list[SubfolderSelection],
) -> tuple[list[SubfolderSelection], list[SubfolderSelection]]:
    shared: list[SubfolderSelection] = []
    individual: list[SubfolderSelection] = []
    for selection in selections:
        if selection.subfolder == INDIVIDUAL_RATE_SUBFOLDER:
            individual.append(selection)
        else:
            shared.append(selection)
    return shared, individual


def validate_rate_card_selections(
    flow: str,
    selections: list[SubfolderSelection],
    shipper: str,
    underflow: str | None = None,
) -> list[str]:
    """Return hard validation errors that prevent running an unsupported flow."""
    del selections, underflow  # Tab and selection gaps are handled as warnings.

    errors: list[str] = []
    if flow not in RATE_CARD_REQUIRED_TABS_BY_FLOW:
        errors.append(f"Flow '{flow}' is not implemented yet.")
    if flow == "Haulage" and shipper != "Siemens Healthineers":
        errors.append("Haulage flow is implemented only for Siemens Healthineers.")
    return errors


def _warn_missing_tabs(
    selection: SubfolderSelection,
    required_tabs: set[str],
    *,
    consequence: str,
) -> list[str]:
    warnings: list[str] = []
    for tab in sorted(required_tabs.difference(selection.tabs)):
        warnings.append(
            f"Tab '{tab}' not selected for {selection.file_path.name} — {consequence}"
        )
    return warnings


def warn_rate_card_selections(
    selections: list[SubfolderSelection],
    flow: str = "FCL",
    underflow: str | None = None,
    shipper: str | None = None,
) -> list[str]:
    if flow not in RATE_CARD_REQUIRED_TABS_BY_FLOW:
        return []

    warnings: list[str] = []
    shared, individual = split_selections(selections)

    if flow in {"FCL", "BCN"}:
        selected = _selected_tabs(shared)
        for tab in RATE_CARD_REQUIRED_TABS_BY_FLOW[flow]:
            if tab not in selected:
                consequence = {
                    "DIGI_FCL_Rates": "transport and base lane data will be omitted.",
                    "Rates": "standard-container surcharges may be omitted.",
                    "Rates_Reefer_Containers": "reefer surcharges may be omitted.",
                    "FCL_THC": "THC lookup values will be omitted (DIGI fallback may still apply).",
                }.get(tab, "related rate card columns may be omitted.")
                warnings.append(f"Tab '{tab}' not selected — {consequence}")
        if not individual and not _main_rates_optional_supplements(shipper or "", flow):
            warnings.append(
                "No individual rate file selected — TMP Fee, Financing Fee, "
                "and related accessorial columns will be omitted."
            )
        for selection in individual:
            if GLOSSARY_TAB not in selection.tabs:
                warnings.append(
                    f"Tab '{GLOSSARY_TAB}' not selected for {selection.file_path.name} — "
                    "TMP Fee and Financing Fee columns will be omitted."
                )
            carrier_key = detect_carrier_key(
                selection.file_path.name,
                shipper=shipper,
                flow=flow,
            )
            if not carrier_key:
                warnings.append(
                    f"Could not detect carrier from individual rate file "
                    f"{selection.file_path.name} — glossary-based columns may be omitted."
                )
        return warnings

    if flow == "Pre/on carriage":
        if underflow == "generic":
            main_rates = [selection for selection in shared if selection.subfolder == "main rates"]
            if not main_rates:
                warnings.append(
                    "No main rates file selected — generic Pre/on carriage build may fail."
                )
            required_tabs = {
                "PreOnCarriage_Containerized_EU",
                "PreOn_Containerized_EU_Services",
                "DIGI_FCL_Rates",
                "HAPAG_Terms & Conditions",
                "MAERSK_Terms & Condition",
                "MSC_Terms & Conditions",
                "ONE_Terms & Conditions",
            }
            tab_consequences = {
                "PreOnCarriage_Containerized_EU": "generic transport rows will be omitted.",
                "PreOn_Containerized_EU_Services": "IMO, positioning, T1, and waiting-fee columns may be omitted.",
                "DIGI_FCL_Rates": "THC Origin columns may be omitted.",
                "HAPAG_Terms & Conditions": "HLCU terms accessorial rows may be omitted.",
                "MAERSK_Terms & Condition": "MAEU terms accessorial rows may be omitted.",
                "MSC_Terms & Conditions": "MSCU terms accessorial rows may be omitted.",
                "ONE_Terms & Conditions": "ONEY terms accessorial rows may be omitted.",
            }
            for selection in main_rates:
                for tab in sorted(required_tabs.difference(selection.tabs)):
                    consequence = tab_consequences.get(
                        tab,
                        "related Pre/on carriage columns may be omitted.",
                    )
                    warnings.append(
                        f"Tab '{tab}' not selected for {selection.file_path.name} — {consequence}"
                    )
                if not any(tab in selection.tabs for tab in PREON_GENERIC_ADD_SERVICES_TABS):
                    warnings.append(
                        f"No Add Services tab selected for {selection.file_path.name} — "
                        "additional services accessorial costs will be omitted."
                    )
            return warnings

        if not individual:
            warnings.append(
                "No individual rate file selected — per-carrier Pre/on carriage build may fail."
            )
        for selection in individual:
            warnings.extend(
                _warn_missing_tabs(
                    selection,
                    {"Pre-On-Carriage_RoW"},
                    consequence="Pre/on carriage transport rows will be omitted.",
                )
            )
            if GLOSSARY_TAB not in selection.tabs:
                warnings.append(
                    f"Tab '{GLOSSARY_TAB}' not selected for {selection.file_path.name} — "
                    "TMP Fee accessorial rows may be omitted."
                )
            if PREON_DIESELFLOATER_TAB not in selection.tabs:
                warnings.append(
                    f"Tab '{PREON_DIESELFLOATER_TAB}' not selected for {selection.file_path.name} — "
                    "Emergency Fuel Surcharge accessorial will be omitted."
                )
        return warnings

    if flow == "Haulage":
        main_rates = [selection for selection in shared if selection.subfolder == "main rates"]
        if not main_rates:
            warnings.append("No main rates file selected — Haulage build may fail.")
        for selection in main_rates:
            warnings.extend(
                _warn_missing_tabs(
                    selection,
                    {
                        "PreOn_Carriage_Car. Haulage",
                        "PreOn_Containerized_EU_Services",
                    },
                    consequence="related Haulage columns may be omitted.",
                )
            )
        return warnings

    if flow == "LCL":
        if not individual:
            warnings.append("No individual rate file selected — LCL build may fail.")
        has_lcl_rates_tab = False
        for selection in individual:
            lcl_tabs = [tab for tab in selection.tabs if is_lcl_data_tab(tab)]
            if not lcl_tabs:
                warnings.append(
                    f"No LCL Rate/LCL_Rates/GST tab selected for {selection.file_path.name} — "
                    "LCL rate rows will be omitted."
                )
            else:
                has_lcl_rates_tab = True
        if individual and not has_lcl_rates_tab:
            warnings.append(
                "No LCL Rate, LCL_Rates, or GST tab selected — LCL build may fail."
            )
        return warnings

    return warnings


def warn_fcl_selections(
    selections: list[SubfolderSelection],
    shipper: str | None = None,
) -> list[str]:
    return warn_rate_card_selections(selections, flow="FCL", shipper=shipper)


def run_extraction(
    flow: str,
    shipper: str,
    selections: list[SubfolderSelection],
    underflow: str | None = None,
) -> ProcessingContext:
    _, context = save_selections_to_excel(flow, shipper, selections, underflow=underflow)
    return context


def run_rate_card_build(
    shipper: str,
    processing_path: Path,
    flow: str = "FCL",
    underflow: str | None = None,
    *,
    glossary_lookups: list[GlossaryFeeLookup] | None = None,
    individual_selections: list[SubfolderSelection] | None = None,
    output_path: Path | None = None,
) -> tuple[Path, object]:
    if flow == "LCL":
        carrier_slug = resolve_lcl_carrier_slug(shipper, individual_selections or [])
        rate_card_path, rate_card_df, _conditional_df = save_lcl_rate_card(
            shipper,
            individual_selections or [],
            processing_path=processing_path,
            output_path=output_path
            or build_lcl_output_rate_card_path("LCL", shipper, carrier_slug),
        )
        return rate_card_path, rate_card_df
    if flow == "Pre/on carriage":
        if underflow == "generic":
            main_rate_selections = [
                selection for selection in (individual_selections or []) if selection.subfolder == "main rates"
            ]
            rate_card_path, rate_card_df = save_preon_generic_rate_card(
                shipper,
                main_rate_selections,
                output_path=output_path or build_preon_output_rate_card_path("Pre_on_carriage_generic", shipper),
            )
            return rate_card_path, rate_card_df

        carrier_slug = resolve_preon_carrier_slug(shipper, individual_selections or [])
        rate_card_path, rate_card_df = save_preon_per_carrier_rate_card(
            shipper,
            individual_selections or [],
            output_path=output_path
            or build_preon_output_rate_card_path("Pre_on_carriage", shipper, carrier_slug),
        )
        return rate_card_path, rate_card_df
    if flow == "Haulage":
        main_rate_selections = [
            selection for selection in (individual_selections or []) if selection.subfolder == "main rates"
        ]
        rate_card_path, rate_card_df = save_preon_generic_rate_card(
            shipper,
            main_rate_selections,
            output_path=output_path or build_preon_output_rate_card_path("Haulage", shipper),
            source_tab="PreOn_Carriage_Car. Haulage",
            include_thc_origin=False,
            include_positioning=False,
            include_terms_accessorial=False,
            include_add_services_accessorial=False,
            services_imo_cost_type="IMO charge - applicable globally",
        )
        return rate_card_path, rate_card_df

    source_df = load_digi_fcl_rates_dataframe(processing_path=processing_path)
    if _main_rates_optional_supplements(shipper, flow):
        thc_lookup = _load_thc_lookup_if_available(processing_path)
    else:
        thc_lookup = load_fcl_thc_lookup(processing_path=processing_path)
    if shipper == "Siemens Healthineers LATAM" and flow == "FCL":
        try:
            surcharge_lookup = load_rates_surcharge_lookup(processing_path=processing_path, strict=False)
        except FileNotFoundError:
            surcharge_lookup = RatesSurchargeLookup({}, {}, {}, {})
    else:
        surcharge_lookup = load_rates_surcharge_lookup(processing_path=processing_path)

    rate_card_path, rate_card_df = save_rate_card(
        source_df,
        shipper,
        flow,
        thc_lookup=thc_lookup,
        surcharge_lookup=surcharge_lookup,
        glossary_lookups=glossary_lookups,
        individual_selections=individual_selections,
        output_path=output_path,
    )
    return rate_card_path, rate_card_df


def _carrier_summaries(
    flow: str,
    shipper: str,
    individual_selections: list[SubfolderSelection],
    glossary_lookups: list[GlossaryFeeLookup],
) -> list[CarrierSummary]:
    lookup_by_key = {lookup.carrier_key.upper(): lookup for lookup in glossary_lookups}
    summaries: list[CarrierSummary] = []

    for selection in individual_selections:
        carrier_key = detect_carrier_key(
            selection.file_path.name,
            shipper=shipper,
            flow=flow,
        )
        if not carrier_key:
            continue
        carrier_code = (
            carrier_code_from_filename(
                selection.file_path.name,
                shipper=shipper,
                flow=flow,
            )
            or carrier_key
        )
        lookup = lookup_by_key.get(carrier_key)
        summary = CarrierSummary(
            carrier_key=carrier_key,
            carrier_code=carrier_code,
            individual_rate_file=selection.file_path.name,
        )
        if lookup is not None:
            summary.tmp_fee_blocks = len(lookup.fees.tmp_fees)
            if lookup.fees.financing_fee is not None:
                financing = lookup.fees.financing_fee
                summary.financing_fee = (
                    f"{financing.currency} {financing.rate_20} / {financing.rate_40}"
                )
        summaries.append(summary)
    return summaries


def run_pipeline(
    flow: str,
    shipper: str,
    selections: list[SubfolderSelection],
    underflow: str | None = None,
) -> PipelineResult:
    if flow not in IMPLEMENTED_FLOWS:
        raise ValueError(f"Flow '{flow}' is not implemented yet.")

    errors = validate_rate_card_selections(flow, selections, shipper, underflow=underflow)
    if errors:
        raise ValueError(f"Cannot build {flow} rate card:\n" + "\n".join(f"  - {e}" for e in errors))

    _, individual_selections = split_selections(selections)
    context = run_extraction(flow, shipper, selections, underflow=underflow)
    source_file = primary_source_file_for_output(flow, selections, underflow=underflow)
    result_output_path = build_flow_result_output_path(flow, source_file, underflow=underflow)

    if flow == "LCL":
        carrier_slug = resolve_lcl_carrier_slug(shipper, individual_selections)
        rate_card_path, rate_card_df = run_rate_card_build(
            shipper,
            context.output_path,
            flow,
            individual_selections=individual_selections,
            output_path=result_output_path,
        )
        glossary_lookups = load_glossary_fee_lookups(
            shipper,
            individual_selections,
            flow=flow,
        )
        return PipelineResult(
            shipper=shipper,
            flow=flow,
            processing_path=context.output_path,
            rate_card_path=rate_card_path,
            source_rows=len(rate_card_df),
            rate_card_rows=len(rate_card_df),
            rate_card_columns=len(rate_card_df.columns),
            carriers=_carrier_summaries(flow, shipper, individual_selections, glossary_lookups),
        )
    if flow == "Pre/on carriage":
        preon_builder_selections = (
            selections if underflow == "generic" else individual_selections
        )
        rate_card_path, rate_card_df = run_rate_card_build(
            shipper,
            context.output_path,
            flow,
            underflow=underflow,
            individual_selections=preon_builder_selections,
            output_path=result_output_path,
        )
        return PipelineResult(
            shipper=shipper,
            flow=flow,
            processing_path=context.output_path,
            rate_card_path=rate_card_path,
            source_rows=0,
            rate_card_rows=len(rate_card_df),
            rate_card_columns=len(rate_card_df.columns),
            carriers=[],
        )
    if flow == "Haulage":
        rate_card_path, rate_card_df = run_rate_card_build(
            shipper,
            context.output_path,
            flow,
            individual_selections=selections,
            output_path=result_output_path,
        )
        return PipelineResult(
            shipper=shipper,
            flow=flow,
            processing_path=context.output_path,
            rate_card_path=rate_card_path,
            source_rows=0,
            rate_card_rows=len(rate_card_df),
            rate_card_columns=len(rate_card_df.columns),
            carriers=[],
        )

    glossary_lookups = load_glossary_fee_lookups(
        shipper,
        individual_selections,
        flow=flow,
    )
    rate_card_path, rate_card_df = run_rate_card_build(
        shipper,
        context.output_path,
        flow,
        glossary_lookups=glossary_lookups,
        individual_selections=individual_selections,
        output_path=result_output_path,
    )

    source_df = load_digi_fcl_rates_dataframe(processing_path=context.output_path)
    return PipelineResult(
        shipper=shipper,
        flow=flow,
        processing_path=context.output_path,
        rate_card_path=rate_card_path,
        source_rows=len(source_df),
        rate_card_rows=len(rate_card_df),
        rate_card_columns=len(rate_card_df.columns),
        carriers=_carrier_summaries(flow, shipper, individual_selections, glossary_lookups),
    )


def run_fcl_pipeline(
    shipper: str,
    selections: list[SubfolderSelection],
) -> PipelineResult:
    return run_pipeline("FCL", shipper, selections)


def run_bcn_pipeline(
    shipper: str,
    selections: list[SubfolderSelection],
) -> PipelineResult:
    return run_pipeline("BCN", shipper, selections)


def run_lcl_pipeline(
    shipper: str,
    selections: list[SubfolderSelection],
) -> PipelineResult:
    return run_pipeline("LCL", shipper, selections)


def run_preon_pipeline(
    shipper: str,
    selections: list[SubfolderSelection],
    underflow: str = "per carrier",
) -> PipelineResult:
    return run_pipeline("Pre/on carriage", shipper, selections, underflow=underflow)


def print_pipeline_summary(result: PipelineResult) -> None:
    print("\nPipeline complete")
    print("=" * 18)
    print(f"  Shipper: {result.shipper}")
    print(f"  Flow: {result.flow}")
    print(f"  Source rows: {result.source_rows}")
    print(f"  Rate card rows: {result.rate_card_rows}")
    print(f"  Rate card columns: {result.rate_card_columns}")
    print(f"  Individual rate cards processed: {len(result.carriers)}")
    for carrier in result.carriers:
        print(f"\n  {carrier.carrier_code} ({carrier.individual_rate_file})")
        print(f"    TMP fee blocks: {carrier.tmp_fee_blocks}")
        if carrier.financing_fee:
            print(f"    Financing fee: {carrier.financing_fee}")
    print(f"\nExtracted data:\n  {result.processing_path}")
    print(f"\nRate card:\n  {result.rate_card_path}")
