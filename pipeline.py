"""End-to-end rate card pipeline: extract input files and build the rate card."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from carrier_lookup import carrier_code_from_filename, detect_carrier_key
from config import GLOSSARY_TAB, IMPLEMENTED_FLOWS, INDIVIDUAL_RATE_SUBFOLDER, RATE_CARD_REQUIRED_TABS_BY_FLOW
from extractor import ProcessingContext, SubfolderSelection, save_selections_to_excel
from glossary_lookup import GlossaryFeeLookup, load_glossary_fee_lookups
from lcl_rate_card_builder import (
    build_output_rate_card_path as build_lcl_output_rate_card_path,
    is_lcl_rates_tab,
    resolve_lcl_carrier_slug,
    save_lcl_rate_card,
)
from preon_carriage_builder import save_preon_generic_rate_card, save_preon_per_carrier_rate_card
from preon_carriage_builder import (
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

HEALTHINEERS_SHIPPERS = frozenset({"Siemens Healthineers", "Siemens Healthineers LATAM"})
HEALTHINEERS_MAIN_RATES_FLOWS = frozenset({"FCL", "BCN"})


def _healthineers_main_rates_optional_supplements(shipper: str, flow: str) -> bool:
    return shipper in HEALTHINEERS_SHIPPERS and flow in HEALTHINEERS_MAIN_RATES_FLOWS


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
    """Return validation errors for rate card build."""
    if flow not in RATE_CARD_REQUIRED_TABS_BY_FLOW:
        return [f"Flow '{flow}' is not implemented yet."]

    if flow == "LCL":
        return _validate_lcl_selections(selections)
    if flow == "Pre/on carriage":
        return _validate_preon_selections(selections, underflow=underflow)
    if flow == "Haulage":
        return _validate_haulage_selections(selections, shipper=shipper)

    shared, individual = split_selections(selections)
    selected = _selected_tabs(shared)
    errors: list[str] = []
    required_tabs = RATE_CARD_REQUIRED_TABS_BY_FLOW[flow]
    if _healthineers_main_rates_optional_supplements(shipper, flow):
        required_tabs = tuple(tab for tab in required_tabs if tab != "FCL_THC")
    for tab in required_tabs:
        if tab not in selected:
            errors.append(f"Missing required tab: {tab}")
    if not individual and not _healthineers_main_rates_optional_supplements(shipper, flow):
        errors.append("At least one individual rate file must be selected.")
    for selection in individual:
        carrier_key = detect_carrier_key(
            selection.file_path.name,
            shipper=shipper,
            flow=flow,
        )
        if not carrier_key:
            errors.append(
                f"Could not detect carrier from individual rate file: {selection.file_path.name}"
            )
    return errors


def _validate_preon_selections(
    selections: list[SubfolderSelection],
    underflow: str | None = None,
) -> list[str]:
    errors: list[str] = []
    if underflow == "generic":
        main_rates = [s for s in selections if s.subfolder == "main rates"]
        if not main_rates:
            errors.append("Main rates file must be selected for Pre/on carriage generic.")
            return errors

        required_tabs = {
            "PreOnCarriage_Containerized_EU",
            "DIGI_FCL_Rates",
            "HAPAG_Terms & Conditions",
            "MAERSK_Terms & Condition",
            "MSC_Terms & Conditions",
            "ONE_Terms & Conditions",
        }
        add_services_aliases = {
            "Add_Services_Glomb_Br. Hafenb",
            "Add_Services_Glomb_Br. Hafenb.",
        }
        for selection in main_rates:
            missing = required_tabs.difference(selection.tabs)
            if not any(tab in selection.tabs for tab in add_services_aliases):
                missing.add("Add_Services_Glomb_Br. Hafenb")
            if missing:
                errors.append(
                    f"{selection.file_path.name} is missing required tabs: {', '.join(sorted(missing))}"
                )
        return errors

    individual = [s for s in selections if s.subfolder == INDIVIDUAL_RATE_SUBFOLDER]
    if not individual:
        errors.append("At least one individual rate file must be selected for Pre/on carriage.")
        return errors

    required_tabs = {"Pre-On-Carriage_RoW", "Glossary", "Emergency Dieselfloater Pre_On"}
    for selection in individual:
        missing = required_tabs.difference(selection.tabs)
        if missing:
            errors.append(
                f"{selection.file_path.name} is missing required tabs: {', '.join(sorted(missing))}"
            )
    return errors


def _validate_haulage_selections(
    selections: list[SubfolderSelection],
    *,
    shipper: str,
) -> list[str]:
    errors: list[str] = []
    if shipper != "Siemens Healthineers":
        errors.append("Haulage flow is implemented only for Siemens Healthineers.")
        return errors
    main_rates = [s for s in selections if s.subfolder == "main rates"]
    if not main_rates:
        errors.append("Main rates file must be selected for Haulage.")
        return errors
    required_tabs = {
        "PreOn_Carriage_Car. Haulage",
        "PreOn_Containerized_EU_Services",
    }
    for selection in main_rates:
        missing = required_tabs.difference(selection.tabs)
        if missing:
            errors.append(
                f"{selection.file_path.name} is missing required tabs: {', '.join(sorted(missing))}"
            )
    return errors


def _validate_lcl_selections(selections: list[SubfolderSelection]) -> list[str]:
    errors: list[str] = []
    individual = [selection for selection in selections if selection.subfolder == INDIVIDUAL_RATE_SUBFOLDER]
    if not individual:
        errors.append("At least one individual rate file must be selected for LCL.")
        return errors

    has_lcl_rates_tab = False
    for selection in individual:
        lcl_tabs = [tab for tab in selection.tabs if is_lcl_rates_tab(tab)]
        if not lcl_tabs:
            errors.append(
                f"No LCL_Rates tab selected for individual rate file: {selection.file_path.name}"
            )
        else:
            has_lcl_rates_tab = True

    if not has_lcl_rates_tab:
        errors.append("At least one tab containing LCL_Rates must be selected.")
    return errors


def validate_fcl_selections(
    selections: list[SubfolderSelection],
    shipper: str,
) -> list[str]:
    return validate_rate_card_selections("FCL", selections, shipper)


def warn_rate_card_selections(
    selections: list[SubfolderSelection],
    flow: str = "FCL",
) -> list[str]:
    if flow == "LCL":
        return []

    warnings: list[str] = []
    _, individual = split_selections(selections)
    for selection in individual:
        if GLOSSARY_TAB not in selection.tabs:
            warnings.append(
                f"Tab '{GLOSSARY_TAB}' not selected for {selection.file_path.name} — "
                "TMP Fee and Financing Fee columns will be omitted."
            )
    return warnings


def warn_fcl_selections(selections: list[SubfolderSelection]) -> list[str]:
    return warn_rate_card_selections(selections)


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
    if _healthineers_main_rates_optional_supplements(shipper, flow):
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

    if flow == "LCL":
        carrier_slug = resolve_lcl_carrier_slug(shipper, individual_selections)
        rate_card_path, rate_card_df = run_rate_card_build(
            shipper,
            context.output_path,
            flow,
            individual_selections=individual_selections,
            output_path=build_lcl_output_rate_card_path(flow, shipper, carrier_slug),
        )
    if flow == "Pre/on carriage":
        carrier_slug = resolve_preon_carrier_slug(shipper, individual_selections)
        preon_builder_selections = (
            selections if underflow == "generic" else individual_selections
        )
        rate_card_path, rate_card_df = run_rate_card_build(
            shipper,
            context.output_path,
            flow,
            underflow=underflow,
            individual_selections=preon_builder_selections,
            output_path=build_preon_output_rate_card_path(
                "Pre_on_carriage_generic" if underflow == "generic" else "Pre_on_carriage",
                shipper,
                None if underflow == "generic" else carrier_slug,
            ),
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
            output_path=build_preon_output_rate_card_path("Haulage", shipper),
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
        output_path=build_output_rate_card_path(flow, shipper),
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
