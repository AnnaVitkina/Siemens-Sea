"""Interactive end-to-end pipeline: file selection, extraction, and rate card build."""

from __future__ import annotations

import sys

from config import IMPLEMENTED_FLOWS
from extractor import save_selections_to_excel
from innomotics_builder import save_innomotics_rate_card
from main import collect_selections, prompt_flow, prompt_preon_underflow, prompt_shipper
from pipeline import (
    print_pipeline_summary,
    run_pipeline,
    validate_rate_card_selections,
    warn_rate_card_selections,
)


def main() -> int:
    print("Siemens Sea Rate Card Pipeline")
    print("=" * 31)
    print("This will extract selected input files and build the final rate card.\n")

    shipper = prompt_shipper()
    if shipper == "Innomotics":
        flow = "Innomotics"
        underflow = None
    elif shipper == "Siemens Healthineers LATAM":
        flow = "FCL"
        underflow = None
    else:
        flow = prompt_flow()
        underflow = prompt_preon_underflow() if flow == "Pre/on carriage" else None
    selections = collect_selections(flow, shipper=shipper, underflow=underflow)

    if not selections:
        print("\nNothing to process. Exiting.")
        return 1

    print("\nProcessing summary:")
    print(f"  Shipper: {shipper}")
    print(f"  Flow: {flow}")
    if underflow:
        print(f"  Variant: {underflow}")
    for selection in selections:
        print(f"  {selection.subfolder}: {selection.file_path.name}")
        for tab in selection.tabs:
            print(f"    - {tab}")

    if shipper == "Innomotics":
        confirm = input("\nProceed with extraction and rate card build? [Y/n]: ").strip().lower()
        if confirm in {"n", "no"}:
            print("Cancelled.")
            return 0
        processing_path, _context = save_selections_to_excel(flow, shipper, selections, underflow=underflow)
        rate_card_path, rate_card_df = save_innomotics_rate_card(shipper, selections)
        print("\nPipeline complete")
        print("==================")
        print(f"  Shipper: {shipper}")
        print(f"  Flow: {flow}")
        print("  Source rows: 0")
        print(f"  Rate card rows: {len(rate_card_df)}")
        print(f"  Rate card columns: {len(rate_card_df.columns)}")
        print("\nExtracted data:")
        print(f"  {processing_path}")
        print("\nRate card:")
        print(f"  {rate_card_path}")
        return 0

    if flow in IMPLEMENTED_FLOWS:
        errors = validate_rate_card_selections(flow, selections, shipper, underflow=underflow)
        for warning in warn_rate_card_selections(selections, flow=flow):
            print(f"\nWarning: {warning}")
        if errors:
            print("\nCannot continue — missing required tabs:")
            for error in errors:
                print(f"  - {error}")
            return 1
    else:
        print(f"\nFlow '{flow}' is not implemented yet.")
        return 1

    confirm = input("\nProceed with extraction and rate card build? [Y/n]: ").strip().lower()
    if confirm in {"n", "no"}:
        print("Cancelled.")
        return 0

    try:
        result = run_pipeline(flow, shipper, selections, underflow=underflow)
    except (FileNotFoundError, ValueError, KeyError, PermissionError) as error:
        print(f"\nPipeline failed: {error}")
        return 1

    print_pipeline_summary(result)
    return 0


if __name__ == "__main__":
    sys.exit(main())
