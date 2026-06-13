"""Build matrix-format FCL rate card from extracted processing data."""

from __future__ import annotations

import sys
from pathlib import Path

from glossary_lookup import individual_rate_selections_from_context, load_glossary_fee_lookups
from pipeline import run_rate_card_build
from rate_card_builder import resolve_shipper


def main() -> int:
    print("Siemens Sea FCL Rate Card Builder")
    print("=" * 34)

    try:
        from extractor import load_processing_context

        shipper = resolve_shipper()
        context = load_processing_context()
        if not context or not context.get("output_path"):
            raise FileNotFoundError(
                "No extracted processing data found. Run main.py or run_pipeline.py first."
            )
        flow = context.get("flow", "FCL")
        underflow = context.get("underflow")
        processing_path = Path(context["output_path"])
        individual_selections = individual_rate_selections_from_context(context)
        glossary_lookups = load_glossary_fee_lookups(
            shipper,
            individual_selections,
            flow=flow,
        )
        if flow == "LCL":
            output_path, rate_card_df = run_rate_card_build(
                shipper,
                processing_path,
                flow,
                underflow=underflow,
                individual_selections=individual_selections,
            )
        else:
            output_path, rate_card_df = run_rate_card_build(
                shipper,
                processing_path,
                flow,
                underflow=underflow,
                glossary_lookups=glossary_lookups,
                individual_selections=individual_selections,
            )
    except FileNotFoundError as error:
        print(error)
        return 1

    print(f"Shipper: {shipper}")
    print(f"Flow: {flow}")
    if flow != "LCL":
        print(f"Individual rate cards loaded: {len(glossary_lookups)}")
        for lookup in glossary_lookups:
            print(f"  {lookup.carrier_name}: {len(lookup.fees.tmp_fees)} TMP block(s)")
            if lookup.fees.financing_fee is not None:
                financing = lookup.fees.financing_fee
                print(
                    f"    Financing fee: {financing.currency} "
                    f"{financing.rate_20} / {financing.rate_40}"
                )

    print(f"Rate card rows: {len(rate_card_df)}")
    print(f"Rate card columns: {len(rate_card_df.columns)}")
    print(f"\nSaved rate card to:\n  {output_path}")
    print("\nFor end-to-end processing, use: python run_pipeline.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
