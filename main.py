"""Interactive rate card extraction entry point."""

from __future__ import annotations

import sys
import re

from config import (
    FLOW_TAB_PRESETS,
    FLOWS,
    IMPLEMENTED_FLOWS,
    INNOMOTICS_SUBFOLDER,
    INDIVIDUAL_RATE_SUBFOLDER,
    INPUT_SUBFOLDERS,
    PREON_UNDERFLOWS,
    SHIPPERS,
)
from extractor import (
    SubfolderSelection,
    list_excel_files,
    list_input_subfolders,
    read_workbook_sheet_names,
    resolve_default_tabs,
    save_selections_to_excel,
)


def _innomotics_default_tabs(available_tabs: list[str]) -> list[str]:
    suggested: list[str] = []
    for tab in available_tabs:
        normalized = re.sub(r"[^a-z0-9]+", " ", tab.lower()).strip()
        if "lcl" not in normalized:
            continue
        if "rate" in normalized or "card" in normalized:
            suggested.append(tab)
    return suggested


def prompt_multiple_choices(
    prompt: str,
    options: list[str],
    allow_skip: bool = False,
) -> list[str]:
    if not options:
        print("  No options available.")
        return []

    print(prompt)
    for index, option in enumerate(options, start=1):
        print(f"  [{index}] {option}")

    if allow_skip:
        print("  [0] Skip this subfolder")

    print("Enter comma-separated numbers to select multiple files.")
    while True:
        raw = input("Enter number(s): ").strip()
        if allow_skip and raw == "0":
            return []
        if not raw:
            print("Invalid choice. Try again.")
            continue

        selected: list[str] = []
        invalid_parts: list[str] = []
        for part in raw.split(","):
            part = part.strip()
            if not part:
                continue
            if part.isdigit():
                choice = int(part)
                if 1 <= choice <= len(options):
                    option = options[choice - 1]
                    if option not in selected:
                        selected.append(option)
                else:
                    invalid_parts.append(part)
            else:
                invalid_parts.append(part)

        if invalid_parts:
            print(f"Invalid choices ignored: {', '.join(invalid_parts)}")

        if selected:
            return selected
        print("No valid files selected. Try again.")


def prompt_choice(prompt: str, options: list[str], allow_skip: bool = False) -> str | None:
    if not options:
        print("  No options available.")
        return None

    print(prompt)
    for index, option in enumerate(options, start=1):
        print(f"  [{index}] {option}")

    if allow_skip:
        print("  [0] Skip this subfolder")

    while True:
        raw = input("Enter number: ").strip()
        if allow_skip and raw == "0":
            return None
        if raw.isdigit():
            choice = int(raw)
            if 1 <= choice <= len(options):
                return options[choice - 1]
        print("Invalid choice. Try again.")


def prompt_shipper() -> str:
    print("\nWhich Shipper is it?")
    for index, shipper in enumerate(SHIPPERS, start=1):
        print(f"  [{index}] {shipper}")

    while True:
        raw = input("Enter number: ").strip()
        if raw.isdigit():
            choice = int(raw)
            if 1 <= choice <= len(SHIPPERS):
                return SHIPPERS[choice - 1]
        print("Invalid choice. Try again.")


def prompt_flow() -> str:
    flow_options = list(FLOWS.keys())
    print("\nSelect processing flow:")
    for index, flow in enumerate(flow_options, start=1):
        status = "available" if flow in IMPLEMENTED_FLOWS else "coming soon"
        print(f"  [{index}] {flow} ({status})")

    while True:
        raw = input("Enter number [default: 1 FCL]: ").strip()
        if not raw:
            return "FCL"
        if raw.isdigit():
            choice = int(raw)
            if 1 <= choice <= len(flow_options):
                selected = flow_options[choice - 1]
                if selected not in IMPLEMENTED_FLOWS:
                    print(f"Flow '{selected}' is not implemented yet. Try again.")
                    continue
                return selected
        print("Invalid choice. Try again.")


def prompt_preon_underflow() -> str:
    print("\nSelect Pre/on carriage variant:")
    for index, underflow in enumerate(PREON_UNDERFLOWS, start=1):
        status = "available"
        print(f"  [{index}] {underflow} ({status})")

    while True:
        raw = input("Enter number [default: 1 per carrier]: ").strip()
        if not raw:
            return "per carrier"
        if raw.isdigit():
            choice = int(raw)
            if 1 <= choice <= len(PREON_UNDERFLOWS):
                return PREON_UNDERFLOWS[choice - 1]
        print("Invalid choice. Try again.")


def prompt_tabs(
    subfolder: str,
    file_name: str,
    default_tabs: list[str],
    available_tabs: list[str],
) -> list[str]:
    default_indices = {
        index
        for index, tab in enumerate(available_tabs, start=1)
        if tab in default_tabs
    }

    print(f"\nTabs for '{subfolder}' -> {file_name}")
    print("Available tabs in file (* = pre-selected):")
    for index, tab in enumerate(available_tabs, start=1):
        marker = "*" if index in default_indices else " "
        print(f"  [{index}] {marker} {tab}")

    if default_indices:
        default_numbers = ", ".join(str(index) for index in sorted(default_indices))
        print(f"\nPre-selected tab numbers: {default_numbers}")
    else:
        print("\nNo pre-selected tabs matched this file.")

    print(
        "Press Enter to confirm pre-selected tabs, "
        "or enter comma-separated tab numbers to override."
    )
    raw = input("Tab numbers: ").strip()
    if not raw:
        return default_tabs

    selected_indices: list[int] = []
    invalid_parts: list[str] = []
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        if part.isdigit():
            choice = int(part)
            if 1 <= choice <= len(available_tabs):
                if choice not in selected_indices:
                    selected_indices.append(choice)
            else:
                invalid_parts.append(part)
        else:
            invalid_parts.append(part)

    if invalid_parts:
        print(f"Invalid tab numbers ignored: {', '.join(invalid_parts)}")

    if not selected_indices:
        return default_tabs

    return [available_tabs[index - 1] for index in selected_indices]


def collect_selections(
    flow: str,
    shipper: str,
    underflow: str | None = None,
) -> list[SubfolderSelection]:
    selections: list[SubfolderSelection] = []

    if shipper == "Innomotics":
        files = list_excel_files(INNOMOTICS_SUBFOLDER)
        print(f"\n--- {INNOMOTICS_SUBFOLDER} ---")
        if not files:
            print("  No Excel files found. Skipping.")
            return selections
        file_labels = [path.name for path in files]
        selected_name = prompt_choice(
            "Select file to process:",
            file_labels,
            allow_skip=True,
        )
        if selected_name is None:
            print(f"Skipped '{INNOMOTICS_SUBFOLDER}'.")
            return selections

        file_path = files[file_labels.index(selected_name)]
        available_tabs = read_workbook_sheet_names(file_path)
        default_tabs = _innomotics_default_tabs(available_tabs)
        tabs = prompt_tabs(
            subfolder=INNOMOTICS_SUBFOLDER,
            file_name=selected_name,
            default_tabs=default_tabs,
            available_tabs=available_tabs,
        )
        if not tabs:
            print(f"No tabs selected for '{selected_name}'.")
            return selections
        selections.append(
            SubfolderSelection(
                subfolder=INNOMOTICS_SUBFOLDER,
                file_path=file_path,
                tabs=tabs,
            )
        )
        return selections

    if shipper == "Siemens Healthineers LATAM":
        files = list_excel_files("main rates")
        print("\n--- main rates ---")
        if not files:
            print("  No Excel files found. Skipping.")
            return selections
        file_labels = [path.name for path in files]
        selected_name = prompt_choice(
            "Select file to process:",
            file_labels,
            allow_skip=True,
        )
        if selected_name is None:
            print("Skipped 'main rates'.")
            return selections
        file_path = files[file_labels.index(selected_name)]
        available_tabs = read_workbook_sheet_names(file_path)
        default_tabs = resolve_default_tabs(flow, "main rates", file_path)
        tabs = prompt_tabs(
            subfolder="main rates",
            file_name=selected_name,
            default_tabs=default_tabs,
            available_tabs=available_tabs,
        )
        if not tabs:
            print(f"No tabs selected for '{selected_name}'.")
            return selections
        selections.append(
            SubfolderSelection(
                subfolder="main rates",
                file_path=file_path,
                tabs=tabs,
            )
        )
        return selections

    subfolders = list_input_subfolders()

    if not subfolders:
        print("No input subfolders found.")
        return selections

    missing = [name for name in INPUT_SUBFOLDERS if name not in subfolders]
    if missing:
        print(f"Warning: missing subfolders: {', '.join(missing)}")

    for subfolder in subfolders:
        if flow == "LCL" and subfolder != INDIVIDUAL_RATE_SUBFOLDER:
            continue
        if flow == "Pre/on carriage":
            if underflow == "per carrier" and subfolder != INDIVIDUAL_RATE_SUBFOLDER:
                continue
            if underflow == "generic" and subfolder != "main rates":
                continue
        if flow == "Haulage" and subfolder != "main rates":
            continue

        files = list_excel_files(subfolder)
        print(f"\n--- {subfolder} ---")
        if not files:
            print("  No Excel files found. Skipping.")
            continue

        file_labels = [path.name for path in files]
        if subfolder == INDIVIDUAL_RATE_SUBFOLDER:
            selected_names = prompt_multiple_choices(
                "Select one or more individual rate files to process:",
                file_labels,
                allow_skip=True,
            )
            if not selected_names:
                print(f"Skipped '{subfolder}'.")
                continue
        else:
            selected_name = prompt_choice(
                "Select file to process:",
                file_labels,
                allow_skip=True,
            )
            if selected_name is None:
                print(f"Skipped '{subfolder}'.")
                continue
            selected_names = [selected_name]

        for selected_name in selected_names:
            file_path = files[file_labels.index(selected_name)]
            available_tabs = read_workbook_sheet_names(file_path)
            default_tabs = resolve_default_tabs(flow, subfolder, file_path)

            if flow not in FLOW_TAB_PRESETS:
                print(f"Flow '{flow}' has no tab presets configured.")
                default_tabs = []

            tabs = prompt_tabs(
                subfolder=subfolder,
                file_name=selected_name,
                default_tabs=default_tabs,
                available_tabs=available_tabs,
            )

            if not tabs:
                print(f"No tabs selected for '{selected_name}'. Skipping.")
                continue

            selections.append(
                SubfolderSelection(
                    subfolder=subfolder,
                    file_path=file_path,
                    tabs=tabs,
                )
            )

    return selections


def main() -> int:
    print("Siemens Sea Rate Card Extractor")
    print("=" * 32)

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

    confirm = input("\nProceed with extraction? [Y/n]: ").strip().lower()
    if confirm in {"n", "no"}:
        print("Cancelled.")
        return 0

    output_path, context = save_selections_to_excel(flow, shipper, selections, underflow=underflow)
    print(f"\nSaved extracted data to:\n  {output_path}")
    print(f"Saved processing context to:\n  {context.output_path.parent / 'latest_processing_context.json'}")
    print("\nTo build the rate card, run: python run_pipeline.py")
    print("Or extraction-only next time: python main.py | rate card only: python build_rate_card.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
