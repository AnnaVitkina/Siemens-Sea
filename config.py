"""Configuration for rate card extraction flows.

Implemented flows:
- FCL — Siemens Divisions (CEV, DHL, KUN) and Siemens Healthineers (FRA, KLN)
- BCN — Siemens Divisions (KUEHNE-DE-FFM-00) and Siemens Healthineers (4 KUEHNE sites)
- LCL — individual-rate LCL_Rates tabs plus Glossary

Other shippers and flows may reuse shared building blocks but can have their
own carrier codes, glossary sections, accessorial rules, and output structure.
"""

import os
from pathlib import Path
from typing import Any

# Default flow/shipper when context does not specify one.
IMPLEMENTED_FLOW = "FCL"
IMPLEMENTED_SHIPPER = "Siemens Divisions"
IMPLEMENTED_FLOWS = ("FCL", "BCN", "LCL", "Pre/on carriage", "Haulage")

CONDITIONAL_RULES_SHEET_NAME = "Conditional Rules"

BCN_EQUIPMENT_TYPE_VALUE = "LTL/Buyer Consolidation"
BCN_MULTIPLIER_LABEL = "Multiplier: Special/BCN ratio"
BCN_POD_COLUMN = "POD+Alternative Destination Terminal"

PROJECT_ROOT = Path(__file__).resolve().parent
COLAB_SHARED_DRIVE_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT Siemens/Siemens Air"
)


def _resolve_data_root() -> Path:
    """Resolve where runtime data folders (input/processing/output) live."""
    override = os.environ.get("SIEMENS_SEA_DATA_ROOT", "").strip()
    if override:
        return Path(override).expanduser()
    if COLAB_SHARED_DRIVE_ROOT.exists():
        return COLAB_SHARED_DRIVE_ROOT
    return PROJECT_ROOT


DATA_ROOT = _resolve_data_root()
INPUT_DIR = DATA_ROOT / "input"
PROCESSING_DIR = DATA_ROOT / "processing"
OUTPUT_DIR = DATA_ROOT / "output"

FCL_BASE_TAB = "DIGI_FCL_Rates"
FCL_THC_BASE_TAB = "FCL_THC"
RATES_BASE_TAB = "Rates"
RATES_REEFER_TAB = "Rates_Reefer_Containers"
GLOSSARY_TAB = "Glossary"
RATE_CARD_SHEET_NAME = "Rate card"
ACCESSORIAL_COSTS_SHEET_NAME = "Accessorial Costs"
ADDON_SMF_TAB = "Add-on SMF (FCL)"

TMP_ACCESSORIAL_COST_NAME = "TMP Fee (TMP-Fee-ALL_IN, FCL; tab 'Glossary')"
LCL_TMP_ACCESSORIAL_COST_NAME = "TMP Fee (TMP-Fee-ALL_IN, LCL; tab 'Glossary')"
FINANCING_ACCESSORIAL_COST_NAME = "Financing Fee"

INPUT_SUBFOLDERS = ("main rates", "THC fee", "individual rate")
INDIVIDUAL_RATE_SUBFOLDER = "individual rate"
INNOMOTICS_SUBFOLDER = "innomotics"

FCL_SHIPPER_CARRIER_CONFIG: dict[str, dict[str, Any]] = {
    "Siemens Divisions": {
        "file_keys": ("CEV", "DHL", "KUN"),
        "carrier_codes": {
            "CEV": "CEVA-DE-FFM-00",
            "DHL": "DHLGLOB-DE-FFM-00",
            "KUN": "KUEHNE-DE-FFM-00",
        },
        "ebs_variants": (
            {
                "key": "kun",
                "cost_label": "Emergency Disruption Surcharge",
                "carrier_code": "KUEHNE-DE-FFM-00",
            },
            {
                "key": "dhl",
                "cost_label": "Emergency Bunker Surcharge",
                "carrier_code": "DHLGLOB-DE-FFM-00",
            },
            {
                "key": "cev",
                "cost_label": "Emergency Imbalance Surcharge",
                "carrier_code": "CEVA-DE-FFM-00",
            },
        ),
        "biofuel_cost_names": {
            "DHL": "Add-on biofuel",
            "KUN": "Biofuel add-on",
        },
    },
    "Siemens Healthineers": {
        "file_keys": ("FRA", "KLN"),
        "carrier_codes": {
            "FRA": "FRACHT-DE-NUER-750",
            "KLN": "KLN FRE-DE-BREM-376",
        },
        "ebs_variants": (
            {
                "key": "fra",
                "cost_label": "Emergency Bunker Surcharge",
                "carrier_code": "FRACHT-DE-NUER-750",
            },
            {
                "key": "kln",
                "cost_label": "Emergency Disruption Surcharge",
                "carrier_code": "KLN FRE-DE-BREM-376",
            },
        ),
        "biofuel_cost_names": {},
    },
}

BCN_SHIPPER_CARRIER_CONFIG: dict[str, dict[str, Any]] = {
    "Siemens Divisions": {
        "file_keys": ("KUN",),
        "carrier_codes": {
            "KUN": "KUEHNE-DE-FFM-00",
        },
        "in_scope_carrier_codes": ("KUEHNE-DE-FFM-00",),
        "ebs_variants": FCL_SHIPPER_CARRIER_CONFIG["Siemens Divisions"]["ebs_variants"],
        "biofuel_cost_names": {
            "KUN": "Biofuel add-on",
        },
    },
    "Siemens Healthineers": {
        "file_keys": ("BREM", "DUES", "FRAN", "NUER"),
        "carrier_codes": {
            "BREM": "KUEHNE-DE-BREM-694",
            "DUES": "KUEHNE-DE-DUES-462",
            "FRAN": "KUEHNE-DE-FRAN-464",
            "NUER": "KUEHNE-DE-NUER-103",
        },
        "in_scope_carrier_codes": (
            "KUEHNE-DE-BREM-694",
            "KUEHNE-DE-DUES-462",
            "KUEHNE-DE-FRAN-464",
            "KUEHNE-DE-NUER-103",
        ),
        "ebs_variants": FCL_SHIPPER_CARRIER_CONFIG["Siemens Healthineers"]["ebs_variants"],
        "biofuel_cost_names": {},
    },
}

SHIPPER_CARRIER_CONFIG_BY_FLOW: dict[str, dict[str, dict[str, Any]]] = {
    "FCL": FCL_SHIPPER_CARRIER_CONFIG,
    "BCN": BCN_SHIPPER_CARRIER_CONFIG,
}

FLOWS = {
    "FCL": "FCL",
    "LCL": "LCL",
    "Pre/on carriage": "Pre/on carriage",
    "Haulage": "Haulage",
    "BCN": "BCN",
}

PREON_UNDERFLOWS = ("per carrier", "generic")

SHIPPERS = (
    "Siemens Divisions",
    "Siemens Healthineers",
    "Innomotics",
    "Siemens Healthineers LATAM",
)

PROCESSING_CONTEXT_FILE = PROCESSING_DIR / "latest_processing_context.json"

RATE_CARD_REQUIRED_TABS_BY_FLOW: dict[str, tuple[str, ...]] = {
    "LCL": (),
    "Pre/on carriage": (),
    "Haulage": (),
    "FCL": (
        FCL_BASE_TAB,
        RATES_BASE_TAB,
        RATES_REEFER_TAB,
        FCL_THC_BASE_TAB,
    ),
    "BCN": (
        FCL_BASE_TAB,
        RATES_BASE_TAB,
        RATES_REEFER_TAB,
        FCL_THC_BASE_TAB,
    ),
}

# Backward-compatible alias.
FCL_RATE_CARD_REQUIRED_TABS = RATE_CARD_REQUIRED_TABS_BY_FLOW["FCL"]

# Default tabs per subfolder for each flow.
FLOW_TAB_PRESETS = {
    "FCL": {
        "main rates": ["DIGI_FCL_Rates", "Rates", "Rates_Reefer_Containers"],
        "THC fee": ["FCL_THC"],
        "individual rate": ["Glossary"],
    },
    "BCN": {
        "main rates": ["DIGI_FCL_Rates", "Rates", "Rates_Reefer_Containers"],
        "THC fee": ["FCL_THC"],
        "individual rate": ["Glossary"],
    },
    "LCL": {
        "individual rate": ["Glossary"],
    },
    "Pre/on carriage": {
        "main rates": [
            "PreOnCarriage_Containerized_EU",
            "PreOn_Containerized_EU_Services",
            "DIGI_FCL_Rates",
            "Add_Services_Glomb_Br. Hafenb",
            "Add_Services_Glomb_Br. Hafenb.",
            "HAPAG_Terms & Conditions",
            "MAERSK_Terms & Condition",
            "MSC_Terms & Conditions",
            "ONE_Terms & Conditions",
            "Agreed Pick-up and Drop-off fees Germany",
        ],
        "individual rate": [
            "Pre-On-Carriage_RoW",
            "Glossary",
            "Emergency Dieselfloater Pre_On",
        ],
    },
    "Haulage": {
        "main rates": [
            "PreOn_Carriage_Car. Haulage",
            "PreOn_Containerized_EU_Services",
        ],
    },
}

# Optional tabs resolved dynamically when present in the workbook.
FLOW_OPTIONAL_TABS = {
    "FCL": {
        "individual rate": ["Add-on SMF (FCL)"],
    },
    "BCN": {
        "individual rate": ["Add-on SMF (FCL)"],
    },
    "LCL": {
        "individual rate": [],
    },
    "Pre/on carriage": {
        "individual rate": [],
    },
}

# Known header row (0-based) per sheet name. None = read sheet without a header row.
SHEET_HEADER_ROWS = {
    "DIGI_FCL_Rates": 0,
    "PreOnCarriage_Containerized_EU": 0,
    "PreOn_Carriage_Car. Haulage": 0,
    "Rates": 2,
    "Rates_Reefer_Containers": 2,
    "FCL_THC": 2,
    "Add-on SMF (FCL)": 3,
    "Pre-On-Carriage_RoW": 3,
    "PreOn_Containerized_EU_Services": 0,
    "Emergency Dieselfloater Pre_On": 5,
    "Glossary": None,
}

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
