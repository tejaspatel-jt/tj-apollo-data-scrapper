#!/usr/bin/env python3
"""
═══════════════════════════════════════════════════════════════
  Campaign Formatter — FINAL STEP (Post-Enrichment)

  PURPOSE:
    Convert Enriched / Master DB → Campaign Ready Excel

  INPUT:
    - Enriched file (multi-sheet OR single sheet)
    OR
    - Master DB (recommended)

  OUTPUT:
    - ignite_ready_contacts_DDMMMYYYY_HHMMSS.xlsx

  FEATURES:
    ✓ Column Mapping (Apollo → Campaign format)
    ✓ Region Mapping (Country → Region)
    ✓ Unknown Country Detection
    ✓ Clean Campaign Structure (LinkedIn + Email workflow)
═══════════════════════════════════════════════════════════════
"""

import pandas as pd
import os
from datetime import datetime


# ─── PATH SETUP ───────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ╔══════════════════════════════════════════════════════════════╗
# ║                     CONFIGURATION                           ║
# ╚══════════════════════════════════════════════════════════════╝

TIMESTAMP = datetime.now().strftime("%d%b%Y_%H%M%S").lower()

# 🔁 Choose your source
USE_MASTER_DB = True   # True → use master DB | False → use enriched file

# MASTER_DB_FILE = "master_lead_database.xlsx"
MASTER_DB_FILE = "master_lead_database_sample.xlsx"    # Delete me later

ENRICHED_FILE = "apollo_enriched_latest.xlsx"  # change if needed
ENRICHED_SHEET_NAME = "Enriched"

OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    f"ignite_ready_contacts_{TIMESTAMP}.xlsx"
)


# ╔══════════════════════════════════════════════════════════════╗
# ║                COUNTRY → REGION MAPPING                     ║
# ╚══════════════════════════════════════════════════════════════╝

COUNTRY_REGION_MAP = {
    "Africa": [
            "Kenya", "Mauritius", "Nigeria", "South Africa", "Sudan"
        ],
        "Asia-Pacific": [
            "Australia", "Bangladesh", "China", "Hong Kong", "India",
            "Japan", "Kazakhstan", "Maldives", "New Zealand", "Pakistan",
            "South Korea", "Sri Lanka", "Papua New Guinea"
        ],
        "Europe": [
            "Austria", "Belarus", "Belgium", "Bosnia and Herzegovina", "Bulgaria",
            "Croatia", "Cyprus", "Czech Republic", "Czechia", "Denmark",
            "Estonia", "Finland", "France", "Georgia", "Germany", "Greece",
            "Hungary", "Iceland", "Ireland", "Italy", "Latvia", "Liechtenstein",
            "Lithuania", "Luxembourg", "Macedonia (FYROM)", "Malta", "Moldova",
            "Monaco", "Netherlands", "Norway", "Poland", "Portugal", "Romania",
            "Russia", "Serbia", "Slovakia", "Slovenia", "Spain", "Sweden",
            "Switzerland", "Ukraine", "United Kingdom","Armenia","Jersey","Kosovo",
            "Montenegro"
        ],
        "Latin America": [
            "Brazil", "British Virgin Islands","Ecuador","Argentina","Colombia","Uruguay"
        ],
        "Middle East": [
            "Bahrain", "Egypt", "Iraq", "Israel", "Jordan", "Kuwait",
            "Oman", "Qatar", "Saudi Arabia", "Tuerkiye", "Turkey",
            "United Arab Emirates"
        ],
        "North America": [
            "Canada", "United States", "Mexico","Puerto Rico"
        ],
        "Southeast Asia": [
            "Indonesia", "Malaysia", "Myanmar (Burma)", "Philippines",
            "Singapore", "Thailand", "Vietnam"
        ]
}


# ╔══════════════════════════════════════════════════════════════╗
# ║                     HELPER FUNCTIONS                        ║
# ╚══════════════════════════════════════════════════════════════╝

def load_input_data():
    """Load data from Master DB or Enriched file."""

    if USE_MASTER_DB:
        path = os.path.join(OUTPUT_DIR, MASTER_DB_FILE)
        print(f"  Source: Master DB → {MASTER_DB_FILE}")
        df = pd.read_excel(path, dtype=str)

    else:
        path = os.path.join(OUTPUT_DIR, ENRICHED_FILE)
        print(f"  Source: Enriched File → {ENRICHED_FILE}")
        df = pd.read_excel(path, sheet_name=ENRICHED_SHEET_NAME, dtype=str)

    df.fillna("", inplace=True)
    print(f"  Loaded {len(df)} rows\n")

    return df


def map_country_to_region(df):
    """Add Region column based on Country mapping."""

    possible_cols = ["Company Country", "Country", "Country Name", "country"]
    country_col = None

    for col in possible_cols:
        if col in df.columns:
            country_col = col
            break

    if not country_col:
        raise ValueError(f"No country column found in {possible_cols}")

    def get_region(country):
        if not country or str(country).strip() == "":
            return "UnknownRegion"

        country_clean = str(country).strip().lower()

        for region, countries in COUNTRY_REGION_MAP.items():
            if any(c.lower() == country_clean for c in countries):
                return region

        return "UnknownRegion"

    df["Region"] = df[country_col].apply(get_region)

    unknowns = df[df["Region"] == "UnknownRegion"][country_col].unique().tolist()

    return df, sorted([u for u in unknowns if u])


def build_campaign_format(df):
    """Convert dataset → Campaign-ready structure."""

    column_map = {
        "Industry": "Industry",
        "Country": "Country",
        "Company Name for Emails": "Company Name",
        "Company Linkedin Url": "Company LinkedIn",
        "Website": "Website",
        "First Name": "First Name",
        "Last Name": "Last Name",
        "Title": "Designation",
        "Person Linkedin Url": "Person LinkedIn",
        "Email": "Email",
    }

    final_df = pd.DataFrame()

    for src, dest in column_map.items():
        final_df[dest] = df[src] if src in df.columns else ""

    # Add workflow columns
    extra_cols = [
        "No", "Region", "Like", "Comment", "Connection Sent",
        "M1", "M2", "M3", "Email Sent", "Status", "Notes"
    ]

    for col in extra_cols:
        final_df[col] = ""

    # Order columns
    ordered_cols = [
        "No", "Industry", "Region", "Country",
        "Company Name", "Company LinkedIn", "Website",
        "First Name", "Last Name", "Designation",
        "Person LinkedIn",
        "Like", "Comment", "Connection Sent",
        "M1", "M2", "M3",
        "Email", "Email Sent", "Status", "Notes"
    ]

    final_df = final_df[ordered_cols]

    # Serial number
    final_df["No"] = range(1, len(final_df) + 1)

    return final_df


# ╔══════════════════════════════════════════════════════════════╗
# ║                        MAIN RUNNER                         ║
# ╚══════════════════════════════════════════════════════════════╝

def main():
    print("\n" + "═" * 62)
    print("  Campaign Formatter — Final Step")
    print("═" * 62 + "\n")

    df = load_input_data()

    # ── Step 1: Region Mapping ─────────────────────────────
    df, unknowns = map_country_to_region(df)

    # ── Step 2: Campaign Format ────────────────────────────
    campaign_df = build_campaign_format(df)

    # ── Step 3: Save Output ────────────────────────────────
    campaign_df.to_excel(OUTPUT_FILE, index=False)

    print(f"  ✅ Campaign File Created → output/{os.path.basename(OUTPUT_FILE)}")
    print(f"     Total rows : {len(campaign_df)}")

    # ── Step 4: Unknown Countries Report ───────────────────
    if unknowns:
        print(f"\n  ⚠ Unmapped Countries ({len(unknowns)}):")
        for c in unknowns[:10]:
            print(f"     - {c}")
        if len(unknowns) > 10:
            print("     ...")

    else:
        print("\n  ✅ All countries mapped successfully")

    print("\n" + "═" * 62 + "\n")


if __name__ == "__main__":
    main()