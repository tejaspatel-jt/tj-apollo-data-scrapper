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

def normalize_columns(df):
    """Normalize column names (lowercase + strip)."""
    df.columns = [c.strip().lower() for c in df.columns]
    return df

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


# v2 : country column name fixed
def map_country_to_region(df):
    possible_cols = ["company country", "country", "country name"]
    
    country_col = next((c for c in df.columns if c in possible_cols), None)

    if not country_col:
        raise ValueError(f"No country column found in {possible_cols}")

    def get_region(country):
        if not country:
            return "UnknownRegion"

        country = str(country).strip().lower()

        for region, countries in COUNTRY_REGION_MAP.items():
            if any(c.lower() == country for c in countries):
                return region

        return "UnknownRegion"

    df["Region"] = df[country_col].apply(get_region)

    unknowns = df[df["Region"] == "UnknownRegion"][country_col].unique().tolist()
    return df, sorted([u for u in unknowns if u])

def map_master_to_campaign(df):
    """
    🔥 CORE FIX: Map Master DB schema → Campaign schema
    """

    mapping = {
        "organization_name": "company name",
        "org_linkedin_url": "company linkedin",
        "website_url": "website",
        "first_name": "first name",
        "last_name": "last name",
        "title": "designation",
        "linkedin_url": "person linkedin",
        "email": "email",
        "org_industry": "industry",
        "country": "country",
    }

    for src, dest in mapping.items():
        if src in df.columns:
            df[dest] = df[src]
        else:
            df[dest] = ""

    return df


# v2 : required_cols added to ensure all necessary columns are present in the final output, even if missing in the input
def build_campaign_format(df):
    final_df = pd.DataFrame()

    # 🔥 Exact mapping (final output format)
    COLUMN_MAPPING = {
    "org_industry": "Industry",
    "country": "Country",
    "Region": "Region",  # already created
    "organization_name": "Company Name",
    "org_linkedin_url": "Company LinkedIn",
    "website_url": "Website",
    "first_name": "First Name",
    "last_name": "Last Name",
    "title": "Designation",
    "linkedin_url": "Person LinkedIn",
    "email": "Email",
}

    for src, dest in COLUMN_MAPPING.items():
        final_df[dest] = df[src] if src in df.columns else ""

    # Workflow columns (exact names)
    workflow_cols = [
        "No", "Like", "Comment", "Connection Sent",
        "M1", "M2", "M3", "Email Sent", "Status", "Notes"
    ]

    for col in workflow_cols:
        final_df[col] = ""

    # Final order (exact as you want)
    ordered = [
        "No", "Industry", "Region", "Country",
        "Company Name", "Company LinkedIn", "Website",
        "First Name", "Last Name", "Designation",
        "Person LinkedIn",
        "Like", "Comment", "Connection Sent",
        "M1", "M2", "M3",
        "Email", "Email Sent", "Status", "Notes"
    ]

    final_df = final_df[ordered]

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

    # 🔥 STEP 1: Schema Mapping (CRITICAL FIX)
    df = map_master_to_campaign(df)

    # 🔥 STEP 2: Region Mapping
    df, unknowns = map_country_to_region(df)

    # 🔥 STEP 3: Build Final Format
    campaign_df = build_campaign_format(df)

    # 🔥 STEP 4: Save
    campaign_df.to_excel(OUTPUT_FILE, index=False)

    print(f"  ✅ Campaign File Created → output/{os.path.basename(OUTPUT_FILE)}")
    print(f"     Total rows : {len(campaign_df)}")

    if unknowns:
        print(f"\n  ⚠ Unmapped Countries ({len(unknowns)}): {', '.join(unknowns[:5])}")
    else:
        print("\n  ✅ All countries mapped successfully")

    print("\n" + "═" * 62 + "\n")


if __name__ == "__main__":
    main()