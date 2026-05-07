# ❌CAN BE REMOVED ❌ NOT USED IN THIS VERSION — IGNORE
def save__output(enrich_map, input_df):
    """
    Build and write 4 sheets:
      Enriched   — ALL input rows in original order (matched + unmatched)
      Matched    — Rows with email
      Unmatched  — Rows without a match
      CRM Save   — Matched rows with contacts/bulk_create field names
                   (review this, then set CREATE_IN_CRM=True or call manually)
    """
    enriched_rows = []
    unmatched_pids = []

    # ── Step 1: Build full dataset (ALL rows) ─────────────────────
    for _, in_row in input_df.iterrows():
        pid = str(in_row.get("apollo_person_id", "")).strip()
        if pid in enrich_map:
            enriched_rows.append(enrich_map[pid])
        else:
            enriched_rows.append(empty_enriched_row(apollo_person_id=pid))
            unmatched_pids.append(pid)

    all_df       = pd.DataFrame(enriched_rows)
    all_df       = enforce_column_order(all_df)

    # has_email    = all_df["Email"].notna() & (all_df["Email"].astype(str).str.strip() != "")
    # matched_df   = all_df[has_email].reset_index(drop=True)
    # unmatched_df = all_df[~has_email].reset_index(drop=True)

    is_matched = all_df["is_matched"] == "true"
    matched_df = all_df[is_matched].reset_index(drop=True)
    unmatched_df = all_df[~is_matched].reset_index(drop=True)



    # ── CRM Save sheet ────────────────────────────────────────────
    # Matched rows only, columns named exactly as contacts/bulk_create API fields
    crm_rows = [build_crm_payload(row) for row in
                [enrich_map[pid] for pid in [
                    str(r.get("apollo_person_id", "")).strip()
                    for _, r in input_df.iterrows()
                ] if pid in enrich_map]]

    crm_df = pd.DataFrame(crm_rows)
    for col in CRM_SAVE_COLUMNS:
        if col not in crm_df.columns:
            crm_df[col] = ""
    crm_df = crm_df[CRM_SAVE_COLUMNS]

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        all_df.to_excel(writer,       index=False, sheet_name="Enriched")
        matched_df.to_excel(writer,   index=False, sheet_name="Matched")
        unmatched_df.to_excel(writer, index=False, sheet_name="Unmatched")
        crm_df.to_excel(writer,       index=False, sheet_name="CRM Save")

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 45)

    out_name = os.path.basename(OUTPUT_FILE)
    print(f"\n  ✅  Saved → 'output/{out_name}'")
    print(f"       ├─ Enriched (all rows)    : {len(all_df)}  ← matches your input count")
    print(f"       ├─ Matched  (has email)   : {len(matched_df)}")
    print(f"       ├─ Unmatched              : {len(unmatched_df)}")
    print(f"       └─ CRM Save (bulk_create) : {len(crm_df)}")
    if unmatched_pids:
        print(f"\n  Unmatched person IDs ({len(unmatched_pids)}): {', '.join(unmatched_pids[:5])}"
              + ("..." if len(unmatched_pids) > 5 else ""))

    return crm_df






# ==========================

        COLUMN_ALIASES = {
            "company name": ["company name", "name"],
            "domain": ["company domain", "website", "domain"],
            "org industry": ["org industry", "industry"],
            "org employees": ["org employees", "employees"],
        }

        def get_col(df, possible_names):
            for col in possible_names:
                if col in df.columns:
                    return col
            return None

        name_col = get_col(df, COLUMN_ALIASES["company name"])
        domain_col = get_col(df, COLUMN_ALIASES["domain"])
        industry_col = get_col(df, COLUMN_ALIASES["org industry"])
        employees_col = get_col(df, COLUMN_ALIASES["org employees"])

        if not name_col or not domain_col:
            raise ValueError("Missing required columns: company name / domain")

        return [
            {
                "name": str(row.get(name_col, "")).strip(),
                "domain": str(row.get(domain_col, "")).strip(),
                "org_industry_input": str(row.get(industry_col, "")).strip(),
                "org_employees_input": str(row.get(employees_col, "")).strip(),
            }
            for _, row in df.iterrows()
        ]


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
# MASTER_DB_FILE = "master_lead_database_sample.xlsx"    # Delete me later
MASTER_DB_FILE = "master_lead_database_perfect_sample_6may2026.xlsx"    # Delete me later

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
            "Australia", "Bangladesh", "China", "Hong Kong", 
            "India",
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
    df = normalize_columns(df)
    print(f"  Loaded {len(df)} rows\n")

    return df


# v2 : country column name fixed
def map_country_to_region(df):
    """
    Map Country → Region
    If unknown countries found:
        - Save to CSV
        - Stop execution
    """

    possible_cols = ["company country", "country", "country name"]

    country_col = next((c for c in df.columns if c in possible_cols), None)

    if not country_col:
        raise ValueError(f"No country column found in {possible_cols}")

    def get_region(country):
        if not country:
            return "UnknownRegion"

        country_clean = str(country).strip().lower()

        for region, countries in COUNTRY_REGION_MAP.items():
            if any(c.lower() == country_clean for c in countries):
                return region

        return "UnknownRegion"

    # ✅ Correct column name
    df["Region"] = df[country_col].apply(get_region)

    # 🔍 Find unknowns
    unknowns = df[df["Region"] == "UnknownRegion"][country_col].dropna().unique().tolist()

    unknowns = sorted([str(u).strip() for u in unknowns if str(u).strip()])

    # 🚨 If unknowns exist → STOP PIPELINE
    if unknowns:
        unknown_file = os.path.join(
            OUTPUT_DIR,
            f"unmapped_countries_{TIMESTAMP}.csv"
        )

        pd.DataFrame({"Unmapped Country": unknowns}).to_csv(unknown_file, index=False)

        print(f"\n  ❌❌❌❌ Unmapped Countries Found : {len(unknowns)}")
        print(f"  Saved → {os.path.basename(unknown_file)}")
        print("  👉 Fix mapping before running again.\n")

        exit(1)  # 🔥 HARD STOP

    return df

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
    print("  Scrapped Data Formatter for Ignite Sheet — Final Step 😎")
    print("═" * 62 + "\n")

    print("\n" + "═" * 62)
    print("  Scrapped Data Formatter for Ignite Sheet — Final Step (3 of 3) 😎")
    print("  Apollo Master Lead DB   →  LinkedIn Ignite Format")
    print("═" * 62)
    print(f"  Source file  : {MASTER_DB_FILE if USE_MASTER_DB else ENRICHED_FILE}")
    print(f"  Output file  : output/ignite_ready_contacts_<timestamp>.xlsx")
    print(f"  Mode         : {'Master DB' if USE_MASTER_DB else 'Enriched File'}")
    print("═" * 62 + "\n")

    df = load_input_data()

    # 🔥 STEP 1: Schema Mapping (CRITICAL FIX)
    df = map_master_to_campaign(df)

    # 🔥 STEP 2: Region Mapping
    df = map_country_to_region(df)

    # 🔥 STEP 3: Build Final Format
    campaign_df = build_campaign_format(df)

    # 🔥 STEP 4: Save
    campaign_df.to_excel(OUTPUT_FILE, index=False)

    print(f"  ✅ Campaign File Created → output/{os.path.basename(OUTPUT_FILE)}")
    print(f"     Total rows : {len(campaign_df)}")

    print("\n" + "═" * 62 + "\n")


if __name__ == "__main__":
    main()