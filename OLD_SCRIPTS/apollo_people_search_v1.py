#!/usr/bin/env python3
"""
═══════════════════════════════════════════════════════════════
  Apollo Contact Scraper — STEP 1: People Search
  Uses: POST /api/v1/mixed_people/api_search

  INPUT : inputs/companies_input.csv   (or hardcoded COMPANIES list)
  OUTPUT: output/apollo_people_search_DDMMMYYYY_HHMMSS.xlsx
          Excel file ready to filter, then feed into Step 2
          (contacts/bulk_create) for email enrichment
═══════════════════════════════════════════════════════════════
"""

import requests
import pandas as pd
import time
import os
from datetime import datetime

# ─── resolve paths relative to this script's location ───────────────────────
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUTS_DIR   = os.path.join(SCRIPT_DIR, "..", "inputs")
OUTPUT_DIR   = os.path.join(SCRIPT_DIR, "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ╔══════════════════════════════════════════════════════════════╗
# ║                     CONFIGURATION                           ║
# ╚══════════════════════════════════════════════════════════════╝

API_KEY = "l4AkVpSXwV9fICkebUMAXw"       # Your Apollo master API key

TIMESTAMP = datetime.now().strftime("%d%b%Y_%H%M%S").lower()   # e.g. 27apr2026_124055

COMPANIES_INPUT_FILE = "companies_input.csv"  # CSV with columns: name, domain
#   Leave as "" to use the hardcoded COMPANIES list below instead

OUTPUT_FILE     = os.path.join(OUTPUT_DIR, f"apollo_people_search_{TIMESTAMP}.xlsx")
#   Timestamp ensures output is never overwritten on repeated runs

PER_PAGE = 100        # Apollo max is 100 per page
MAX_PAGES = 500       # Apollo hard cap (50k results); reduce to limit credits
REQUEST_DELAY = 1.2   # Seconds to wait between API calls (avoid rate-limiting)

# Decision-maker title filters (edit freely)
PERSON_TITLES = [
    "ceo",
    "co-founder",
    "founder",
    "co-founder and ceo",
    "cto",
    "chief technology officer",
    "vp engineering",
    "vice president engineering",
    "vp of engineering",
    "director of engineering",
    "engineering director",
    "head of engineering",
    "head of platform engineering",
    "platform engineering",
    "head of quality assurance",
    "qa director",
    "director of qa",
    "head of qa",
    "head of product",
    "vp product",
    "vice president product",
    "director of product",
    "chief product officer",
    "director of product and engineering",
    "product engineering",
    "senior vice president engineering",
    "svp engineering",
    "director of research and development",
    "head of r&d",
    "field cto",
    "vpe",
    "vp eng",
]

# Hardcoded fallback company list (used when COMPANIES_INPUT_FILE = "")
COMPANIES = [
    {"name": "Alara Imaging",                           "domain": "alaragateway.com"},
    {"name": "Vertis Therapy",                          "domain": "vertistherapy.com"},
    {"name": "Medical Informatics",                     "domain": "sickbay.com"},
    {"name": "Pro EMS",                                 "domain": "proems.com"},
    {"name": "XP Health",                               "domain": "xphealth.co"},
    {"name": "Cordata Healthcare Innovations",          "domain": "cordatahealth.com"},
    {"name": "Agamon Health",                           "domain": "agamonhealth.com"},
    {"name": "ImageMover",                              "domain": "imagemovermd.com"},
    {"name": "Alloy Health",                            "domain": "myalloy.com"},
    {"name": "FHAS",                                    "domain": "fhas.com"},
    {"name": "MedArrive",                               "domain": "medarrive.com"},
    {"name": "Health Gorilla",                          "domain": "healthgorilla.com"},
    {"name": "The Wound Company",                       "domain": "thewound.co"},
    {"name": "Envoy America",                           "domain": "envoyamerica.com"},
    {"name": "Digital Health Strategies",               "domain": "digitalhealthstrategies.com"},
    {"name": "Synergy Billing",                         "domain": "synergybilling.com"},
    {"name": "careMESH",                                "domain": "caremesh.com"},
    {"name": "Navis Clinical Laboratories",             "domain": "navisclinical.com"},
    {"name": "Yosemite Pathology & Precision Pathology","domain": "ypmg.com"},
    {"name": "Sleep Reset",                             "domain": "thesleepreset.com"},
]

# ╔══════════════════════════════════════════════════════════════╗
# ║                     HELPER FUNCTIONS                        ║
# ╚══════════════════════════════════════════════════════════════╝

BASE_URL = "https://api.apollo.io/api/v1"

def load_companies():
    """Load companies from CSV file or fall back to hardcoded list."""
    if COMPANIES_INPUT_FILE and os.path.exists(COMPANIES_INPUT_FILE):
        df = pd.read_csv(COMPANIES_INPUT_FILE)
        df.columns = [c.strip().lower() for c in df.columns]
        if "domain" not in df.columns:
            raise ValueError(f"'domain' column not found in {COMPANIES_INPUT_FILE}")
        name_col = "name" if "name" in df.columns else "domain"
        print(f"  Loaded {len(df)} companies from '{COMPANIES_INPUT_FILE}'")
        return [{"name": row[name_col], "domain": str(row["domain"]).strip()} for _, row in df.iterrows()]
    else:
        print(f"  Using hardcoded COMPANIES list ({len(COMPANIES)} companies)")
        return COMPANIES


def search_people(domain, company_name):
    """
    Paginate through /mixed_people/api_search for one domain.
    Returns list of raw Apollo person objects.
    """
    all_people = []
    page = 1

    while page <= MAX_PAGES:
        response = requests.post(
            f"{BASE_URL}/mixed_people/api_search",
            params={"per_page": PER_PAGE, "page": page},
            json={
                "q_organization_domains_list": [domain],
                "person_titles": PERSON_TITLES,
            },
            headers={
                "Cache-Control": "no-cache",
                "Content-Type": "application/json",
                "accept": "application/json",
                "x-api-key": API_KEY,
            },
            timeout=30,
        )

        if response.status_code == 429:
            wait = int(response.headers.get("Retry-After", 15))
            print(f"    ⚠  Rate limited — waiting {wait}s...")
            time.sleep(wait)
            continue

        if response.status_code != 200:
            print(f"    ✗  HTTP {response.status_code}: {response.text[:300]}")
            break

        data = response.json()
        people      = data.get("people", [])
        # total_entries      = data.get("total_entries", 0)
        pagination  = data.get("pagination", {})
        total_pages = pagination.get("total_pages", 1)
        total       = pagination.get("total_entries", "?")

        print(f"    Page {page:>3}/{min(total_pages, MAX_PAGES)} — {len(people)} results  (total ≈ {total})")
        # print(f"    Page {page:>3}/{min(total_pages, MAX_PAGES)} — {len(people)} results  (total ≈ {total_entries})")
        all_people.extend(people)

        if page >= total_pages:
            break
        page += 1
        time.sleep(REQUEST_DELAY)

    return all_people


def flatten_person(person, company_name, domain):
    """
    Flatten one Apollo person object into a flat dict.

    Columns are chosen so the row can be directly fed into
    contacts/bulk_create after filtering — just keep the rows
    you want and pass them to Script 2 (bulk_create).

    bulk_create fields mapped:
      first_name, last_name, title, organization_name,
      linkedin_url, apollo_person_id (→ used for deduplication),
      website_url (company domain)
    """
    org = person.get("organization") or {}

    return {
        # ── Identity ─────────────────────────────────────────
        "apollo_person_id":     person.get("id", ""),          # needed for bulk_create deduplication
        "first_name":           person.get("first_name", ""),
        "last_name":            person.get("last_name", ""),
        "full_name":            person.get("name", ""),
        "title":                person.get("title", ""),
        "seniority":            person.get("seniority", ""),
        "linkedin_url":         person.get("linkedin_url", ""),

        # ── Contact (usually masked at this stage) ────────────
        "email":                person.get("email", ""),
        "email_status":         person.get("email_status", ""),

        # ── Location ─────────────────────────────────────────
        "city":                 person.get("city", ""),
        "state":                person.get("state", ""),
        "country":              person.get("country", ""),

        # ── Source company (what you searched for) ────────────
        "searched_company":     company_name,
        "searched_domain":      domain,

        # ── Apollo org data (for bulk_create) ─────────────────
        "organization_name":    org.get("name", ""),           # → bulk_create: organization_name
        "website_url":          org.get("website_url", ""),    # → bulk_create: website_url
        "org_linkedin_url":     org.get("linkedin_url", ""),
        "org_industry":         org.get("industry", ""),
        "org_employees":        org.get("estimated_num_employees", ""),
        "apollo_org_id":        org.get("id", ""),
    }


def save_to_excel(df):
    """Write output to Excel with two sheets: Contacts + Summary."""
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        # Sheet 1 — all contacts
        df.to_excel(writer, index=False, sheet_name="Contacts")

        # Sheet 2 — per-company summary
        summary = (
            df.groupby(["searched_company", "searched_domain"])
              .agg(contacts_found=("apollo_person_id", "count"))
              .reset_index()
        )
        summary.to_excel(writer, index=False, sheet_name="Summary")

        # Auto-fit column widths on Contacts sheet
        ws = writer.sheets["Contacts"]
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    print(f"\n  ✅  Saved {len(df)} contacts → '{OUTPUT_FILE}'")


# ╔══════════════════════════════════════════════════════════════╗
# ║                        MAIN RUNNER                           ║
# ╚══════════════════════════════════════════════════════════════╝

def main():
    print("\n" + "═" * 62)
    print("  Apollo People Search Scraper — Step 1 of 2")
    print("═" * 62)

    companies = load_companies()
    print(f"  Titles filter : {len(PERSON_TITLES)} roles")
    print(f"  Output file   : {OUTPUT_FILE}")
    print("═" * 62 + "\n")

    all_rows = []

    for idx, company in enumerate(companies, 1):
        name   = company["name"]
        domain = company["domain"]

        print(f"[{idx:>2}/{len(companies)}]  {name}  ({domain})")
        people = search_people(domain, name)

        if not people:
            print(f"         No results.\n")
            continue

        rows = [flatten_person(p, name, domain) for p in people]
        print(f"         ✓ {len(rows)} contacts collected\n")
        all_rows.extend(rows)
        time.sleep(REQUEST_DELAY)

    if not all_rows:
        print("\n  ⚠  No contacts found for any company.")
        return

    df = pd.DataFrame(all_rows)

    # Drop full duplicates (same person appearing under multiple domains)
    df.drop_duplicates(subset=["apollo_person_id"], keep="first", inplace=True)
    df.reset_index(drop=True, inplace=True)

    print("─" * 62)
    print(f"  Total unique contacts : {len(df)}")
    save_to_excel(df)

    print()
    print("  ─── NEXT STEP ───────────────────────────────────────────")
    print(f"  1. Open '{OUTPUT_FILE}'")
    print("  2. Filter / keep only the contacts you want (Contacts sheet)")
    print("  3. Save filtered rows as 'bulk_create_input.xlsx'")
    print("  4. Run Script 2 (apollo_bulk_create.py) to enrich emails")
    print("─" * 62 + "\n")


if __name__ == "__main__":
    main()
