#!/usr/bin/env python3
"""
═══════════════════════════════════════════════════════════════
  Apollo Contact Scraper — STEP 1: People Search
  Uses: 
    POST /api/v1/mixed_people/api_search (New Leads)
    POST /api/v1/contacts/search         (Already Saved Contacts)

  INPUT : inputs/companies_input.csv   (or hardcoded COMPANIES list)
  OUTPUT: output/apollo_people_search_DDMMMYYYY_HHMMSS.xlsx
          output/master_lead_database.xlsx (Maintains state across runs)
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
# os.makedirs(INPUTS_DIR, exist_ok=True)

# ╔══════════════════════════════════════════════════════════════╗
# ║                     CONFIGURATION                           ║
# ╚══════════════════════════════════════════════════════════════╝

# API_KEY = "l4AkVpSXwV9fICkebUMAXw"       # Your Apollo master API key
API_KEY = "vWTBtYd1P9IpMLV-wghAzw"        # Your Apollo master API key

TIMESTAMP = datetime.now().strftime("%d%b%Y_%H%M%S").lower()   # e.g. 27apr2026_124055

COMPANIES_INPUT_FILE = "companies_input.csv"  # CSV with columns: name, domain

OUTPUT_FILE     = os.path.join(OUTPUT_DIR, f"apollo_people_search_{TIMESTAMP}.xlsx")
MASTER_DB       = os.path.join(OUTPUT_DIR, "master_lead_database.xlsx")

PER_PAGE = 100        # Apollo max is 100 per page
MAX_PAGES = 500       # Apollo hard cap (50k results)
REQUEST_DELAY = 1.2   # Seconds to wait between API calls

# Decision-maker title filters (kept exactly as provided)
PERSON_TITLES = [
    "ceo", "co-founder", "founder", "co-founder and ceo", "cto",
    "chief technology officer", "vp engineering", "vice president engineering",
    "vp of engineering", "director of engineering", "engineering director",
    "head of engineering", "head of platform engineering", "platform engineering",
    "head of quality assurance", "qa director", "director of qa", "head of qa",
    "head of product", "vp product", "vice president product", "director of product",
    "chief product officer", "director of product and engineering", "product engineering",
    "senior vice president engineering", "svp engineering", "director of research and development",
    "head of r&d", "field cto", "vpe", "vp eng",
]

# Hardcoded fallback company list
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
    # path = os.path.join(OUTPUT_DIR, BULK_ENRICH_INPUT_FILE)
    # if COMPANIES_INPUT_FILE and os.path.exists(os.path.join(INPUTS_DIR, COMPANIES_INPUT_FILE)):
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
    """Paginate through /mixed_people/api_search for new leads."""
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
        people = data.get("people", [])
        pagination = data.get("pagination", {})
        total_pages = pagination.get("total_pages", 1)
        
        all_people.extend(people)
        if page >= total_pages: break
        page += 1
        time.sleep(REQUEST_DELAY)

    return all_people


def search_contacts(domain):
    """Paginate through /contacts/search for ALREADY SAVED people."""
    all_contacts = []
    page = 1

    while page <= MAX_PAGES:
        response = requests.post(
            f"{BASE_URL}/contacts/search",
            json={
                "q_organization_domains_list": [domain],
                "person_titles": PERSON_TITLES,
                "page": page,
            },
            headers={
                "x-api-key": API_KEY,
                "Content-Type": "application/json",
            },
            timeout=30,
        )

        if response.status_code == 429:
            wait = int(response.headers.get("Retry-After", 15))
            print(f"    ⚠  CRM Rate limited — waiting {wait}s...")
            time.sleep(wait)
            continue

        if response.status_code != 200:
            break

        data = response.json()
        contacts = data.get("contacts", [])
        all_contacts.extend(contacts)
        
        if page >= data.get("pagination", {}).get("total_pages", 1): break
        page += 1
        time.sleep(REQUEST_DELAY)

    return all_contacts


def flatten_person(person, company_name, domain, source_type="NEW_PROSPECT"):
    """Flatten one Apollo object into a flat dict with status tracking."""
    org = person.get("organization") or {}
    # Use person_id for CRM contacts, id for new prospects
    pid = person.get("id") if source_type == "NEW_PROSPECT" else person.get("person_id")
    
    return {
        "apollo_person_id":     pid,
        "first_name":           person.get("first_name", ""),
        "last_name":            person.get("last_name", ""),
        "full_name":            person.get("name", ""),
        "title":                person.get("title", ""),
        "seniority":            person.get("seniority", ""),
        "linkedin_url":         person.get("linkedin_url", ""),
        "email":                person.get("email", ""),
        "email_status":         person.get("email_status", ""),
        "city":                 person.get("city", ""),
        "state":                person.get("state", ""),
        "country":              person.get("country", ""),
        "searched_company":     company_name,
        "searched_domain":      domain,
        "organization_name":    org.get("name", ""),
        "website_url":          org.get("website_url", ""),
        "source_type":          source_type,
        "is_enriched":          True if person.get("email") else False,
        "is_saved_to_crm":      True if source_type == "CRM_CONTACT" else False
    }


def save_to_excel(df):
    """Write output to Excel with Contacts + Summary sheets."""
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Contacts")
        summary = df.groupby(["searched_company", "source_type"]).size().reset_index(name='count')
        summary.to_excel(writer, index=False, sheet_name="Summary")


# ╔══════════════════════════════════════════════════════════════╗
# ║                        MAIN RUNNER                           ║
# ╚══════════════════════════════════════════════════════════════╝

def main():
    print("\n" + "═" * 62)
    print("  Apollo People Search Scraper (Enhanced with Contact Search)")
    print("═" * 62)

    companies = load_companies()
    all_rows = []

    if os.path.exists(MASTER_DB):
        master_df = pd.read_excel(MASTER_DB)
        print(f"  Loaded Master DB: {len(master_df)} existing entries.")
    else:
        master_df = pd.DataFrame()

    for idx, company in enumerate(companies, 1):
        name, domain = company["name"], company["domain"]
        print(f"[{idx:>2}/{len(companies)}] Processing {name}...")

        # 1. Search CRM (Saved)
        contacts = search_contacts(domain)
        if contacts:
            all_rows.extend([flatten_person(c, name, domain, "CRM_CONTACT") for c in contacts])
            print(f"         ✓ {len(contacts)} existing contacts found")

        # 2. Search People (New)
        people = search_people(domain, name)
        if people:
            all_rows.extend([flatten_person(p, name, domain, "NEW_PROSPECT") for p in people])
            print(f"         ✓ {len(people)} new prospects collected")
        
        time.sleep(REQUEST_DELAY)

    if not all_rows:
        print("\n  ⚠ No results found.")
        return

    current_run_df = pd.DataFrame(all_rows)
    current_run_df.drop_duplicates(subset=["apollo_person_id"], keep="first", inplace=True)
    
    # Update Master Lead Database
    final_master = pd.concat([master_df, current_run_df]).drop_duplicates(subset=["apollo_person_id"], keep="first")
    final_master.to_excel(MASTER_DB, index=False)
    
    save_to_excel(current_run_df)

    print("─" * 62)
    print(f"  Master DB Updated: {MASTER_DB}")
    print(f"  Run Output Saved: {OUTPUT_FILE}")
    print("═" * 62 + "\n")

if __name__ == "__main__":
    main()