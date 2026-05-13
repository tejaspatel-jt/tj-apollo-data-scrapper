#!/usr/bin/env python3
"""
═══════════════════════════════════════════════════════════════
  Apollo Contact Scraper — STEP 1: People Search
  Uses:
    POST /api/v1/mixed_people/api_search   (New Leads / Prospects)
    POST /api/v1/contacts/search           (Already Saved CRM Contacts)

  INPUT : inputs/companies_input.csv  (or hardcoded COMPANIES list)
  OUTPUT: output/apollo_people_search_DDMMMYYYY_HHMMSS.xlsx
            Sheet 1 — Contacts  (all results, ready for Step 2)
            Sheet 2 — Summary   (per-company breakdown)
          output/master_lead_database.xlsx
            Persistent across runs; never overwrites existing entries
═══════════════════════════════════════════════════════════════
"""

import requests
import pandas as pd
import time
import os
from datetime import datetime
import html   # for unescaping HTML entities in strings

# ─── resolve paths relative to this script's location ───────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUTS_DIR = os.path.join(SCRIPT_DIR, "..", "inputs")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ╔══════════════════════════════════════════════════════════════╗
# ║                     CONFIGURATION                           ║
# ╚══════════════════════════════════════════════════════════════╝

# API_KEY = "l4AkVpSXwV9fICkebUMAXw"        # Normal API key
API_KEY = "vWTBtYd1P9IpMLV-wghAzw"        # Your Apollo master API key

TIMESTAMP = datetime.now().strftime("%d%b%Y_%H%M%S").lower()   # e.g. 27apr2026_124055

COMPANIES_INPUT_FILE = "companies_input.csv"   # CSV with columns: name, domain
# Leave as "" to use the hardcoded COMPANIES list below instead

OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"apollo_people_search_{TIMESTAMP}.xlsx")
# Timestamp ensures output is never overwritten on repeated runs

# ── 9. MASTER DB ──────────────────────────────────────────────
# True  → update master_lead_database.xlsx after every run (original behaviour)
# False → skip master DB entirely — only the search output file is saved
CREATE_MASTER_DB = False

MASTER_DB = os.path.join(OUTPUT_DIR, "master_lead_database.xlsx")
# Persistent master database — updated on every run, never loses old entries

PER_PAGE      = 100   # Apollo max is 100 per page
MAX_PAGES     = 500   # Apollo hard cap (50k results); reduce to limit credits
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
    {"name": "Alara Imaging",                             "domain": "alaragateway.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Vertis Therapy",                            "domain": "vertistherapy.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Medical Informatics",                       "domain": "sickbay.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Pro EMS",                                   "domain": "proems.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "XP Health",                                 "domain": "xphealth.co", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Cordata Healthcare Innovations",            "domain": "cordatahealth.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Agamon Health",                             "domain": "agamonhealth.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "ImageMover",                                "domain": "imagemovermd.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Alloy Health",                              "domain": "myalloy.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "FHAS",                                      "domain": "fhas.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "MedArrive",                                 "domain": "medarrive.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Health Gorilla",                            "domain": "healthgorilla.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "The Wound Company",                         "domain": "thewound.co", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Envoy America",                             "domain": "envoyamerica.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Digital Health Strategies",                 "domain": "digitalhealthstrategies.com", "org_industry_input": "","org_employees_input": ""},
    {"name": "careMESH",                                  "domain": "caremesh.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Navis Clinical Laboratories",               "domain": "navisclinical.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Yosemite Pathology & Precision Pathology",  "domain": "ypmg.com", "org_industry_input": "", "org_employees_input": ""},
    {"name": "Sleep Reset",                               "domain": "thesleepreset.com", "org_industry_input": "", "org_employees_input": ""},
]


# ╔══════════════════════════════════════════════════════════════╗
# ║                     HELPER FUNCTIONS                        ║
# ╚══════════════════════════════════════════════════════════════╝

BASE_URL = "https://api.apollo.io/api/v1"

# v7 changes : input columns: company name, website, org industry, org employees
def load_companies():
    """Load companies from CSV file or fall back to hardcoded list."""
    
    file_path = os.path.join(INPUTS_DIR, COMPANIES_INPUT_FILE)

    if COMPANIES_INPUT_FILE and os.path.exists(file_path):
        df = pd.read_csv(file_path)
        df.columns = [c.strip().lower() for c in df.columns]

        
        # if "company name" not in df.columns:
        #     raise ValueError(f"'company name' column not found in {COMPANIES_INPUT_FILE}")
        # if "company domain" not in df.columns:
        #     raise ValueError(f"'company domain' column not found in {COMPANIES_INPUT_FILE}")
        # if "org industry" not in df.columns:
        #     raise ValueError(f"'org industry' column not found in {COMPANIES_INPUT_FILE}")
        # if "org employees" not in df.columns:
        #     raise ValueError(f"'org employees' column not found in {COMPANIES_INPUT_FILE}")

        # print(f"  Loaded {len(df)} companies from '{COMPANIES_INPUT_FILE}'")
        # # return [{"name": row[name_col], "domain": str(row["domain"]).strip()} for _, row in df.iterrows()]

        # return [
        #     {
        #         "name": str(row.get("company name", "")).strip(),
        #         "domain": str(row.get("company domain", "")).strip(),
        #         "org_industry_input": str(row.get("org industry", "")).strip(),
        #         "org_employees_input": str(row.get("org employees", "")).strip(),
        #     }
        #     for _, row in df.iterrows()
        # ]
    
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

        if name_col not in df.columns:
            raise ValueError(f"Missing required column: {name_col}")
        if domain_col not in df.columns:
            raise ValueError(f"Missing required column: {domain_col}")
        if industry_col not in df.columns:
            print(f"Warning: 'org industry' column not found. Industry input will be empty.")
        if employees_col not in df.columns:
            print(f"Warning: 'org employees' column not found. Employees input will be empty.")

        if not name_col or not domain_col:
            missing = []
            if not name_col:
                missing.append("company name")
            if not domain_col:
                missing.append("domain")

            raise ValueError(f"Missing required columns: {missing}")

        return [
            {
                "name": str(row.get(name_col, "")).strip(),
                "domain": str(row.get(domain_col, "")).strip(),
                "org_industry_input": str(row.get(industry_col, "")).strip() if industry_col else "",
                "org_employees_input": str(row.get(employees_col, "")).strip() if employees_col else "",
            }
            for _, row in df.iterrows()
        ]

    else:
        print(f"  Using hardcoded COMPANIES list ({len(COMPANIES)} companies)")
        return COMPANIES

# API - FETCH NEW PEOPLE, NOT SAVED IN CRM ( NO CREDITS CONSUMED )
def search_people(domain, company_name):
    """
    Paginate through /mixed_people/api_search for one domain.
    Returns new/prospect people not yet saved to Apollo CRM.
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
                "Content-Type":  "application/json",
                "accept":        "application/json",
                "x-api-key":     API_KEY,
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

        data        = response.json()
        people      = data.get("people", [])
        pagination  = data.get("pagination", {})
        total_pages = pagination.get("total_pages", 1)
        total       = pagination.get("total_entries", "?")

        print(f"    [NEW]  Page {page:>3}/{min(total_pages, MAX_PAGES)} — {len(people)} results (total ≈ {total})")
        all_people.extend(people)

        if page >= total_pages:
            break
        page += 1
        time.sleep(REQUEST_DELAY)

    return all_people


# API - FETCH ALREADY SAVED CONTACTS ( NO CREDITS CONSUMED )
def search_contacts(domain, company_name):
    """
    Paginate through /contacts/search for one domain.
    Returns contacts already saved in your Apollo CRM.
    These are returned without consuming enrichment credits.
    """
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
                "Cache-Control": "no-cache",
                "Content-Type":  "application/json",
                "accept":        "application/json",
                "x-api-key":     API_KEY,
            },
            timeout=30,
        )

        if response.status_code == 429:
            wait = int(response.headers.get("Retry-After", 15))
            print(f"    ⚠  CRM rate limited — waiting {wait}s...")
            time.sleep(wait)
            continue

        if response.status_code != 200:
            print(f"    ✗  CRM HTTP {response.status_code}: {response.text[:300]}")
            break

        data        = response.json()
        contacts    = data.get("contacts", [])
        pagination  = data.get("pagination", {})
        total_pages = pagination.get("total_pages", 1)
        total       = pagination.get("total_entries", "?")

        print(f"    [CRM]  Page {page:>3}/{min(total_pages, MAX_PAGES)} — {len(contacts)} results (total ≈ {total})")
        all_contacts.extend(contacts)

        if page >= total_pages:
            break
        page += 1
        time.sleep(REQUEST_DELAY)

    return all_contacts


# changes in v7: added company input data to flatten_person parameters; updated logic to use company dict instead of separate name/domain params
def flatten_person(person, company, source_type="NEW_PROSPECT"):
    """
    Flatten one Apollo person/contact object into a flat dict.

    source_type:
      "NEW_PROSPECT"  — from /mixed_people/api_search
      "CRM_CONTACT"   — from /contacts/search (already saved in Apollo)

    Columns are chosen so the row can be directly fed into Step 2
    (apollo_bulk_create.py) after filtering.

    bulk_create fields mapped:
      first_name, last_name, title, organization_name,
      linkedin_url, apollo_person_id (→ used for enrichment deduplication),
      website_url (company domain)
    """
    org = person.get("organization") or {}

    def normalize_domain(url):
        if not url:
            return ""
        return url.replace("http://", "").replace("https://", "").replace("www.", "").strip().lower()  

    # Extract company input data
    company_name_input = company.get("name", "")
    domain = company.get("domain", "")
    org_industry_input = company.get("org_industry_input", "")
    org_employees_input = company.get("org_employees_input", "")

    # /contacts/search returns person_id (not id) for the Apollo person record
    if source_type == "CRM_CONTACT":
        pid = person.get("person_id") or person.get("id", "")
    else:
        pid = person.get("id", "")

    # 🔥 Smart override logic
    org_industry = org_industry_input if org_industry_input else org.get("industry", "")
    org_employees = org_employees_input if org_employees_input else org.get("estimated_num_employees", "")

    return {
        # ── Identity ──────────────────────────────────────────────
        "apollo_person_id":  pid,                              # needed for Step 2 enrichment
        "first_name":        person.get("first_name", ""),
        "last_name":         person.get("last_name",  ""),
        # "full_name":         person.get("name",       ""),
        "full_name": (
            person.get("name")
            or 
            f"{person.get('first_name', '')} {person.get('last_name', '')}".strip()
        ),
        "title":             person.get("title",      ""),
        "seniority":         person.get("seniority",  ""),
        "linkedin_url":      person.get("linkedin_url", ""),

        # ── Contact (usually masked for NEW_PROSPECT at this stage) ──
        "email":             person.get("email",        ""),
        "email_status":      person.get("email_status", ""),

        # ── Location ──────────────────────────────────────────────
        "city":    person.get("city",    ""),
        "state":   person.get("state",   ""),
        "country": person.get("country", ""),

        # ── Source company (what you searched for) ─────────────────
        "searched_company": company_name_input,
        "searched_domain":  domain,

        # ── Apollo org data (for Step 2 bulk_create) ───────────────
        "organization_name": org.get("name", company_name_input),
        # "website_url":       org.get("website_url",domain),  # → bulk_create: website_url
        "website_url":       normalize_domain(org.get("website_url", domain)),
        "org_linkedin_url":  org.get("linkedin_url",             ""),
        "org_industry":      org_industry,
        "org_employees":     org_employees,
        "apollo_org_id":     org.get("id",                       ""),

        # 🔥 NEW: track origin of org data
        "org_data_source": "INPUT" if org_industry_input or org_employees_input else "APOLLO",

        # ── Source tracking ────────────────────────────────────────
        "source_type":      source_type,    # NEW_PROSPECT | CRM_CONTACT

        # "is_enriched":      True if person.get("email") else False,
        # "is_crm_contact":   source_type == "CRM_CONTACT",   # True = already in Apollo

        # "is_enriched": "true" if person.get("email") else "false",
        # "is_crm_contact": "true" if source_type == "CRM_CONTACT" else "false",

        # is_enriched is True always because CRM_CONTACTs are already enriched/saved in Apollo, and NEW_PROSPECTs with no email will be enriched in Step 2 and get an email then 
        "is_enriched": "true" if (source_type == "CRM_CONTACT" or person.get("email")) else "false",
        "is_crm_contact": "true" if source_type == "CRM_CONTACT" else "false",
    }


def save_to_excel(df):
    """
    Write output to Excel with two sheets:
      Contacts — all contacts (flat, ready for Step 2)
      Summary  — per-company + per-source-type breakdown
    """
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        # Sheet 1 — all contacts
        df.to_excel(writer, index=False, sheet_name="Contacts")

        # Sheet 2 — per-company summary (new + crm breakdown)
        summary = (
            df.groupby(["searched_company", "searched_domain", "source_type"])
            .agg(contacts_found=("apollo_person_id", "count"))
            .reset_index()
        )
        summary.to_excel(writer, index=False, sheet_name="Summary")

        # Auto-fit column widths on Contacts sheet
        ws = writer.sheets["Contacts"]
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    # print(f"\n  ✅  Saved {len(df)} contacts → '{OUTPUT_FILE}'")

    out_name = os.path.basename(OUTPUT_FILE)
    crm_count = (df["source_type"] == "CRM_CONTACT").sum()
    new_count = (df["source_type"] == "NEW_PROSPECT").sum()
    enriched_count = (df["is_enriched"].astype(str).str.lower() == "true").sum()

    print(f"\n  ✅  Searched API Output Saved → 'output/{out_name}'")
    print(f"            ├───── Total contacts     : {len(df)}")
    print(f"            ├───── CRM contacts       : {crm_count}  (already saved in Apollo)")
    print(f"            ├───── New prospects      : {new_count}  (not yet saved)")
    print(f"            └───── Pre-enriched       : {enriched_count}  (have email already)")


def update_master_db(current_df):
    """
    Merge current run into the persistent master_lead_database.xlsx.
    Existing entries (by apollo_person_id) are never overwritten.
    New entries from this run are appended.
    """
    if os.path.exists(MASTER_DB):
        master_df = pd.read_excel(MASTER_DB, dtype=str)

        # 🔥 Fix: align schema - start
        for col in current_df.columns:
            if col not in master_df.columns:
                master_df[col] = ""

        for col in master_df.columns:
            if col not in current_df.columns:
                current_df[col] = ""
        # 🔥 Fix: align schema - end

        before    = len(master_df)
        combined  = pd.concat([master_df, current_df.astype(str)], ignore_index=True)
        combined.drop_duplicates(subset=["apollo_person_id"], keep="first", inplace=True)
        combined.reset_index(drop=True, inplace=True)
        added = len(combined) - before
    else:
        combined = current_df.astype(str).copy()
        combined.drop_duplicates(subset=["apollo_person_id"], keep="first", inplace=True)
        before, added = 0, len(combined)

    combined.to_excel(MASTER_DB, index=False, engine="openpyxl")

    # print(f"  ✅  Master DB → '{MASTER_DB}'")
    # print(f"       Before: {before}  |  Added this run: {added}  |  Total: {len(combined)}")

    total_enriched = (combined["is_enriched"].astype(str).str.lower() == "true").sum()
    total_crm = (combined["is_crm_contact"].astype(str).str.lower() == "true").sum()
    to_enrich = len(combined) - total_enriched

    print(f"\n  🧠  Master DB Updated ✅")
    print(f"            ├───── Before this run    : {before}")
    print(f"            ├───── Added this run     : {added}")
    print(f"            ├───── Total records      : {len(combined)}")
    print(f"            ├───── Already enriched   : {total_enriched}  (will be skipped in Script 2)")
    print(f"            ├───── Needs enrichment   : {to_enrich}  (will be processed in Script 2)")
    print(f"            └───── CRM contacts       : {total_crm}")
    print(f"\n  👀  View Master DB → 'output/master_lead_database.xlsx'")


# ╔══════════════════════════════════════════════════════════════╗
# ║                        MAIN RUNNER                         ║
# ╚══════════════════════════════════════════════════════════════╝

def main():
    print("\n" + "═" * 62)
    print("  Apollo People Search Scraper — Step 1 of 2")
    print("  Searches: New Prospects + Already Saved CRM Contacts")
    print("═" * 62)

    companies = load_companies()
    print(f"  Titles filter  : {len(PERSON_TITLES)} roles")
    print(f"  Output file    : {OUTPUT_FILE}")
    print(f"  Master DB      : {MASTER_DB}")
    print("═" * 62 + "\n")

    all_rows = []

    for idx, company in enumerate(companies, 1):
        name   = company["name"]
        domain = company["domain"]

        print(f"[{idx:>2}/{len(companies)}] {name} ({domain})")

        # ── 1. Search already-saved CRM contacts (no enrichment credits) ──
        crm_contacts = search_contacts(domain, name)
        if crm_contacts:
            # rows = [flatten_person(c, name, domain, "CRM_CONTACT") for c in crm_contacts]
            rows = [flatten_person(c, company, "CRM_CONTACT") for c in crm_contacts]
            all_rows.extend(rows)
            print(f"    ✓  {len(rows)} CRM contacts found")
        else:
            print(f"    –  No CRM contacts found")

        # ── 2. Search new/prospect people ─────────────────────────────────
        new_people = search_people(domain, name)
        if new_people:
            # rows = [flatten_person(p, name, domain, "NEW_PROSPECT") for p in new_people]
            rows = [flatten_person(p, company, "NEW_PROSPECT") for p in new_people]
            all_rows.extend(rows)
            print(f"    ✓  {len(rows)} new prospects collected")
        else:
            print(f"    –  No new prospects found")

        print()
        time.sleep(REQUEST_DELAY)

    if not all_rows:
        print("\n  ⚠  No contacts found for any company.")
        return

    df = pd.DataFrame(all_rows)

    # 🔥 Fix common HTML entities in all string fields (e.g. &amp; → &)
    # df = df.applymap(lambda x: html.unescape(x) if isinstance(x, str) else x)
    df = df.map(lambda x: html.unescape(x) if isinstance(x, str) else x) # because applymap renamed to map in newer pandas versions, we can use map instead of applymap


    # Drop full duplicates (same apollo_person_id appearing under multiple domains)
    # CRM contacts take priority: sort so CRM_CONTACT rows come first before dedup
    df.sort_values(
        "source_type",
        key=lambda s: s.map({"CRM_CONTACT": 0, "NEW_PROSPECT": 1}),
        inplace=True,
    )
    df.drop_duplicates(subset=["apollo_person_id"], keep="first", inplace=True)
    df.reset_index(drop=True, inplace=True)

    print("─" * 62)
    crm_count  = (df["source_type"] == "CRM_CONTACT").sum()
    new_count  = (df["source_type"] == "NEW_PROSPECT").sum()
    print(f"  Total unique contacts  : {len(df)}")
    print(f"    ├─ CRM contacts      : {crm_count}  (already saved in Apollo)")
    print(f"    └─ New prospects     : {new_count}  (new leads, not yet saved)")

    save_to_excel(df)

    if CREATE_MASTER_DB:
        update_master_db(df)

    print()
    print("  ─── NEXT STEP ───────────────────────────────────────────")
    print(f"  1. Open '{OUTPUT_FILE}'")
    print("  2. Review the Contacts sheet — filter rows you want")
    print("  3. Save filtered rows as a new file")
    print("  4. Run Script 2 (apollo_bulk_create.py) to enrich + save")
    print("     Note: CRM_CONTACT rows are already saved — tracker will")
    print("     skip re-enriching them if already processed before.")
    print("─" * 62 + "\n")


if __name__ == "__main__":
    main()
