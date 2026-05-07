# Write the complete unified Apollo Master Utility script
# script = r'''#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════════╗
║         APOLLO CONTACT SCRAPER — MASTER UTILITY                     ║
║         Centralized, Credit-Safe, Full-Field Pipeline               ║
╠══════════════════════════════════════════════════════════════════════╣
║  STEP 1 │ People Search     │ POST /api/v1/mixed_people/api_search  ║
║  STEP 2 │ Enrich (NEW only) │ POST /api/v1/people/bulk_match        ║
║  STEP 3 │ Save to Apollo    │ POST /api/v1/contacts/bulk_create     ║
╠══════════════════════════════════════════════════════════════════════╣
║  CENTRAL DB : output/apollo_contacts_db.xlsx                        ║
║    • Tracks is_enriched + is_saved_to_apollo per contact            ║
║    • Skip enrichment API for already-enriched contacts (save $$$)   ║
║    • Skip CRM save for already-saved contacts                       ║
╠══════════════════════════════════════════════════════════════════════╣
║  INPUT  : inputs/companies_input.csv  (columns: name, domain)       ║
║  OUTPUT : output/apollo_contacts_db.xlsx  (master, always updated)  ║
║           output/apollo_export_DDMMMYYYY_HHMMSS.xlsx (run snapshot) ║
╚══════════════════════════════════════════════════════════════════════╝
"""

import requests
import pandas as pd
import time
import os
from datetime import datetime

# ─── Resolve paths ──────────────────────────────────────────────────────────
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
INPUTS_DIR  = os.path.join(SCRIPT_DIR, "..", "inputs")
OUTPUT_DIR  = os.path.join(SCRIPT_DIR, "..", "output")
os.makedirs(OUTPUT_DIR,  exist_ok=True)
os.makedirs(INPUTS_DIR,  exist_ok=True)

# ╔══════════════════════════════════════════════════════════════════════╗
# ║  CONFIGURATION  — edit this section only                            ║
# ╚══════════════════════════════════════════════════════════════════════╝

API_KEY = "YOUR_APOLLO_API_KEY"          # ← paste your Apollo API key here

# ── Which steps to run ───────────────────────────────────────────────────────
# Options: "1"  = People Search only
#          "2"  = Enrich only  (unenriched contacts in DB)
#          "3"  = CRM Save only (enriched-but-unsaved contacts in DB)
#          "2+3"= Enrich then auto-save (recommended for daily runs)
#          "ALL"= Run all three steps end-to-end
STEPS_TO_RUN = "ALL"

# ── File settings ────────────────────────────────────────────────────────────
COMPANIES_INPUT_FILE  = "companies_input.csv"   # inside inputs/ folder
                                                # columns: name, domain
DB_FILE               = os.path.join(OUTPUT_DIR, "apollo_contacts_db.xlsx")
TIMESTAMP             = datetime.now().strftime("%d%b%Y_%H%M%S").lower()
SNAPSHOT_FILE         = os.path.join(OUTPUT_DIR, f"apollo_export_{TIMESTAMP}.xlsx")

# ── People Search (Step 1) settings ─────────────────────────────────────────
PER_PAGE      = 100     # Apollo max 100 per page
MAX_PAGES     = 10      # Reduce to limit credits on large companies
REQUEST_DELAY = 1.2     # Seconds between API calls

PERSON_TITLES = [
    "ceo", "co-founder", "founder", "co-founder and ceo",
    "cto", "chief technology officer",
    "vp engineering", "vice president engineering", "vp of engineering",
    "director of engineering", "engineering director", "head of engineering",
    "head of platform engineering", "platform engineering",
    "head of quality assurance", "qa director", "director of qa", "head of qa",
    "head of product", "vp product", "vice president product",
    "director of product", "chief product officer",
    "director of product and engineering", "product engineering",
    "senior vice president engineering", "svp engineering",
    "director of research and development", "head of r&d",
    "field cto", "vpe", "vp eng",
]

# ── Enrichment (Step 2) settings ─────────────────────────────────────────────
ENRICH_BATCH_SIZE         = 10      # Apollo hard limit for bulk_match
REVEAL_PERSONAL_EMAILS    = True    # Uses extra credits
REVEAL_PHONE_NUMBER       = False   # Uses extra credits

# ── CRM Save (Step 3) settings ────────────────────────────────────────────────
CRM_BATCH_SIZE = 100                # Apollo limit for bulk_create
RUN_DEDUPE     = True               # Prevents duplicates in Apollo CRM
LABEL_NAMES    = []                 # e.g. ["Decision Maker", "Apollo Scraped"]
AUTO_SAVE_ENRICHED = True           # True = save contacts right after enrichment
                                    # (no manual review step needed)

# ── Hardcoded fallback companies (used when companies_input.csv not found) ───
COMPANIES_FALLBACK = [
    {"name": "Apollo.io",   "domain": "apollo.io"},
    {"name": "HubSpot",     "domain": "hubspot.com"},
]

# ─────────────────────────────────────────────────────────────────────────────
BASE_URL = "https://api.apollo.io/api/v1"

# ╔══════════════════════════════════════════════════════════════════════╗
# ║  DATABASE SCHEMA                                                     ║
# ╚══════════════════════════════════════════════════════════════════════╝

# Status + tracking columns (always present in DB)
STATUS_COLUMNS = [
    "is_enriched",            # bool — True after bulk_match succeeds
    "is_saved_to_apollo",     # bool — True after bulk_create succeeds
    "enriched_at",            # datetime string
    "saved_at",               # datetime string
    "step1_added_at",         # datetime string — when row first appeared
    "apollo_crm_contact_id",  # ID returned by bulk_create (CRM contact ID)
]

# All enriched data columns (from bulk_match response)
ENRICHED_COLUMNS = [
    "First Name", "Last Name", "Title", "Company Name",
    "Company Name for Emails", "Email", "Email Status",
    "Primary Email Source", "Primary Email Verification Source",
    "Email Confidence", "Primary Email Catch-all Status",
    "Primary Email Last Verified At",
    "Seniority", "Departments", "Sub Departments",
    "Contact Owner", "Work Direct Phone", "Home Phone", "Mobile Phone",
    "Corporate Phone", "Other Phone", "Do Not Call",
    "Stage", "Lists", "Last Contacted", "Account Owner",
    "# Employees", "Industry", "Keywords",
    "Person Linkedin Url", "Website", "Company Linkedin Url",
    "Facebook Url", "Twitter Url",
    "City", "State", "Country",
    "Company Address", "Company City", "Company State", "Company Country",
    "Company Phone", "Technologies", "Annual Revenue", "Total Funding",
    "Latest Funding", "Latest Funding Amount", "Last Raised At",
    "Subsidiary of", "Subsidiary of (Organization ID)",
    "Email Sent", "Email Open", "Email Bounced", "Replied", "Demoed",
    "Number of Retail Locations",
    "Apollo Contact Id", "Apollo Account Id",
    "Secondary Email", "Secondary Email Source", "Secondary Email Status",
    "Secondary Email Verification Source",
    "Tertiary Email", "Tertiary Email Source",
    "Tertiary Email Status", "Tertiary Email Verification Source",
    "Primary Intent Topic", "Primary Intent Score",
    "Secondary Intent Topic", "Secondary Intent Score",
    "Qualify Contact", "photo_url", "_org_id",
]

# Step-1 raw columns (before enrichment)
STEP1_COLUMNS = [
    "apollo_person_id", "first_name", "last_name", "full_name",
    "title", "seniority", "linkedin_url", "email", "email_status",
    "city", "state", "country", "searched_company", "searched_domain",
    "organization_name", "website_url", "org_linkedin_url",
    "org_industry", "org_employees", "apollo_org_id",
]

DB_ALL_COLUMNS = STEP1_COLUMNS + ENRICHED_COLUMNS + STATUS_COLUMNS


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  DATABASE HELPERS                                                    ║
# ╚══════════════════════════════════════════════════════════════════════╝

def load_db():
    """Load central database Excel. Creates empty DB if not found."""
    if os.path.exists(DB_FILE):
        df = pd.read_excel(DB_FILE, sheet_name="Database", dtype=str)
        df.fillna("", inplace=True)
        # Ensure all expected columns exist
        for col in DB_ALL_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        print(f"  📂 Loaded DB: {len(df)} contacts from '{DB_FILE}'")
    else:
        df = pd.DataFrame(columns=DB_ALL_COLUMNS)
        print(f"  📂 No existing DB found — starting fresh.")
    return df


def save_db(df, run_log_entry=None):
    """Save central database to Excel with Database + Run Log sheets."""
    # Load existing run log if present
    run_log_rows = []
    if os.path.exists(DB_FILE):
        try:
            run_log_rows = pd.read_excel(DB_FILE, sheet_name="Run Log", dtype=str).to_dict("records")
        except Exception:
            pass

    if run_log_entry:
        run_log_rows.append(run_log_entry)

    run_log_df = pd.DataFrame(run_log_rows) if run_log_rows else pd.DataFrame(
        columns=["timestamp", "step", "action", "count", "notes"]
    )

    # Ensure all DB columns are present
    for col in DB_ALL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Database")
        run_log_df.to_excel(writer, index=False, sheet_name="Run Log")

        # Auto-fit widths
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 45)

    enriched_count  = (df["is_enriched"] == "True").sum()
    saved_count     = (df["is_saved_to_apollo"] == "True").sum()
    print(f"\n  💾 DB saved → '{DB_FILE}'")
    print(f"     Total contacts : {len(df)}")
    print(f"     Enriched       : {enriched_count}")
    print(f"     Saved to Apollo: {saved_count}")


def save_snapshot(df):
    """Save a timestamped read-only snapshot for this run."""
    enriched_df  = df[df["is_enriched"] == "True"].reset_index(drop=True)
    unenriched_df = df[df["is_enriched"] != "True"].reset_index(drop=True)
    saved_df     = df[df["is_saved_to_apollo"] == "True"].reset_index(drop=True)
    pending_df   = df[(df["is_enriched"] == "True") &
                      (df["is_saved_to_apollo"] != "True")].reset_index(drop=True)

    with pd.ExcelWriter(SNAPSHOT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer,          index=False, sheet_name="All Contacts")
        enriched_df.to_excel(writer, index=False, sheet_name="Enriched")
        unenriched_df.to_excel(writer, index=False, sheet_name="Not Enriched Yet")
        saved_df.to_excel(writer,    index=False, sheet_name="Saved to Apollo")
        pending_df.to_excel(writer,  index=False, sheet_name="Pending CRM Save")

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 45)

    print(f"  📸 Snapshot saved → '{SNAPSHOT_FILE}'")


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  API HELPERS                                                         ║
# ╚══════════════════════════════════════════════════════════════════════╝

def _headers():
    return {
        "Cache-Control": "no-cache",
        "Content-Type":  "application/json",
        "Accept":        "application/json",
        "x-api-key":     API_KEY,
    }


def api_people_search(domain, page):
    """One page of /mixed_people/api_search for a domain."""
    resp = requests.post(
        f"{BASE_URL}/mixed_people/api_search",
        params={"per_page": PER_PAGE, "page": page},
        json={
            "q_organization_domains_list": [domain],
            "person_titles": PERSON_TITLES,
        },
        headers=_headers(),
        timeout=30,
    )
    if resp.status_code == 429:
        wait = int(resp.headers.get("Retry-After", 30))
        print(f"\n    ⚠ Rate limited — waiting {wait}s...")
        time.sleep(wait)
        return api_people_search(domain, page)
    resp.raise_for_status()
    return resp.json()


def api_bulk_match(person_ids):
    """POST /people/bulk_match for up to 10 IDs."""
    resp = requests.post(
        f"{BASE_URL}/people/bulk_match",
        params={
            "reveal_personal_emails": str(REVEAL_PERSONAL_EMAILS).lower(),
            "reveal_phone_number":    str(REVEAL_PHONE_NUMBER).lower(),
        },
        json={"details": [{"id": pid} for pid in person_ids]},
        headers=_headers(),
        timeout=60,
    )
    if resp.status_code == 429:
        wait = int(resp.headers.get("Retry-After", 30))
        print(f"\n    ⚠ Rate limited — waiting {wait}s...")
        time.sleep(wait)
        return api_bulk_match(person_ids)
    resp.raise_for_status()
    data = resp.json()
    return data.get("matches", []), data


def api_bulk_create(contacts_payload):
    """POST /contacts/bulk_create for up to 100 contacts."""
    body = {"contacts": contacts_payload, "run_dedupe": RUN_DEDUPE}
    if LABEL_NAMES:
        body["append_label_names"] = LABEL_NAMES
    resp = requests.post(
        f"{BASE_URL}/contacts/bulk_create",
        json=body,
        headers=_headers(),
        timeout=60,
    )
    if resp.status_code == 429:
        wait = int(resp.headers.get("Retry-After", 30))
        print(f"\n    ⚠ Rate limited — waiting {wait}s...")
        time.sleep(wait)
        return api_bulk_create(contacts_payload)
    resp.raise_for_status()
    return resp.json()


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  DATA PARSERS                                                        ║
# ╚══════════════════════════════════════════════════════════════════════╝

def flatten_person_step1(person, company_name, domain):
    """Flatten a raw Apollo person (Step 1) into a DB row dict."""
    org = person.get("organization") or {}
    return {
        "apollo_person_id":  person.get("id", ""),
        "first_name":        person.get("first_name", ""),
        "last_name":         person.get("last_name", ""),
        "full_name":         person.get("name", ""),
        "title":             person.get("title", ""),
        "seniority":         person.get("seniority", ""),
        "linkedin_url":      person.get("linkedin_url", ""),
        "email":             person.get("email", ""),
        "email_status":      person.get("email_status", ""),
        "city":              person.get("city", ""),
        "state":             person.get("state", ""),
        "country":           person.get("country", ""),
        "searched_company":  company_name,
        "searched_domain":   domain,
        "organization_name": org.get("name", ""),
        "website_url":       org.get("website_url", ""),
        "org_linkedin_url":  org.get("linkedin_url", ""),
        "org_industry":      org.get("industry", ""),
        "org_employees":     org.get("estimated_num_employees", ""),
        "apollo_org_id":     org.get("id", ""),
        "is_enriched":          "False",
        "is_saved_to_apollo":   "False",
        "enriched_at":          "",
        "saved_at":             "",
        "step1_added_at":       datetime.now().isoformat(timespec="seconds"),
        "apollo_crm_contact_id": "",
    }


def parse_phones(match):
    phones  = match.get("phone_numbers") or []
    buckets = {
        "Work Direct Phone": "", "Home Phone": "",
        "Mobile Phone": "", "Corporate Phone": "", "Other Phone": "",
    }
    type_map = {
        "work_direct": "Work Direct Phone", "home": "Home Phone",
        "mobile": "Mobile Phone", "corporate": "Corporate Phone", "other": "Other Phone",
    }
    for p in phones:
        raw   = p.get("sanitized_number") or p.get("raw_number", "")
        ptype = (p.get("type") or "other").lower().replace(" ", "_")
        bucket = type_map.get(ptype, "Other Phone")
        if raw and not buckets[bucket]:
            buckets[bucket] = raw
    return buckets


def parse_emails(match):
    emails = sorted(match.get("contact_emails") or [], key=lambda x: x.get("position", 99))
    def get(idx, field):
        if idx < len(emails):
            v = emails[idx].get(field)
            return v if v is not None else ""
        return ""
    return {
        "Email":                                get(0, "email") or match.get("email", ""),
        "Email Status":                         get(0, "email_status") or match.get("email_status", ""),
        "Primary Email Source":                 get(0, "email_source"),
        "Primary Email Verification Source":    get(0, "verification_source"),
        "Email Confidence":                     match.get("extrapolated_email_confidence", ""),
        "Primary Email Catch-all Status":       get(0, "catch_all_status"),
        "Primary Email Last Verified At":       get(0, "last_verified_at"),
        "Secondary Email":                      get(1, "email"),
        "Secondary Email Source":               get(1, "email_source"),
        "Secondary Email Status":               get(1, "email_status"),
        "Secondary Email Verification Source":  get(1, "verification_source"),
        "Tertiary Email":                       get(2, "email"),
        "Tertiary Email Source":                get(2, "email_source"),
        "Tertiary Email Status":                get(2, "email_status"),
        "Tertiary Email Verification Source":   get(2, "verification_source"),
    }


def flatten_match_enriched(match):
    """Flatten a bulk_match result into enriched-column dict."""
    def safe(obj, *keys):
        for k in keys:
            if not isinstance(obj, dict): return ""
            obj = obj.get(k)
            if obj is None: return ""
        return obj if obj is not None else ""

    org    = match.get("organization") or {}
    acct   = match.get("account")      or {}
    intent = match.get("intent_signals") or []

    keywords     = ", ".join(org.get("keywords") or [])
    technologies = ", ".join([t.get("name", "") for t in (org.get("technologies") or [])]) \
                   if isinstance(org.get("technologies"), list) else ""

    row = {
        "First Name":         safe(match, "first_name"),
        "Last Name":          safe(match, "last_name"),
        "Title":              safe(match, "title"),
        "Seniority":          safe(match, "seniority"),
        "Departments":        ", ".join(match.get("departments")    or []),
        "Sub Departments":    ", ".join(match.get("subdepartments") or []),
        "Person Linkedin Url": safe(match, "linkedin_url"),
        "Facebook Url":        safe(match, "facebook_url"),
        "Twitter Url":         safe(match, "twitter_url"),
        "photo_url":           safe(match, "photo_url"),
        "City":                safe(match, "city"),
        "State":               safe(match, "state"),
        "Country":             safe(match, "country"),
        "Apollo Contact Id":   safe(match, "id"),
        "Company Name":        safe(org,   "name"),
        "Company Name for Emails": safe(acct, "name"),
        "# Employees":         safe(org,   "estimated_num_employees"),
        "Industry":            safe(org,   "industry"),
        "Keywords":            keywords,
        "Website":             safe(org,   "website_url"),
        "Company Linkedin Url": safe(org,  "linkedin_url"),
        "Company Address":     safe(org,   "raw_address"),
        "Company City":        safe(org,   "city"),
        "Company State":       safe(org,   "state"),
        "Company Country":     safe(org,   "country"),
        "Company Phone":       safe(org,   "phone") or safe(acct, "phone"),
        "Technologies":        technologies,
        "Annual Revenue":      safe(org,   "annual_revenue"),
        "Total Funding":       safe(org,   "total_funding"),
        "Latest Funding":      safe(org,   "latest_funding_stage"),
        "Latest Funding Amount": safe(org, "latest_funding_amount"),
        "Last Raised At":      safe(org,   "last_funding_date"),
        "Number of Retail Locations": safe(org, "retail_location_count"),
        "Apollo Account Id":   safe(acct,  "id"),
        "_org_id":             safe(org,   "id"),
        "Primary Intent Topic":   intent[0].get("topic", "") if len(intent) > 0 else "",
        "Primary Intent Score":   intent[0].get("score", "") if len(intent) > 0 else "",
        "Secondary Intent Topic": intent[1].get("topic", "") if len(intent) > 1 else "",
        "Secondary Intent Score": intent[1].get("score", "") if len(intent) > 1 else "",
        # CRM/activity — always blank at scrape time
        "Contact Owner": "", "Do Not Call": "", "Stage": "", "Lists": "",
        "Last Contacted": "", "Account Owner": "",
        "Subsidiary of": "", "Subsidiary of (Organization ID)": "",
        "Email Sent": "", "Email Open": "", "Email Bounced": "",
        "Replied": "", "Demoed": "", "Qualify Contact": "",
    }
    row.update(parse_emails(match))
    row.update(parse_phones(match))
    return row


def build_crm_payload_full(row):
    """
    Build a complete contacts/bulk_create payload using ALL available API fields.
    Sends multiple emails via contact_emails[] and multiple phones via phone_numbers[].
    Only includes non-empty fields in the payload.
    """
    def v(col):
        val = row.get(col, "")
        s = str(val).strip() if val is not None else ""
        return s if s not in ("", "nan", "None") else ""

    payload = {}

    # ── Core identity ──────────────────────────────────────────────────
    if v("First Name"):         payload["first_name"]        = v("First Name")
    if v("Last Name"):          payload["last_name"]         = v("Last Name")
    if v("Title"):              payload["title"]             = v("Title")
    if v("Title"):              payload["primary_title"]     = v("Title")
    if v("Email"):              payload["email"]             = v("Email")
    org_name = v("Company Name") or v("organization_name")
    if org_name:                payload["organization_name"] = org_name

    # ── Location ──────────────────────────────────────────────────────
    parts = [p for p in [v("City"), v("State"), v("Country")] if p]
    if parts: payload["present_raw_address"] = ", ".join(parts)

    # ── Best single phone (for backward compat) ───────────────────────
    best_phone = (v("Work Direct Phone") or v("Mobile Phone") or
                  v("Corporate Phone")   or v("Home Phone")   or v("Other Phone"))
    if best_phone: payload["phone"] = best_phone

    # ── Social links ──────────────────────────────────────────────────
    if v("Person Linkedin Url"): payload["linkedin_url"]  = v("Person Linkedin Url")
    if v("Facebook Url"):        payload["facebook_url"]  = v("Facebook Url")
    if v("Twitter Url"):         payload["twitter_url"]   = v("Twitter Url")
    if v("photo_url"):           payload["photo_url"]     = v("photo_url")

    # ── Apollo account/org linkage ────────────────────────────────────
    if v("Apollo Account Id"):  payload["account_id"]      = v("Apollo Account Id")
    org_id = v("_org_id") or v("apollo_org_id")
    if org_id:                  payload["organization_id"]  = org_id

    # ── Multiple emails array (primary, secondary, tertiary) ──────────
    contact_emails = []
    for pos, col in enumerate(["Email", "Secondary Email", "Tertiary Email"]):
        email = v(col)
        if email:
            contact_emails.append({"email": email, "position": pos})
    if contact_emails:
        payload["contact_emails"] = contact_emails

    # ── Multiple phones array ─────────────────────────────────────────
    phone_numbers = []
    phone_cols = [
        "Work Direct Phone", "Mobile Phone",
        "Corporate Phone", "Home Phone", "Other Phone",
    ]
    for pos, col in enumerate(phone_cols):
        ph = v(col)
        if ph:
            phone_numbers.append({"raw_number": ph, "position": pos})
    if phone_numbers:
        payload["phone_numbers"] = phone_numbers

    return payload


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  STEP 1 — PEOPLE SEARCH                                             ║
# ╚══════════════════════════════════════════════════════════════════════╝

def load_companies():
    path = os.path.join(INPUTS_DIR, COMPANIES_INPUT_FILE)
    if os.path.exists(path):
        df = pd.read_csv(path)
        df.columns = [c.strip().lower() for c in df.columns]
        if "domain" not in df.columns:
            raise ValueError(f"'domain' column not found in {path}")
        name_col = "name" if "name" in df.columns else "domain"
        print(f"  Loaded {len(df)} companies from '{path}'")
        return [{"name": row[name_col], "domain": str(row["domain"]).strip()} for _, row in df.iterrows()]
    else:
        print(f"  ⚠ '{path}' not found — using hardcoded fallback list.")
        return COMPANIES_FALLBACK


def run_step1(db_df):
    """Search for people by company domain and add NEW contacts to DB."""
    print("\n" + "═" * 68)
    print("  STEP 1 — People Search  │  POST /mixed_people/api_search")
    print("═" * 68)

    companies     = load_companies()
    existing_pids = set(db_df["apollo_person_id"].astype(str).str.strip().tolist())
    new_rows      = []

    for idx, company in enumerate(companies, 1):
        name   = company["name"]
        domain = company["domain"]
        print(f"\n  [{idx:>2}/{len(companies)}] {name}  ({domain})")

        page = 1
        company_people = []

        while page <= MAX_PAGES:
            try:
                data       = api_people_search(domain, page)
                people     = data.get("people", [])
                pagination = data.get("pagination", {})
                total_pages = pagination.get("total_pages", 1)
                total       = pagination.get("total_entries", "?")

                print(f"    Page {page:>3}/{min(total_pages, MAX_PAGES)}  —  "
                      f"{len(people)} results  (total ≈ {total})")
                company_people.extend(people)

                if page >= total_pages:
                    break
                page += 1
                time.sleep(REQUEST_DELAY)

            except requests.exceptions.HTTPError as e:
                print(f"    ✗ HTTP Error: {e}")
                break
            except Exception as e:
                print(f"    ✗ Error: {e}")
                break

        new_for_company = 0
        for p in company_people:
            pid = str(p.get("id", "")).strip()
            if pid and pid not in existing_pids:
                row = flatten_person_step1(p, name, domain)
                new_rows.append(row)
                existing_pids.add(pid)
                new_for_company += 1

        dup = len(company_people) - new_for_company
        print(f"    ✓ {new_for_company} NEW  |  {dup} already in DB (skipped)")
        time.sleep(REQUEST_DELAY)

    if new_rows:
        new_df = pd.DataFrame(new_rows)
        # Ensure all columns present
        for col in DB_ALL_COLUMNS:
            if col not in new_df.columns:
                new_df[col] = ""
        db_df = pd.concat([db_df, new_df[DB_ALL_COLUMNS]], ignore_index=True)
        print(f"\n  ✅ Step 1 complete — {len(new_rows)} new contacts added to DB")
    else:
        print("\n  ℹ  Step 1 complete — no new contacts found")

    save_db(db_df, run_log_entry={
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "step": "1 - People Search",
        "action": "search_people",
        "count": len(new_rows),
        "notes": f"{len(companies)} companies searched",
    })
    return db_df


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  STEP 2 — BULK ENRICHMENT (credit-safe)                             ║
# ╚══════════════════════════════════════════════════════════════════════╝

def run_step2(db_df):
    """Enrich only contacts where is_enriched != 'True'."""
    print("\n" + "═" * 68)
    print("  STEP 2 — Bulk Enrichment  │  POST /people/bulk_match")
    print("═" * 68)

    to_enrich = db_df[db_df["is_enriched"].astype(str) != "True"].copy()
    already   = (db_df["is_enriched"].astype(str) == "True").sum()

    print(f"  To enrich   : {len(to_enrich)}")
    print(f"  Already done: {already}  (skipping — credits protected 💰)")

    if to_enrich.empty:
        print("  ℹ  Nothing to enrich.")
        return db_df

    person_ids   = to_enrich["apollo_person_id"].astype(str).str.strip().tolist()
    batches      = [person_ids[i:i+ENRICH_BATCH_SIZE]
                    for i in range(0, len(person_ids), ENRICH_BATCH_SIZE)]
    total_credits = 0
    enrich_map    = {}  # pid → enriched dict

    print(f"\n  Processing {len(person_ids)} contacts in {len(batches)} batch(es)...\n")

    for b_idx, batch_ids in enumerate(batches, 1):
        print(f"    Batch {b_idx:>3}/{len(batches)}  ({len(batch_ids)} IDs)...", end=" ", flush=True)
        try:
            matches, raw_resp = api_bulk_match(batch_ids)
        except requests.exceptions.HTTPError as e:
            print(f"✗ HTTP Error: {e}")
            continue
        except Exception as e:
            print(f"✗ Error: {e}")
            continue

        credits  = raw_resp.get("credits_consumed", 0)
        enriched = raw_resp.get("unique_enriched_records", len(matches))
        missing  = raw_resp.get("missing_records", 0)
        total_credits += credits
        print(f"✓  matched={enriched}  missing={missing}  credits={credits}")

        for m in matches:
            flat = flatten_match_enriched(m)
            pid  = flat.get("Apollo Contact Id", "").strip()
            if pid:
                enrich_map[pid] = flat

        if b_idx < len(batches):
            time.sleep(REQUEST_DELAY)

    # ── Update DB rows with enriched data ────────────────────────────
    now_str = datetime.now().isoformat(timespec="seconds")
    updated = 0

    for i, row in db_df.iterrows():
        pid = str(row.get("apollo_person_id", "")).strip()
        if pid in enrich_map:
            enriched_data = enrich_map[pid]
            for col, val in enriched_data.items():
                db_df.at[i, col] = val
            db_df.at[i, "is_enriched"]  = "True"
            db_df.at[i, "enriched_at"]  = now_str
            updated += 1

    print(f"\n  ✅ Step 2 complete — {updated} contacts enriched")
    print(f"     Total credits consumed: {total_credits}")

    save_db(db_df, run_log_entry={
        "timestamp": now_str,
        "step": "2 - Enrichment",
        "action": "bulk_match",
        "count": updated,
        "notes": f"credits={total_credits}  missing={len(to_enrich)-updated}",
    })
    return db_df


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  STEP 3 — SAVE TO APOLLO CRM (skip already-saved)                  ║
# ╚══════════════════════════════════════════════════════════════════════╝

def run_step3(db_df):
    """Save enriched-but-not-yet-saved contacts to Apollo CRM."""
    print("\n" + "═" * 68)
    print("  STEP 3 — Save to Apollo CRM  │  POST /contacts/bulk_create")
    print("═" * 68)

    to_save  = db_df[
        (db_df["is_enriched"].astype(str) == "True") &
        (db_df["is_saved_to_apollo"].astype(str) != "True")
    ].copy()
    already  = (db_df["is_saved_to_apollo"].astype(str) == "True").sum()

    print(f"  To save     : {len(to_save)}")
    print(f"  Already done: {already}  (skipping — no duplicates 🛡️)")

    if to_save.empty:
        print("  ℹ  Nothing to save.")
        return db_df

    contacts_list = [build_crm_payload_full(row) for _, row in to_save.iterrows()]
    # Map index position → db_df index for updating after API call
    pos_to_idx    = {pos: idx for pos, (idx, _) in enumerate(to_save.iterrows())}

    batches       = [contacts_list[i:i+CRM_BATCH_SIZE]
                     for i in range(0, len(contacts_list), CRM_BATCH_SIZE)]
    total_created  = 0
    total_existing = 0
    now_str        = datetime.now().isoformat(timespec="seconds")
    processed_pos  = 0

    print(f"\n  Saving {len(contacts_list)} contacts in {len(batches)} batch(es)...\n")

    for b_idx, batch in enumerate(batches, 1):
        print(f"    CRM Batch {b_idx:>3}/{len(batches)}  ({len(batch)} contacts)...", end=" ", flush=True)
        try:
            resp = api_bulk_create(batch)
        except requests.exceptions.HTTPError as e:
            print(f"✗ HTTP Error: {e}")
            processed_pos += len(batch)
            continue
        except Exception as e:
            print(f"✗ Error: {e}")
            processed_pos += len(batch)
            continue

        created_contacts  = resp.get("created_contacts",  [])
        existing_contacts = resp.get("existing_contacts", [])
        total_created  += len(created_contacts)
        total_existing += len(existing_contacts)
        print(f"✓  created={len(created_contacts)}  existing={len(existing_contacts)}")

        # Mark created contacts as saved + store their CRM ID
        for c in created_contacts:
            crm_id    = c.get("id", "")
            crm_email = c.get("email", "").lower().strip()
            # Find matching row in to_save by email
            for pos in range(processed_pos, processed_pos + len(batch)):
                db_idx = pos_to_idx.get(pos)
                if db_idx is None:
                    continue
                row_email = str(db_df.at[db_idx, "Email"]).lower().strip()
                if row_email == crm_email or not crm_email:
                    db_df.at[db_idx, "is_saved_to_apollo"]   = "True"
                    db_df.at[db_idx, "saved_at"]             = now_str
                    db_df.at[db_idx, "apollo_crm_contact_id"] = crm_id
                    break

        # Mark existing contacts as saved (they were already there)
        for c in existing_contacts:
            crm_id    = c.get("id", "")
            crm_email = c.get("email", "").lower().strip()
            for pos in range(processed_pos, processed_pos + len(batch)):
                db_idx = pos_to_idx.get(pos)
                if db_idx is None:
                    continue
                row_email = str(db_df.at[db_idx, "Email"]).lower().strip()
                if row_email == crm_email or not crm_email:
                    db_df.at[db_idx, "is_saved_to_apollo"]   = "True"
                    db_df.at[db_idx, "saved_at"]             = now_str
                    db_df.at[db_idx, "apollo_crm_contact_id"] = crm_id
                    break

        processed_pos += len(batch)
        if b_idx < len(batches):
            time.sleep(REQUEST_DELAY)

    print(f"\n  ✅ Step 3 complete")
    print(f"     Created : {total_created}")
    print(f"     Existing: {total_existing}  (already in CRM — marked as saved)")

    save_db(db_df, run_log_entry={
        "timestamp": now_str,
        "step": "3 - CRM Save",
        "action": "bulk_create",
        "count": total_created + total_existing,
        "notes": f"created={total_created}  existing={total_existing}",
    })
    return db_df


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  MAIN ORCHESTRATOR                                                  ║
# ╚══════════════════════════════════════════════════════════════════════╝

def main():
    print("\n" + "╔" + "═" * 66 + "╗")
    print("║" + "  APOLLO MASTER UTILITY — Credit-Safe Contact Pipeline".center(66) + "║")
    print("╚" + "═" * 66 + "╝")
    print(f"  Steps    : {STEPS_TO_RUN}")
    print(f"  DB file  : {DB_FILE}")
    print(f"  Snapshot : {SNAPSHOT_FILE}")
    print("─" * 68)

    db_df = load_db()

    steps = STEPS_TO_RUN.upper().strip()

    if steps in ("1", "ALL"):
        db_df = run_step1(db_df)

    if steps in ("2", "2+3", "ALL"):
        db_df = run_step2(db_df)

    if steps in ("3", "2+3", "ALL") or (steps == "2+3" and AUTO_SAVE_ENRICHED):
        db_df = run_step3(db_df)

    # Final snapshot for this run
    save_snapshot(db_df)

    print("\n" + "═" * 68)
    print("  🏁  Run complete!")
    print(f"  Master DB : {DB_FILE}")
    print(f"  Snapshot  : {SNAPSHOT_FILE}")
    print("═" * 68 + "\n")


if __name__ == "__main__":
    main()
'''

output_path = "/root/apollo_master_utility.py"
with open(output_path, "w") as f:
    f.write(script)

print(f"Script written: {len(script)} characters")
print("Lines:", script.count('\n'))