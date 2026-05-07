#!/usr/bin/env python3
"""
═══════════════════════════════════════════════════════════════
  Apollo Contact Scraper — STEP 2: Bulk People Enrichment
                         + STEP 3: Save to Apollo CRM
  Uses:
    POST /api/v1/people/bulk_match       (enrich)
    POST /api/v1/contacts/bulk_create    (save to CRM)

  INPUT : Filtered Excel/CSV from Step 1  (output/ folder)
  OUTPUT: output/apollo_enriched_DDMMMYYYY_HHMMSS.xlsx
          4 sheets:
            Enriched   — ALL input rows (matched + unmatched)
            Matched    — Rows where Apollo returned an email
            Unmatched  — Rows Apollo could not find
            CRM Save   — Matched rows formatted for contacts/bulk_create
                         (review this sheet before saving to CRM)

  LIMIT : bulk_match      → 10  contacts per API call
          contacts/bulk_create → 100 contacts per API call
═══════════════════════════════════════════════════════════════
"""

import requests
import pandas as pd
import time
import os
from datetime import datetime

# ─── resolve paths relative to this script's location ───────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUTS_DIR   = os.path.join(SCRIPT_DIR, "..", "inputs")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ╔══════════════════════════════════════════════════════════════╗
# ║                     CONFIGURATION                           ║
# ╚══════════════════════════════════════════════════════════════╝

API_KEY = "vWTBtYd1P9IpMLV-wghAzw"        # Your Apollo master API key

TIMESTAMP = datetime.now().strftime("%d%b%Y_%H%M%S").lower()   # e.g. 27apr2026_124055

# Input file — filename inside output/ folder (your filtered Step 1 output)
# BULK_ENRICH_INPUT_FILE = "apollo_people_search_20260423_144441.xlsx"
BULK_ENRICH_INPUT_FILE = "apollo_people_search_sample.xlsx"

# Output file — timestamped so it is never overwritten
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"apollo_enriched_{TIMESTAMP}.xlsx")

MASTER_DB = os.path.join(OUTPUT_DIR, "master_lead_database.xlsx")

# ── Enrichment settings ──────────────────────────────────────────────────────
BATCH_SIZE             = 10     # Apollo hard limit: max 10 IDs per bulk_match call
REQUEST_DELAY          = 1.2    # Seconds between API calls (avoid rate-limiting)
REVEAL_PERSONAL_EMAILS = True   # True = uses extra credits (requires supported plan)
REVEAL_PHONE_NUMBER    = False  # True = uses extra credits

# 
SINGLE_SHEET_OUTPUT = False   # 🔁 Toggle this

# ── CRM save settings ────────────────────────────────────────────────────────
CREATE_IN_CRM   = True   # Set True to auto-call contacts/bulk_create after enrichment
                          # Set False to only write the "CRM Save" sheet for manual review
CRM_BATCH_SIZE  = 100     # contacts/bulk_create limit: max 100 per request
RUN_DEDUPE      = True    # Prevents duplicate contacts; existing ones returned separately
LABEL_NAMES     = []      # e.g. ["Decision Maker", "Apollo Scraped"] — applied to all created contacts


# ╔══════════════════════════════════════════════════════════════╗
# ║              OUTPUT COLUMN ORDER (Apollo format)            ║
# ╚══════════════════════════════════════════════════════════════╝

DESIRED_COLUMNS = [
    "apollo_person_id",   # 🔥 ADD THIS
    "First Name",
    "Last Name",
    "Title",
    "Company Name",
    "Company Name for Emails",
    "Email",
    "Email Status",
    "Primary Email Source",
    "Primary Email Verification Source",
    "Email Confidence",
    "Primary Email Catch-all Status",
    "Primary Email Last Verified At",
    "Seniority",
    "Departments",
    "Sub Departments",
    "Contact Owner",
    "Work Direct Phone",
    "Home Phone",
    "Mobile Phone",
    "Corporate Phone",
    "Other Phone",
    "Do Not Call",
    "Stage",
    "Lists",
    "Last Contacted",
    "Account Owner",
    "# Employees",
    "Industry",
    "Keywords",
    "Person Linkedin Url",
    "Website",
    "Company Linkedin Url",
    "Facebook Url",
    "Twitter Url",
    "City",
    "State",
    "Country",
    "Company Address",
    "Company City",
    "Company State",
    "Company Country",
    "Company Phone",
    "Technologies",
    "Annual Revenue",
    "Total Funding",
    "Latest Funding",
    "Latest Funding Amount",
    "Last Raised At",
    "Subsidiary of",
    "Subsidiary of (Organization ID)",
    "Email Sent",
    "Email Open",
    "Email Bounced",
    "Replied",
    "Demoed",
    "Number of Retail Locations",
    "Apollo Contact Id",
    "Apollo Account Id",
    "Secondary Email",
    "Secondary Email Source",
    "Secondary Email Status",
    "Secondary Email Verification Source",
    "Tertiary Email",
    "Tertiary Email Source",
    "Tertiary Email Status",
    "Tertiary Email Verification Source",
    "Primary Intent Topic",
    "Primary Intent Score",
    "Secondary Intent Topic",
    "Secondary Intent Score",
    "Qualify Contact",
]

# Columns for the "CRM Save" sheet — named exactly as contacts/bulk_create API fields
CRM_SAVE_COLUMNS = [
    "first_name",
    "last_name",
    "title",
    "organization_name",
    "email",
    "phone",
    "linkedin_url",
    "twitter_url",
    "facebook_url",
    "photo_url",
    "account_id",       # Apollo Account Id — links contact to existing CRM account
    "organization_id",  # Apollo Org Id — links to Apollo org record
]


# ╔══════════════════════════════════════════════════════════════╗
# ║                     HELPER FUNCTIONS                        ║
# ╚══════════════════════════════════════════════════════════════╝

BASE_URL = "https://api.apollo.io/api/v1"


def load_input_file():
    path = os.path.join(OUTPUT_DIR, BULK_ENRICH_INPUT_FILE)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Input file not found: '{path}'\n"
            f"Set BULK_ENRICH_INPUT_FILE to a filename inside the output/ folder."
        )
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(path)
    elif ext == ".csv":
        df = pd.read_csv(path)
    else:
        raise ValueError(f"Unsupported file type: {ext}  (use .xlsx or .csv)")

    df.columns = [c.strip().lower() for c in df.columns]
    if "apollo_person_id" not in df.columns:
        raise ValueError("Column 'apollo_person_id' not found. Use Step 1 output.")

    before = len(df)
    df = df[df["apollo_person_id"].notna() & (df["apollo_person_id"].astype(str).str.strip() != "")]
    dropped = before - len(df)
    if dropped:
        print(f"  ⚠  Dropped {dropped} rows with empty apollo_person_id")

    print(f"  Loaded {len(df)} contacts from 'output/{BULK_ENRICH_INPUT_FILE}'")
    return df.reset_index(drop=True)

# API CALL - BULK MATCH + ENRICH
def bulk_match_batch(person_ids):
    """POST one batch (max 10 IDs) to /people/bulk_match."""
    response = requests.post(
        f"{BASE_URL}/people/bulk_match",
        params={
            "reveal_personal_emails": str(REVEAL_PERSONAL_EMAILS).lower(),
            "reveal_phone_number":    str(REVEAL_PHONE_NUMBER).lower(),
        },
        json={"details": [{"id": pid} for pid in person_ids]},
        headers={
            "Cache-Control": "no-cache",
            "Content-Type":  "application/json",
            "Accept":        "application/json",
            "x-api-key":     API_KEY,
        },
        timeout=60,
    )
    if response.status_code == 429:
        wait = int(response.headers.get("Retry-After", 30))
        print(f"\n    ⚠  Rate limited — waiting {wait}s...")
        time.sleep(wait)
        return bulk_match_batch(person_ids)
    response.raise_for_status()
    data = response.json()
    return data.get("matches", []), data

# API CALL - BULK CREATE CONTACTS IN CRM
def bulk_create_crm_batch(contacts_payload):
    """POST one batch (max 100 contacts) to /contacts/bulk_create."""
    body = {
        "contacts":   contacts_payload,
        "run_dedupe": RUN_DEDUPE,
    }
    if LABEL_NAMES:
        body["append_label_names"] = LABEL_NAMES

    response = requests.post(
        f"{BASE_URL}/contacts/bulk_create",
        json=body,
        headers={
            "Cache-Control": "no-cache",
            "Content-Type":  "application/json",
            "Accept":        "application/json",
            "x-api-key":     API_KEY,
        },
        timeout=60,
    )
    if response.status_code == 429:
        wait = int(response.headers.get("Retry-After", 30))
        print(f"\n    ⚠  Rate limited — waiting {wait}s...")
        time.sleep(wait)
        return bulk_create_crm_batch(contacts_payload)
    response.raise_for_status()
    return response.json()

# BUILD SKIP LIST
def get_enrichment_candidates(input_df):
    """
    Decide which records should go to enrichment.
    Rules:
      1. If Apollo Person ID is already enriched in MASTER DB (has email) → skip forever
      2. If CRM contact already enriched (is_crm_contact=True and is_enriched=True) → skip
      3. Otherwise → enrich
    Returns:
        - to_enrich : List of Apollo Person IDs to enrich in this run
        - skipped_df : DataFrame of skipped rows (for logging/debugging)
    """

    # Load master DB
    if os.path.exists(MASTER_DB):
        master_df = pd.read_excel(MASTER_DB, dtype=str)
        master_df["apollo_person_id"] = master_df["apollo_person_id"].astype(str).str.strip()

        enriched_ids = set(
            master_df[
                master_df["is_enriched"].astype(str).str.lower() == "true"
            ]["apollo_person_id"]
        )
    else:
        enriched_ids = set()

    input_df["apollo_person_id"] = input_df["apollo_person_id"].astype(str).str.strip()

    to_enrich = []
    skipped_rows = []

    for _, row in input_df.iterrows():
        pid = row["apollo_person_id"]

        is_enriched = str(row.get("is_enriched", "")).lower() == "true"
        is_crm = str(row.get("is_crm_contact", "")).lower() == "true"

        # ✅ Rule 1: Already enriched in MASTER → skip forever
        if pid in enriched_ids:
            skipped_rows.append((pid, "MASTER_ALREADY_ENRICHED"))
            continue

        # ✅ Rule 2: CRM contact already enriched → skip
        if is_crm and is_enriched:
            skipped_rows.append((pid, "CRM_ALREADY_ENRICHED"))
            continue

        # ✅ Rule 3: Needs enrichment
        to_enrich.append(pid)

    print(f"\n  🧠 Smart Filter Summary:")
    print(f"     Total input          : {len(input_df)}")
    print(f"     To enrich (NEW)      : {len(to_enrich)}")
    print(f"     Skipped              : {len(skipped_rows)}")

    return to_enrich

def parse_phones(match):
    phones  = match.get("phone_numbers") or []
    buckets = {"Work Direct Phone": "", "Home Phone": "", "Mobile Phone": "", "Corporate Phone": "", "Other Phone": ""}
    type_map = {"work_direct": "Work Direct Phone", "home": "Home Phone",
                "mobile": "Mobile Phone", "corporate": "Corporate Phone", "other": "Other Phone"}
    for p in phones:
        raw    = p.get("sanitized_number") or p.get("raw_number", "")
        ptype  = (p.get("type") or "other").lower().replace(" ", "_")
        bucket = type_map.get(ptype, "Other Phone")
        if raw and not buckets[bucket]:
            buckets[bucket] = raw
    return buckets


def parse_emails(match):
    contact_emails = sorted(match.get("contact_emails") or [], key=lambda x: x.get("position", 99))
    def get(idx, field):
        if idx < len(contact_emails):
            v = contact_emails[idx].get(field)
            return v if v is not None else ""
        return ""
    return {
        "Email":                               get(0, "email") or match.get("email", ""),
        "Email Status":                        get(0, "email_status") or match.get("email_status", ""),
        "Primary Email Source":                get(0, "email_source"),
        "Primary Email Verification Source":   get(0, "verification_source"),
        "Email Confidence":                    match.get("extrapolated_email_confidence", ""),
        "Primary Email Catch-all Status":      get(0, "catch_all_status"),
        "Primary Email Last Verified At":      get(0, "last_verified_at"),
        "Secondary Email":                     get(1, "email"),
        "Secondary Email Source":              get(1, "email_source"),
        "Secondary Email Status":              get(1, "email_status"),
        "Secondary Email Verification Source": get(1, "verification_source"),
        "Tertiary Email":                      get(2, "email"),
        "Tertiary Email Source":               get(2, "email_source"),
        "Tertiary Email Status":               get(2, "email_status"),
        "Tertiary Email Verification Source":  get(2, "verification_source"),
    }


def flatten_match(match):
    def safe(obj, *keys):
        for k in keys:
            if not isinstance(obj, dict): return ""
            obj = obj.get(k)
            if obj is None: return ""
        return obj if obj is not None else ""

    org  = match.get("organization") or {}
    acct = match.get("account") or {}

    keywords     = ", ".join(org.get("keywords") or [])
    technologies = ", ".join([t.get("name", "") for t in (org.get("technologies") or [])]) \
                   if isinstance(org.get("technologies"), list) else ""
    
    intent = match.get("intent_signals") or []

    pid = str(match.get("id", "")).strip()  # 🔥 PRIMARY KEY

    row = {
        # 🔥 PRIMARY KEY (used across pipeline)
        "apollo_person_id": pid,

        # Keep this for backward compatibility
        "Apollo Contact Id": pid,

        "First Name":                       safe(match, "first_name"),
        "Last Name":                        safe(match, "last_name"),
        "Title":                            safe(match, "title"),
        "Seniority":                        safe(match, "seniority"),
        "Departments":                      ", ".join(match.get("departments") or []),
        "Sub Departments":                  ", ".join(match.get("subdepartments") or []),
        "Person Linkedin Url":              safe(match, "linkedin_url"),
        "Facebook Url":                     safe(match, "facebook_url"),
        "Twitter Url":                      safe(match, "twitter_url"),
        "City":                             safe(match, "city"),
        "State":                            safe(match, "state"),
        "Country":                          safe(match, "country"),
        "Apollo Contact Id":                safe(match, "id"),

        "Company Name":                     safe(org, "name"),
        "Company Name for Emails":          safe(acct, "name"),
        "# Employees":                      safe(org, "estimated_num_employees"),
        "Industry":                         safe(org, "industry"),
        "Keywords":                         keywords,
        "Website":                          safe(org, "website_url"),
        "Company Linkedin Url":             safe(org, "linkedin_url"),
        "Company Address":                  safe(org, "raw_address"),
        "Company City":                     safe(org, "city"),
        "Company State":                    safe(org, "state"),
        "Company Country":                  safe(org, "country"),
        "Company Phone":                    safe(org, "phone") or safe(acct, "phone"),
        "Technologies":                     technologies,
        "Annual Revenue":                   safe(org, "annual_revenue"),
        "Total Funding":                    safe(org, "total_funding"),
        "Latest Funding":                   safe(org, "latest_funding_stage"),
        "Latest Funding Amount":            safe(org, "latest_funding_amount"),
        "Last Raised At":                   safe(org, "last_funding_date"),
        "Number of Retail Locations":       safe(org, "retail_location_count"),
        "Apollo Account Id":                safe(acct, "id"),

        "Primary Intent Topic":             intent[0].get("topic", "") if len(intent) > 0 else "",
        "Primary Intent Score":             intent[0].get("score", "") if len(intent) > 0 else "",
        "Secondary Intent Topic":           intent[1].get("topic", "") if len(intent) > 1 else "",
        "Secondary Intent Score":           intent[1].get("score", "") if len(intent) > 1 else "",
        # CRM / activity — always empty
        "Contact Owner": "", "Do Not Call": "", "Stage": "", "Lists": "",
        "Last Contacted": "", "Account Owner": "", "Subsidiary of": "",
        "Subsidiary of (Organization ID)": "", "Email Sent": "", "Email Open": "",
        "Email Bounced": "", "Replied": "", "Demoed": "", "Qualify Contact": "",

        # internal field used to build CRM Save sheet
        "_org_id": safe(org, "id"),

        # 🔥 YOUR FINAL RULE
        "is_enriched": "true",
    }

    # Optional: keep if still using somewhere
    row["is_matched"] = "true"

    row.update(parse_emails(match))
    row.update(parse_phones(match))

    return row


def build_crm_payload(row):
    """
    Map one enriched row → contacts/bulk_create contact object.
    Only non-empty fields are included so the API ignores truly missing data.
    """
    # Best available phone: prefer Work Direct, fallback through others
    phone = (row.get("Work Direct Phone") or row.get("Mobile Phone") or
             row.get("Corporate Phone") or row.get("Home Phone") or
             row.get("Other Phone") or "")

    mapping = {
        "first_name":       row.get("First Name", ""),
        "last_name":        row.get("Last Name", ""),
        "title":            row.get("Title", ""),
        "organization_name":row.get("Company Name", ""),
        "email":            row.get("Email", ""),
        "phone":            phone,
        "linkedin_url":     row.get("Person Linkedin Url", ""),
        "twitter_url":      row.get("Twitter Url", ""),
        "facebook_url":     row.get("Facebook Url", ""),
        "photo_url":        row.get("photo_url", ""),  # available in flatten_match
        "account_id":       row.get("Apollo Account Id", ""),
        "organization_id":  row.get("_org_id", ""),
    }
    # Strip empty values — send only populated fields
    return {k: v for k, v in mapping.items() if str(v).strip()}


# ❌REMOVE - NOT USED❌
# def empty_enriched_row(apollo_person_id=""):
#     row = {col: "" for col in DESIRED_COLUMNS}
#     row["Apollo Contact Id"] = apollo_person_id
#     row["is_matched"] = "false"
#     return row

def empty_enriched_row(apollo_person_id=""):
    pid = str(apollo_person_id).strip()

    row = {col: "" for col in DESIRED_COLUMNS}

    # 🔥 PRIMARY KEY consistency
    row["apollo_person_id"] = pid
    row["Apollo Contact Id"] = pid

    # 🔥 Since API was called but no data returned
    row["is_enriched"] = "true"

    # Optional (if still referenced anywhere)
    row["is_matched"] = "false"

    return row

def enforce_column_order(df):
    for col in DESIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[DESIRED_COLUMNS]



# BUILD OUTPUT EXCEL WITH 4 SHEETS: Enriched (all), Matched, Unmatched, CRM Save
def save_output(enrich_map, input_df):
    """
    Build and write 4 sheets:

      Enriched   — ALL input rows (same order as input)
      Matched    — Rows where enrichment API returned data
      Unmatched  — Rows where API returned nothing
      CRM Save   — Only matched rows (ready for bulk_create)

    NOTE:
    - "Matched" = API returned result (NOT based on email)
    - "is_enriched" = API was called (your final rule)
    """

    enriched_rows = []
    unmatched_pids = []

    # ── Step 1: Build full dataset (ALL rows) ─────────────────────
    for _, in_row in input_df.iterrows():
        pid = str(in_row.get("apollo_person_id", "")).strip()

        if pid in enrich_map:
            # ✅ API returned result → matched
            row = enrich_map[pid]
            row["is_enriched"] = "true"   # 🔥 mark as enriched (API hit)
            enriched_rows.append(row)
        else:
            # ❌ No result from API → unmatched
            row = empty_enriched_row(apollo_person_id=pid)
            row["is_enriched"] = "true"   # 🔥 still enriched (API was called)
            enriched_rows.append(row)
            unmatched_pids.append(pid)

    # Convert to DataFrame
    all_df = pd.DataFrame(enriched_rows)
    all_df = enforce_column_order(all_df)

    # ── Step 2: Split Matched vs Unmatched ────────────────────────
    # 👉 Based on enrich_map presence (NOT email, NOT is_matched)

    matched_df = all_df[
        all_df["apollo_person_id"].isin(enrich_map.keys())
    ].reset_index(drop=True)

    unmatched_df = all_df[
        ~all_df["apollo_person_id"].isin(enrich_map.keys())
    ].reset_index(drop=True)

    # ── Step 3: CRM Save Sheet (ONLY matched rows) ────────────────
    crm_rows = []

    for _, row in matched_df.iterrows():
        payload = build_crm_payload(row.to_dict())
        crm_rows.append(payload)

    crm_df = pd.DataFrame(crm_rows)

    # Ensure all required columns exist
    for col in CRM_SAVE_COLUMNS:
        if col not in crm_df.columns:
            crm_df[col] = ""

    crm_df = crm_df[CRM_SAVE_COLUMNS]

    # 🔥 Important: remove NaN before API usage
    crm_df = crm_df.fillna("")

    # ── Step 4: Write Excel ──────────────────────────────────────
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

        if SINGLE_SHEET_OUTPUT:
            # 🔥 Single sheet mode
            all_df.to_excel(writer, index=False, sheet_name="All Data")

        else:
            # 🔥 Multi-sheet mode (existing behavior)  
            all_df.to_excel(writer,       index=False, sheet_name="Enriched")
            matched_df.to_excel(writer,   index=False, sheet_name="Matched")
            unmatched_df.to_excel(writer, index=False, sheet_name="Unmatched")
            crm_df.to_excel(writer,       index=False, sheet_name="CRM Save")

        # Auto column width (nice touch 👌)
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 45)

    # ── Step 5: Output Name ──────────────────────────────────────
    out_name = os.path.basename(OUTPUT_FILE)

    # ── Step 6: Logging ──────────────────────────────────────────
    print(f"\n  ✅  Saved → 'output/{out_name}'")
    if SINGLE_SHEET_OUTPUT:
        print(f"       ├─ Mode: Single Sheet (All Data)")
        print(f"       └─ Enriched (all rows)    : {len(all_df)}")

    else:
        print(f"       ├─ Mode: Multi Sheets")
        print(f"       ├─ Enriched (all rows)    : {len(all_df)}")
        print(f"       ├─ Matched  (API success) : {len(matched_df)}")
        print(f"       ├─ Unmatched              : {len(unmatched_df)}")
        print(f"       └─ CRM Save (bulk_create) : {len(crm_df)}")

    if unmatched_pids:
        print(
            f"\n  Unmatched person IDs ({len(unmatched_pids)}): "
            f"{', '.join(unmatched_pids[:5])}"
            + ("..." if len(unmatched_pids) > 5 else "")
        )

    return crm_df




def run_bulk_create_crm(crm_df):
    """
    Call POST /api/v1/contacts/bulk_create for all rows in crm_df.
    Batches of 100. Prints created vs existing counts per batch.
    """
    
    # This line can't handle NaN => so Crashing ❌
    # contacts_list = crm_df.to_dict("records")


    # Fill NaN with empty strings before converting to list of dicts, so we don't lose any rows due to NaN values in the CRM payload
    crm_df = crm_df.fillna("")

    # (more robust) Convert DataFrame to list of dicts, stripping whitespace and excluding empty values to minimize payload size
    contacts_list = [
        {k: v for k, v in row.items() if str(v).strip()}
        for row in crm_df.to_dict("records")
    ]

    batches       = [contacts_list[i:i+CRM_BATCH_SIZE] for i in range(0, len(contacts_list), CRM_BATCH_SIZE)]
    total_created  = 0
    total_existing = 0

    print(f"\n  Saving {len(contacts_list)} contacts to Apollo CRM in {len(batches)} batch(es)...\n")

    for b_idx, batch in enumerate(batches, 1):
        print(f"  CRM Batch {b_idx:>3}/{len(batches)}  ({len(batch)} contacts)...", end=" ", flush=True)
        try:
            resp     = bulk_create_crm_batch(batch)
            created  = len(resp.get("created_contacts", []))
            existing = len(resp.get("existing_contacts", []))
            total_created  += created
            total_existing += existing
            print(f"✓  created={created}  existing={existing}")
        except requests.exceptions.HTTPError as e:
            print(f"\n    ✗  HTTP Error: {e}")
        except Exception as e:
            print(f"\n    ✗  Error: {e}")

        if b_idx < len(batches):
            time.sleep(REQUEST_DELAY)

    print(f"\n  ── CRM Save Complete ──────────────────────────────────")
    print(f"     Total created  : {total_created}")
    print(f"     Total existing : {total_existing}  (already in your CRM, not duplicated)")


# LAST STEP: UPDATE MASTER DB WITH NEW ENRICHED DATA & STATUS
def update_master_db(input_df, enrich_map):
    # MASTER_DB = os.path.join(OUTPUT_DIR, "master_lead_database.xlsx")

    records = []

    for _, row in input_df.iterrows():
        pid = str(row.get("apollo_person_id", "")).strip()

        if pid in enrich_map:
            enriched = enrich_map[pid]

            record = {
                "apollo_person_id": pid,
                "first_name": enriched.get("First Name", ""),
                "last_name": enriched.get("Last Name", ""),
                "full_name": f"{enriched.get('First Name','')} {enriched.get('Last Name','')}".strip(),
                "title": enriched.get("Title", ""),
                "linkedin_url": enriched.get("Person Linkedin Url", ""),
                "email": enriched.get("Email", ""),
                "email_status": enriched.get("Email Status", ""),
                "organization_name": enriched.get("Company Name", ""),
                "website_url": enriched.get("Website", ""),
                "org_linkedin_url": enriched.get("Company Linkedin Url", ""),
                "org_industry": enriched.get("Industry", ""),
                "org_employees": enriched.get("# Employees", ""),
                "apollo_org_id": enriched.get("_org_id", ""),
                "source_type": row.get("source_type", ""),
                # "is_enriched": "true" if enriched.get("Email") else "false", # Don't Give Fuck Whether Email is Found or Not 😎✌️
                "is_enriched": "true",
                "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

        else:
            record = row.to_dict()
            record["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        records.append(record)

    new_df = pd.DataFrame(records)

    # Load existing master
    if os.path.exists(MASTER_DB):
        master_df = pd.read_excel(MASTER_DB, dtype=str)

        combined = pd.concat([master_df, new_df], ignore_index=True)

        # Deduplicate
        combined.drop_duplicates(subset=["apollo_person_id"], keep="last", inplace=True)
    else:
        combined = new_df

    combined.fillna("", inplace=True)
    combined.to_excel(MASTER_DB, index=False)

    print(f"\n  🧠 Master DB Updated → {len(combined)} total records")


# ╔══════════════════════════════════════════════════════════════╗
# ║                        MAIN RUNNER                         ║
# ╚══════════════════════════════════════════════════════════════╝

def main():
    print("\n" + "═" * 62)
    print("  Apollo Bulk People Enrichment — Step 2 of 2")
    print("  Endpoint: POST /api/v1/people/bulk_match")
    print("═" * 62)
    print(f"  Input file            : output/{BULK_ENRICH_INPUT_FILE}")
    print(f"  Output file           : output/{os.path.basename(OUTPUT_FILE)}")
    print(f"  Batch size (enrich)   : {BATCH_SIZE}  (Apollo max per bulk_match call)")
    print(f"  Batch size (CRM save) : {CRM_BATCH_SIZE}  (Apollo max per bulk_create call)")
    print(f"  Reveal personal email : {REVEAL_PERSONAL_EMAILS}")
    print(f"  Reveal phone number   : {REVEAL_PHONE_NUMBER}")
    print(f"  Create in CRM now     : {CREATE_IN_CRM}")
    if LABEL_NAMES:
        print(f"  Labels to apply       : {', '.join(LABEL_NAMES)}")
    print("═" * 62 + "\n")

    input_df = load_input_file()
    total    = len(input_df)

    # Don't Take the Ids from the input file directly Now
    # person_ids_all = input_df["apollo_person_id"].astype(str).str.strip().tolist()

    # Take Only the IDs that need enrichment based on our smart filter
    person_ids_all = get_enrichment_candidates(input_df)

    # If no IDs to enrich after filtering, skip API calls and just save empty enriched rows
    if not person_ids_all:
        print("\n  ⚠ No contacts to enrich after smart filtering. Skipping API calls.")
        return

    batches        = [person_ids_all[i:i+BATCH_SIZE] for i in range(0, total, BATCH_SIZE)]
    total_batches  = len(batches)
    total_credits  = 0

    print(f"  Processing {total} contacts in {total_batches} batch(es) of {BATCH_SIZE}...\n")

    enrich_map = {}   # apollo_person_id → flattened enriched dict

    for b_idx, batch_ids in enumerate(batches, 1):
        print(f"  Batch {b_idx:>3}/{total_batches}  ({len(batch_ids)} IDs)...", end=" ", flush=True)
        try:
            matches, raw_resp = bulk_match_batch(batch_ids)
        except requests.exceptions.HTTPError as e:
            print(f"\n    ✗  HTTP Error: {e}")
            continue
        except Exception as e:
            print(f"\n    ✗  Error: {e}")
            continue

        credits  = raw_resp.get("credits_consumed", 0)
        enriched = raw_resp.get("unique_enriched_records", len(matches))
        missing  = raw_resp.get("missing_records", 0)
        total_credits += credits
        print(f"✓  matched={enriched}  missing={missing}  credits={credits}")

        for m in matches:
            flat = flatten_match(m)
            pid  = flat.get("Apollo Contact Id", "").strip()
            if pid:
                enrich_map[pid] = flat

        if b_idx < total_batches:
            time.sleep(REQUEST_DELAY)

    if not enrich_map:
        print("\n  ⚠  No matches returned from API. Check your API key and input file.")
        return

    print("\n" + "─" * 62)
    crm_df = save_output(enrich_map, input_df)

    # Update master DB with enriched data and enrichment status (is_enriched=True if email found, else False)
    update_master_db(input_df, enrich_map)

    out_name = os.path.basename(OUTPUT_FILE)
    print()
    print(f"  Total enrichment credits consumed : {total_credits}")

    # ── Optional: auto-save to CRM ────────────────────────────────
    if CREATE_IN_CRM:
        if crm_df.empty:
            print("\n  ⚠  No matched contacts to save to CRM.")
        else:
            run_bulk_create_crm(crm_df)
    else:
        print()
        print("  ─── TO SAVE CONTACTS TO APOLLO CRM ─────────────────────")
        print(f"  1. Review the 'CRM Save' sheet in 'output/{out_name}'")
        print("  2. Remove any rows you DON'T want to create")
        print("  3. Set  CREATE_IN_CRM = True  in CONFIGURATION")
        print("  4. Re-run this script  (it will skip enrichment step?)")
        print("     — OR just set CREATE_IN_CRM = True on first run")
        print("─" * 62 + "\n")

    print("  ─── SHEETS SUMMARY ──────────────────────────────────────")
    print("  'Enriched'  = ALL input rows (matched + unmatched, same order)")
    print("  'Matched'   = Rows where Apollo returned an email")
    print("  'Unmatched' = Rows Apollo could not find")
    print("  'CRM Save'  = Ready for contacts/bulk_create API")
    print("─" * 62 + "\n")


if __name__ == "__main__":
    main()
