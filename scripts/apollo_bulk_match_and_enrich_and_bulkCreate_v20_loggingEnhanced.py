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

  CURRENT_FLOW : ✅
  Enrichment → Save Output → Update Master DB → CRM Save
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
# BULK_ENRICH_INPUT_FILE = "apollo_people_search_04may2026_161617.xlsx"
# BULK_ENRICH_INPUT_FILE = "apollo_people_search_sample.xlsx"
# BULK_ENRICH_INPUT_FILE = "apollo_people_search_06may2026_125052.xlsx"
# BULK_ENRICH_INPUT_FILE = "apollo_people_search_06may2026_130139.xlsx"
BULK_ENRICH_INPUT_FILE = "apollo_people_search_06may2026_131327.xlsx"

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
LABEL_NAMES     = ["Created_with_ApolloAPI"]      # e.g. ["Decision Maker", "Apollo Scraped"] — applied to all created contacts


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
    # "apollo_person_id"
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

FINAL_COLUMNS = [
    "apollo_person_id","first_name","last_name","full_name",
    "title","seniority","linkedin_url","email","email_status",
    "city","state","country","searched_company","searched_domain",
    "organization_name","website_url","org_linkedin_url",
    "org_industry","org_employees","apollo_org_id",
    "source_type","is_enriched","is_crm_contact",
    "last_updated","is_matched"
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



    master_skip   = sum(1 for _, r in skipped_rows if r == "MASTER_ALREADY_ENRICHED")
    crm_skip      = sum(1 for _, r in skipped_rows if r == "CRM_ALREADY_ENRICHED")
    credits_saved = len(skipped_rows)

    # Logging summary
    print(f"\n  🧠  Smart Filter Summary Before Enrichment:")
    print(f"       ├─ Total input            : {len(input_df)}")
    print(f"       ├─ To enrich (NEW)        : {len(to_enrich)}")
    print(f"       ├─ Skipped (total)        : {len(skipped_rows)}")
    print(f"       │   ├─ Master DB dupes    : {master_skip}")
    print(f"       │   └─ CRM pre-enriched   : {crm_skip}")
    print(f"       └─ Credits saved          : {credits_saved}  🎉🫀")

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

    # # ── Step 3: CRM Save Sheet (ONLY matched rows) ────────────────
    # crm_rows = []

    # for _, row in matched_df.iterrows():
    #     payload = build_crm_payload(row.to_dict())
    #     crm_rows.append(payload)

    # crm_df = pd.DataFrame(crm_rows)

    # # ── Step 3: CRM Save Sheet (ONLY matched rows) ────────────────
    # crm_rows = []
    
    # for _, row in matched_df.iterrows():
    #     payload = build_crm_payload(row.to_dict())
    #     # 🔥 CRITICAL: carry apollo_person_id so we can match CRM response back
    #     payload["apollo_person_id"] = str(row.get("apollo_person_id", "")).strip()
    #     crm_rows.append(payload)
    
    # crm_df = pd.DataFrame(crm_rows)

    # # Ensure all required columns exist
    # for col in CRM_SAVE_COLUMNS:
    #     if col not in crm_df.columns:
    #         crm_df[col] = ""

    # # crm_df = crm_df[CRM_SAVE_COLUMNS]
    # crm_df = crm_df[["apollo_person_id"] + CRM_SAVE_COLUMNS]

    # # ── Step 3: CRM Save Sheet (ONLY matched rows) ────────────────
    crm_rows = []

    for _, row in matched_df.iterrows():
        payload = build_crm_payload(row.to_dict())
        # 🔥 CRITICAL: carry apollo_person_id so we can match CRM response back
        payload["apollo_person_id"] = str(row.get("apollo_person_id", "")).strip()
        crm_rows.append(payload)

    FULL_CRM_COLUMNS = ["apollo_person_id"] + CRM_SAVE_COLUMNS

    if crm_rows:
        crm_df = pd.DataFrame(crm_rows)
        # Ensure all required columns exist
        for col in FULL_CRM_COLUMNS:
            if col not in crm_df.columns:
                crm_df[col] = ""
        crm_df = crm_df[FULL_CRM_COLUMNS]
    else:
        # Empty DataFrame but with correct columns so downstream code doesn't crash
        crm_df = pd.DataFrame(columns=FULL_CRM_COLUMNS)

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
        print(f"       ├───── Mode: Single Sheet (All Data)")
        print(f"       └───── Enriched (all rows)    : {len(all_df)}")

    else:
        print(f"       ├───── Mode: Multi Sheets")
        print(f"       ├───── Enriched (all rows)    : {len(all_df)}")
        print(f"       ├───── Matched  (API success) : {len(matched_df)}")
        print(f"       ├───── Unmatched              : {len(unmatched_df)}")
        print(f"       └───── CRM Save (bulk_create) : {len(crm_df)}")

    if unmatched_pids:
        print(
            f"\n  Unmatched person IDs ({len(unmatched_pids)}): "
            f"{', '.join(unmatched_pids[:5])}"
            + ("..." if len(unmatched_pids) > 5 else "")
        )

    # 🔥 Clean full dataset before returning
    all_df = all_df.fillna("").astype(str)
    all_df = all_df.apply(lambda col: col.str.strip())

    # return crm_df, 
    return crm_df, all_df




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

    # created_ids = set()
    # existing_ids = set()

    print(f"\n  Saving {len(contacts_list)} contacts to Apollo CRM in {len(batches)} batch(es)...\n")

    for b_idx, batch in enumerate(batches, 1):
        print(f"  CRM Batch {b_idx:>3}/{len(batches)}  ({len(batch)} contacts)...", end=" ", flush=True)
        try:
            resp     = bulk_create_crm_batch(batch)
            
            created_list  = resp.get("created_contacts",  [])
            existing_list = resp.get("existing_contacts", [])

            total_created  += len(created_list)
            total_existing += len(existing_list)

            print(f"✓  created={len(created_list)}  existing={len(existing_list)}")

        except requests.exceptions.HTTPError as e:
            print(f"\n    ✗  HTTP Error: {e}")
        except Exception as e:
            print(f"\n    ✗  Error: {e}")

        if b_idx < len(batches):
            time.sleep(REQUEST_DELAY)

    print(f"\n  ── CRM Save Completed ✅ ──────────────────────────────────")
    print(f"     Total created  : {total_created}")
    print(f"     Total existing : {total_existing}  (already in your CRM, not duplicated)")
    # print(f"     Total created  : {len(created_ids)}")
    # print(f"     Total existing : {len(existing_ids)}  (already in your CRM, not duplicated)")

    # 🔥 RETURN crm_df (we use it to mark CRM contacts)
    # return crm_df

    # 🔥 RETURN THIS (CRITICAL)
    # return created_ids, existing_ids


"""
═══════════════════════════════════════════════════════════════
  Update MASTER DB with enriched + processed data

  Rules:
    1. NEVER create duplicates (based on apollo_person_id)
    2. UPDATE existing rows intelligently (field-level)
    3. INSERT new rows if not already present
    4. PRESERVE existing enriched data (do not overwrite with empty values)
    5. ALWAYS update important fields like email if new value exists

  Input:
    new_df → Final processed DataFrame (after enrichment merge)

  Output:
    Updates MASTER_DB file in-place
═══════════════════════════════════════════════════════════════
"""
# v14 : is_enriched → sticky TRUE
#       is_crm_contact → controlled TRUE only when contact saved to CRM
def update_master_db_enriched(new_df):

    # ════════════════════════════════════════
    # STEP 0: FORCE CLEAN SCHEMA (CRITICAL FIX)
    # ════════════════════════════════════════

    new_df = new_df.copy()

    # Normalize column names (avoid case mismatch bugs)
    new_df.columns = [c.strip().lower() for c in new_df.columns]

    # ════════════════════════════════════════
    # COLUMN MAPPING (OUTPUT → MASTER SCHEMA)
    # ════════════════════════════════════════
    COLUMN_MAPPING = {
        "first name": "first_name",
        "last name": "last_name",
        "person linkedin url": "linkedin_url",
        "company linkedin url": "org_linkedin_url",
        "company name": "organization_name",
        "website": "website_url",
        "industry": "org_industry",
        "# employees": "org_employees",
        "email": "email",
        "email status": "email_status",
        "city": "city",
        "state": "state",
        "country": "country",
        # 'is_enriched': 'isenriched',      # ← ADD THIS
        # 'is_crm_contact':        'iscrmcontact',    # ← ADD THIS
        # 'last_updated':          'lastupdated',     # ← ADD THIS
    }

    # 🔥 FORCE enrichment flag (critical fix)
    # if "is_enriched" in new_df.columns:
    #     new_df["is_enriched"] = new_df["is_enriched"].apply(
    #         lambda x: "true" if str(x).strip().lower() == "true" else ""
    #     )

    # Rename columns
    new_df.rename(columns=COLUMN_MAPPING, inplace=True)

    # REMOVE DUPLICATE COLUMNS (CRITICAL)
    new_df = new_df.loc[:, ~new_df.columns.duplicated()]

    # Keep only allowed columns
    new_df = new_df[[col for col in FINAL_COLUMNS if col in new_df.columns]]

    # Add missing columns (to maintain structure)
    for col in FINAL_COLUMNS:
        if col not in new_df.columns:
            new_df[col] = ""

    # # STEP 1: Load existing master DB (if exists)
    # if os.path.exists(MASTER_DB):
    #     master_df = pd.read_excel(MASTER_DB, dtype=str)
    # else:
    #     # Create empty master DB with same schema
    #     master_df = pd.DataFrame(columns=new_df.columns)

    # STEP 1: Load existing master DB (if exists)
    if os.path.exists(MASTER_DB):
        master_df = pd.read_excel(MASTER_DB, dtype=str).copy()  # 🔥 .copy() prevents read-only error
    else:
        master_df = pd.DataFrame(columns=new_df.columns)

    # ════════════════════════════════════════
    # CLEAN MASTER DB SCHEMA
    # ════════════════════════════════════════
    master_df.columns = [c.strip().lower() for c in master_df.columns]

    # REMOVE DUPLICATE COLUMNS (CRITICAL)
    master_df = master_df.loc[:, ~master_df.columns.duplicated()]

    # Keep only valid columns
    master_df = master_df[[col for col in FINAL_COLUMNS if col in master_df.columns]]

    # Add missing columns
    for col in FINAL_COLUMNS:
        if col not in master_df.columns:
            master_df[col] = ""

    # STEP 2: Normalize nulls (avoid NaN issues)
    master_df = master_df.fillna("")
    new_df = new_df.fillna("")

    # 🔥 FORCE ENRICHMENT FLAG (CRITICAL FIX)
    new_df["is_enriched"] = "true"

    # 🔥 DEFAULT CRM FLAG
    if "is_crm_contact" not in new_df.columns:
        new_df["is_crm_contact"] = "false"

    # STEP 3: Use apollo_person_id as index for fast lookup
    master_df.set_index("apollo_person_id", inplace=True)

    # ════════════════════════════════════════
    # REMOVE DUPLICATE IDS (CRITICAL FIX)
    # ════════════════════════════════════════
    # master_df = master_df[~master_df.index.duplicated(keep="first")]
    master_df = master_df[~master_df.index.duplicated(keep="first")].copy()  # 🔥 force writable

    new_df.set_index("apollo_person_id", inplace=True)

    # ════════════════════════════════════════
    # REMOVE DUPLICATES FROM NEW DATA (CRITICAL)
    # ════════════════════════════════════════
    new_df = new_df[~new_df.index.duplicated(keep="last")]

    updated_count = 0
    inserted_count = 0

    # STEP 4: Iterate through new data
    for pid, new_row in new_df.iterrows():

        if pid in master_df.index:

            # ───────────────────────────────────────────────
            # UPDATE existing row (field-level intelligent update)
            # ───────────────────────────────────────────────
            for col in new_df.columns:

                # Get existing value (safe fallback)
                # old_val = master_df.at[pid, col] if col in master_df.columns else ""
                if col in master_df.columns:
                    old_val = master_df.loc[pid, col]

                    # 🔥 HANDLE SERIES CASE (FINAL FIX)
                    if isinstance(old_val, pd.Series):
                        old_val = old_val.iloc[0]
                else:
                    old_val = ""

                # 🔥 Rule 1: Fill missing values
                # if not old_val or old_val in ["", "nan", None]:
                if pd.isna(old_val) or str(old_val).strip().lower() in ["", "nan", "none"]:
                    master_df.at[pid, col] = new_row[col]

                # 🔥 Rule 2: Always update email if new one exists
                # elif col in ["Email", "Email Status"] and new_row[col]:
                # elif col in ["email", "email_status"] and new_row[col]:
                elif col in ["email", "email_status"] and str(new_row[col]).strip():
                    master_df.at[pid, col] = new_row[col]

                # 🔥 Rule 3: ALWAYS update enrichment flags - AGAIN ISSUE.
                # elif col in ["is_enriched", "is_crm_contact"]:
                #     master_df.at[pid, col] = new_row[col]

                # 🔥 Rule 3: Smart update for flags
                # elif col == "is_enriched":
                #     old_val = str(old_val).lower()
                #     new_val = str(new_row[col]).lower()

                #     # Once TRUE → always TRUE
                #     if new_val == "true":
                #         master_df.at[pid, col] = "true"
                #     elif old_val != "true":
                #         master_df.at[pid, col] = new_val

                # elif col == "is_enriched":
                # # 🔥 FINAL RULE: enrichment attempted → always true
                #     master_df.at[pid, col] = "true"

                elif col == "is_enriched":
                    new_val = str(new_row[col]).strip().lower()

                    # Always enforce valid values
                    if new_val not in ["true", "false"]:
                        new_val = "false"

                    # Once true → always true
                    if new_val == "true":
                        master_df.at[pid, col] = "true"
                    elif str(old_val).lower() != "true":
                        master_df.at[pid, col] = "false"

                elif col == "is_crm_contact":
                    new_val = str(new_row[col]).strip().lower()
                
                    # 🔥 STICKY TRUE — once a CRM contact, always a CRM contact
                    # Never downgrade from "true" to "false"
                    if new_val == "true":
                        master_df.at[pid, col] = "true"
                    # else: keep whatever is already in master DB — do NOT overwrite with "false"


                # Causing issue and changing is_crm_contact to false for existing CRM contacts - START
                # elif col == "is_crm_contact":
                #     new_val = str(new_row[col]).strip().lower()

                #     # Only update when explicitly TRUE
                #     if new_val == "true":
                #         master_df.at[pid, col] = "true"
                #     else:
                #         master_df.at[pid, col] = "false"
                # Causing issue and changing is_crm_contact to false for existing CRM contacts - END

                # RULE 3: Keep existing value otherwise
                # (do nothing)

            # Update timestamp
            master_df.at[pid, "last_updated"] = TIMESTAMP

            updated_count += 1

        # else:
        #     # ════════════════════════════════════════
        #     # INSERT new record
        #     # ════════════════════════════════════════
        #     master_df.loc[pid] = new_row

        #     # 🔥 ensure flags are correct for new records
        #     master_df.at[pid, "is_enriched"] = "true"

        #     inserted_count += 1

        else:
            # ════════════════════════════════════════
            # INSERT new record
            # ════════════════════════════════════════
            new_row_dict = new_row.to_dict() if hasattr(new_row, 'to_dict') else dict(new_row)
            new_row_dict["is_enriched"] = "true"  # 🔥 set flag BEFORE inserting
            master_df.loc[pid] = new_row_dict

            inserted_count += 1

    # STEP 5: Reset index back to column
    master_df.reset_index(inplace=True)

    # ════════════════════════════════════════
    # FINAL COLUMN CONTROL (LAST DEFENSE)
    # ════════════════════════════════════════
    master_df = master_df[FINAL_COLUMNS]

    # 🔥 FINAL SAFETY — NO EMPTY FLAGS EVER
    master_df["is_enriched"] = master_df["is_enriched"].apply(
        lambda x: "true" if str(x).strip().lower() == "true" else "false"
    )
    
    master_df["is_crm_contact"] = master_df["is_crm_contact"].apply(
        lambda x: "true" if str(x).strip().lower() == "true" else "false"
    )

    # STEP 6: Save updated master DB
    master_df.to_excel(MASTER_DB, index=False)

    total_enriched = (master_df["is_enriched"].astype(str).str.lower() == "true").sum()
    total_crm      = (master_df["is_crm_contact"].astype(str).str.lower() == "true").sum()
    total_emails   = (master_df["email"].astype(str).str.strip() != "").sum()

    # STEP 7: Logging summary
    print(f"\n  🧠  Master DB Updated Successfully ✅")
    print(f"            ├───── Updated existing      : {updated_count}")
    print(f"            ├───── Inserted new          : {inserted_count}")
    print(f"            ├───── Total records in DB   : {len(master_df)}")
    print(f"            ├───── Enriched (API called) : {total_enriched}")
    print(f"            ├───── In CRM                : {total_crm}")
    print(f"            └───── Have email            : {total_emails}")
    print(f"\n  👀  View Master DB File 👉 → 'output/master_lead_database.xlsx'")

    # STEP 7: Logging summary
    # print(f"\n  🧠 Master DB Updated Successfully ✅")
    # print(f"     Updated existing records : {updated_count}")
    # print(f"     Inserted new records     : {inserted_count}")
    # print(f"     Total records in DB      : {len(master_df)}")
    # print(f"\n  👀 View Master DB File 👉 'output/master_lead_database.xlsx'")


def print_run_complete_summary(enriched_count: int, credits: int):
    print(f"\n  {'═' * 60}")
    print(f"  ✅  Run Complete 🏁")
    print(f"       ├───── ✨ Enriched this run      : {enriched_count}  {'(all skipped)' if enriched_count == 0 else 'contacts'}")
    print(f"       ├───── 💰 Credits consumed       : {credits}  {'🎉' if credits == 0 else ''}")
    print(f"       ├───── 🔽 Output saved           : output/{os.path.basename(OUTPUT_FILE)}")
    print(f"       └───── 🔐 Master DB              : output/master_lead_database.xlsx")
    print(f"  {'═' * 60}\n")


# ╔══════════════════════════════════════════════════════════════╗
# ║                         MAIN RUNNER                         ║
# ╚══════════════════════════════════════════════════════════════╝

# v11 : update_master_db changed to update_master_db_enriched => it cleans data + batches logic changed to only use the length of person_ids_all instead of total.
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

    # STEP 1: Load Input
    input_df = load_input_file()
    total = len(input_df)

    # Don't Take the Ids from the input file directly Now
    # person_ids_all = input_df["apollo_person_id"].astype(str).str.strip().tolist()

    # STEP 2 : Take Only the IDs that need enrichment based on our smart filter (cost control)
    person_ids_all = get_enrichment_candidates(input_df)

    # If no IDs to enrich after filtering, skip API calls and just save empty enriched rows
    # if not person_ids_all:
    #     print("\n ⚠ No contacts to enrich after smart filtering. Skipping API calls.")
    #     return

    # 🔥 IMPORTANT: If no IDs to enrich,
    # we still want to save an output file with empty enriched rows and update the master DB with is_enriched=True for those records,
    # so we don't lose track of them in future runs.
    # So we don't return here, but instead create an empty enrich_map and continue to save_output and update_master_db with empty data.

    # ───────────────────────────────────────────────
    # STEP 3: Handle NO ENRICHMENT case
    # ───────────────────────────────────────────────
    if not person_ids_all:
        print("\n ⚠ No contacts to enrich after smart filtering. Skipping API calls.")
        print("\n ✅ Completed (no enrichment needed)")
        print("\n ↗️  Still We will save OUTPUT to File + update MASTER DB with is_enriched=True for these records, so they are not lost in future runs.")

        # enrich_map = {} # 🔥 important: empty map
        # crm_df = save_output(enrich_map, input_df)
        # update_master_db(input_df, enrich_map)

        enrich_map = {}

        # crm_df = save_output(enrich_map, input_df)

        # ✅ SINGLE SOURCE OF TRUTH
        crm_df, full_df = save_output(enrich_map, input_df)

        # ───────────────────────────────────────────────
        # STEP: No enrichment, but mark records as processed
        # ───────────────────────────────────────────────

        # 🔥 IMPORTANT: Since no enrichment happened,
        # mark all rows as NOT enriched (False)
        # input_df["is_enriched"] = False

        # Clean string values (avoid hidden bugs like spaces, NaN, etc.)
        # input_df = input_df.applymap(lambda x: str(x).strip() if pd.notna(x) else "")

        # ✅ Update master DB from final dataset
        update_master_db_enriched(full_df)

        if CREATE_IN_CRM and not crm_df.empty:
            # ✅ Run CRM save (no return value needed)
            run_bulk_create_crm(crm_df)

            # ✅ Mark CRM contacts using crm_df directly
            # apollo_person_id is ALWAYS present — works even for contacts with no email
            # (Apollo's response does NOT return apollo_person_id reliably, so we use what WE sent)
            crm_sent_pids = set(crm_df["apollo_person_id"].astype(str).str.strip())
            full_df["is_crm_contact"] = full_df["apollo_person_id"].apply(
                lambda x: "true" if str(x).strip() in crm_sent_pids else "false"
            )

            # 🔥 Final update to master DB with correct CRM flag
            update_master_db_enriched(full_df)

        # full_df = pd.read_excel(OUTPUT_FILE, sheet_name="Enriched")
        # update_master_db_enriched(full_df)
        print_run_complete_summary(enriched_count=0, credits=0)
        return

    # batches = [person_ids_all[i:i+BATCH_SIZE] for i in range(0, total, BATCH_SIZE)] # Commented because we should batch based on the number of IDs we are enriching, not the total input rows

    # ───────────────────────────────────────────────
    # STEP 4: Batch Preparation for Enrichment API Calls
    # ───────────────────────────────────────────────

    # 🔥 IMPORTANT: Use length of person_ids_all (after filtering) to determine batches, not total input rows
    batches = [
        person_ids_all[i:i+BATCH_SIZE]
        for i in range(0, len(person_ids_all), BATCH_SIZE)
    ]

    total_batches = len(batches)
    total_credits = 0

    print(f"  Processing {total} contacts in {total_batches} batch(es) of {BATCH_SIZE}...\n")

    enrich_map = {}  # apollo_person_id → flattened enriched dict

    # ───────────────────────────────────────────────
    # STEP 5: Enrichment Loop
    # ───────────────────────────────────────────────
    for b_idx, batch_ids in enumerate(batches, 1):
        print(f"  Batch {b_idx:>3}/{total_batches}  ({len(batch_ids)} IDs)...", end=" ", flush=True)
        try:
            matches, raw_resp = bulk_match_batch(batch_ids)
        except requests.exceptions.HTTPError as e:
            print(f"\n  ✗ HTTP Error: {e}")
            continue
        except Exception as e:
            print(f"\n  ✗ Error: {e}")
            continue

        credits  = raw_resp.get("credits_consumed", 0)
        enriched = raw_resp.get("unique_enriched_records", len(matches))
        missing  = raw_resp.get("missing_records", 0)

        total_credits += credits

        print(f"✓  matched={enriched}  missing={missing}  credits={credits}")

        # ✅ Store using SINGLE PRIMARY KEY
        for m in matches:
            flat = flatten_match(m)
            # pid = flat.get("Apollo Contact Id", "").strip()
            pid = flat.get("apollo_person_id", "").strip()

            if pid:
                enrich_map[pid] = flat

        if b_idx < total_batches:
            time.sleep(REQUEST_DELAY)

    # ───────────────────────────────────────────────
    # STEP 6: Handle ZERO MATCHES edge case
    # ───────────────────────────────────────────────
    if not enrich_map:
        print("\n  ⚠ No matches returned from API. Check your API key and input file.")
        print("   Saving output + updating master DB anyway...\n")

        crm_df, full_df = save_output(enrich_map, input_df)

        # added at last moment by GPT, lets hope not fuck up everything 🙏
        full_df["is_crm_contact"] = "false"

        update_master_db_enriched(full_df)
        return

    # ───────────────────────────────────────────────
    # STEP 7: Save Output (SINGLE SOURCE OF TRUTH)
    # ───────────────────────────────────────────────
    print("\n" + "─" * 62)

    # crm_df = save_output(enrich_map, input_df)
    crm_df, full_df = save_output(enrich_map, input_df)


    # ── Sheet reference (helpful for whoever opens the Excel) ──
    print(f"  ─── SHEETS GUIDE ────────────────────────────────────────")
    print(f"  'Enriched'  = ALL input rows (matched + unmatched, same order)")
    print(f"  'Matched'   = Rows where Apollo API returned data (not guaranteed to have email) ")
    print(f"  'Unmatched' = Rows Apollo could not find any match")
    print(f"  'CRM Save'  = Ready for CRM upload via 'contacts/bulk_create' API")
    print(f"  ─────────────────────────────────────────────────────────")

    print(f"  🚨🚨🚨 Total enrichment credits consumed : {total_credits}")

    # Update master DB with enriched data and enrichment status (is_enriched=True if email found, else False)
    # update_master_db(input_df, enrich_map)

    # ───────────────────────────────────────────────
    # STEP: Merge enrichment into input_df
    # ───────────────────────────────────────────────

    # for idx, row in input_df.iterrows():
    #     pid = row["apollo_person_id"]

    #     if pid in enrich_map:
    #         enriched = enrich_map[pid]

    #         # 🔥 Apply enrichment fields
    #         for key, value in enriched.items():
    #             input_df.at[idx, key] = value

    # Update enrichment flag
    # input_df["is_enriched"] = input_df["email"].apply(lambda x: True if x else False)

    # 🔥 FINAL RULE: If record reached here → enrichment was attempted
    # input_df["is_enriched"] = "true"

    # Clean string values (avoid hidden bugs like spaces, NaN, etc.)
    # input_df = input_df.applymap(lambda x: str(x).strip() if pd.notna(x) else "")

    # ───────────────────────────────────────────────
    # STEP: Update master DB
    # ───────────────────────────────────────────────

    # update_master_db_enriched(full_df)

    # print(f"  Total enrichment credits consumed : {total_credits}")

    # ───────────────────────────────────────────────
    # STEP 8: CRM Save + update CRM Flag in master DB
    # ───────────────────────────────────────────────
    if CREATE_IN_CRM:
        if crm_df.empty:
            print("\n  ⚠ No matched contacts to save to CRM.")
            full_df["is_crm_contact"] = "false"
        else:
            # ✅ Run CRM save (no return value needed)
            # Apollo's response contacts have CRM IDs (not apollo_person_id),
            # so we match using what WE sent — crm_df already has apollo_person_id for every row.
            # This works for contacts with OR without email.
            run_bulk_create_crm(crm_df)

            # ✅ Mark CRM contacts using crm_df directly (apollo_person_id is ALWAYS present)
            crm_sent_pids = set(crm_df["apollo_person_id"].astype(str).str.strip())
            full_df["is_crm_contact"] = full_df["apollo_person_id"].apply(
                lambda x: "true" if str(x).strip() in crm_sent_pids else "false"
            )

    else:
        full_df["is_crm_contact"] = "false"

        print()
        out_name = os.path.basename(OUTPUT_FILE)
        print("  ─── TO SAVE CONTACTS TO APOLLO CRM ─────────────────────")
        print(f"  1. Review the 'CRM Save' sheet in 'output/{out_name}'")
        print("  2. Remove any rows you DON'T want to create")
        print("  3. Set CREATE_IN_CRM = True in CONFIGURATION")
        print("  4. Re-run this script (it will skip enrichment step?)")
        print("  — OR just set CREATE_IN_CRM = True on first run")
        print("─" * 62 + "\n")

    # ───────────────────────────────────────────────
    # STEP 9: FINAL MASTER DB UPDATE (AFTER CRM)
    # ───────────────────────────────────────────────
    # 🔥 Single call here — is_crm_contact is already correctly set above
    update_master_db_enriched(full_df)

    # Just Call the Run Summary function at the end with correct enriched count and credits
    print_run_complete_summary(enriched_count=len(person_ids_all), credits=total_credits)

    # ───────────────────────────────────────────────
    # FINAL RUN SUMMARY
    # ───────────────────────────────────────────────
    # print(f"\n  {'*' * 60}")
    # print(f"   ✅  Run Complete 🏁!")
    # print(f"   ├─ ✨Enriched this run      : {len(person_ids_all)} contacts")
    # print(f"   ├─ 💰Credits consumed       : {total_credits}")
    # print(f"   ├─ 🔽Output saved           : output/{os.path.basename(OUTPUT_FILE)}")
    # print(f"   └─ 🔐Master DB              : output/master_lead_database.xlsx")
    # print(f"  {'*' * 60}\n")


if __name__ == "__main__":
    main()