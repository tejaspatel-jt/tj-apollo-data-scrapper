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
    Returns:
        - list of apollo_person_ids to enrich
        - skipped_df (for logging/debug)
    """

    # Load master DB if exists
    if os.path.exists(MASTER_DB):
        master_df = pd.read_excel(MASTER_DB, dtype=str)
        master_df["apollo_person_id"] = master_df["apollo_person_id"].astype(str).str.strip()
        enriched_ids = set(
            master_df[
                master_df["email"].notna() & (master_df["email"].str.strip() != "")
            ]["apollo_person_id"]
        )
    else:
        enriched_ids = set()

    input_df["apollo_person_id"] = input_df["apollo_person_id"].astype(str).str.strip()

    to_enrich = []
    skipped_rows = []

    for _, row in input_df.iterrows():
        pid = row["apollo_person_id"]
        email = str(row.get("email", "")).strip()
        is_crm = str(row.get("is_crm_contact", "")).lower() == "true"

        if pid in enriched_ids:
            skipped_rows.append((pid, "ALREADY_ENRICHED_MASTER"))
            continue

        if is_crm and email:
            skipped_rows.append((pid, "CRM_WITH_EMAIL"))
            continue

        if email:
            skipped_rows.append((pid, "ALREADY_HAS_EMAIL_STEP1"))
            continue

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

    row = {
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
    }
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


def empty_enriched_row(apollo_person_id=""):
    row = {col: "" for col in DESIRED_COLUMNS}
    row["Apollo Contact Id"] = apollo_person_id
    return row


def enforce_column_order(df):
    for col in DESIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[DESIRED_COLUMNS]


def save_output(enrich_map, input_df):
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

    for _, in_row in input_df.iterrows():
        pid = str(in_row.get("apollo_person_id", "")).strip()
        if pid in enrich_map:
            enriched_rows.append(enrich_map[pid])
        else:
            enriched_rows.append(empty_enriched_row(apollo_person_id=pid))
            unmatched_pids.append(pid)

    all_df       = pd.DataFrame(enriched_rows)
    all_df       = enforce_column_order(all_df)

    has_email    = all_df["Email"].notna() & (all_df["Email"].astype(str).str.strip() != "")
    matched_df   = all_df[has_email].reset_index(drop=True)
    unmatched_df = all_df[~has_email].reset_index(drop=True)

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


def run_bulk_create_crm(crm_df):
    """
    Call POST /api/v1/contacts/bulk_create for all rows in crm_df.
    Batches of 100. Prints created vs existing counts per batch.
    """
    contacts_list = crm_df.to_dict("records")
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
