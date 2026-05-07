#!/usr/bin/env python3
"""
═══════════════════════════════════════════════════════════════
  Apollo Contact Scraper — STEP 2: Bulk People Enrichment
  Uses: POST /api/v1/people/bulk_match

  INPUT : Filtered Excel/CSV from Step 1  (output/ folder)
  OUTPUT: output/apollo_enriched_DDMMMYYYY_HHMMSS.xlsx
          Columns ordered to match Apollo CSV export format.

  LIMIT : 10 people per API call (Apollo hard limit)
          Script batches automatically.

Credits are consumed every time you call /people/bulk_match, even for unmatched records, means for those contacts as well for which no Email Id exist
═══════════════════════════════════════════════════════════════
"""

import requests
import pandas as pd
import time
import os
from datetime import datetime

# ─── resolve paths relative to this script's location ───────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ╔══════════════════════════════════════════════════════════════╗
# ║                     CONFIGURATION                           ║
# ╚══════════════════════════════════════════════════════════════╝

API_KEY = "vWTBtYd1P9IpMLV-wghAzw"        # Your Apollo master API key

TIMESTAMP = datetime.now().strftime("%d%b%Y_%H%M%S").lower()   # e.g. 27apr2026_124055

# Input file — filename inside output/ folder (your filtered Step 1 output)
BULK_ENRICH_INPUT_FILE = "apollo_people_search_20260423_144441.xlsx"

# Output file — timestamped so it is never overwritten
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"apollo_enriched_{TIMESTAMP}.xlsx")

BATCH_SIZE             = 10     # Apollo hard limit: max 10 IDs per bulk_match call
REQUEST_DELAY          = 1.2    # Seconds between batches (avoid rate-limiting)
REVEAL_PERSONAL_EMAILS = True   # True = uses extra credits (requires supported plan)
REVEAL_PHONE_NUMBER    = False  # True = uses extra credits


# ╔══════════════════════════════════════════════════════════════╗
# ║              OUTPUT COLUMN ORDER (Apollo format)            ║
# ╚══════════════════════════════════════════════════════════════╝
#
# Matches Apollo's own CSV export column order exactly.
# Columns not returned by the API are included as empty columns
# so the file is always structurally consistent.
#
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
    "Contact Owner",                        # not in API — always empty
    "Work Direct Phone",
    "Home Phone",
    "Mobile Phone",
    "Corporate Phone",
    "Other Phone",
    "Do Not Call",                          # not in API — always empty
    "Stage",                                # not in API — always empty
    "Lists",                                # not in API — always empty
    "Last Contacted",                       # not in API — always empty
    "Account Owner",                        # not in API — always empty
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
    "Subsidiary of",                        # not in API — always empty
    "Subsidiary of (Organization ID)",      # not in API — always empty
    "Email Sent",                           # CRM activity — always empty
    "Email Open",                           # CRM activity — always empty
    "Email Bounced",                        # CRM activity — always empty
    "Replied",                              # CRM activity — always empty
    "Demoed",                               # CRM activity — always empty
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
    "Qualify Contact",                      # custom field — always empty
]


# ╔══════════════════════════════════════════════════════════════╗
# ║                     HELPER FUNCTIONS                        ║
# ╚══════════════════════════════════════════════════════════════╝

BASE_URL = "https://api.apollo.io/api/v1"


def load_input_file():
    """Load filtered Step 1 output from output/ folder."""
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
        raise ValueError(
            "Column 'apollo_person_id' not found in input file.\n"
            "Make sure you are using the output from Step 1 (apollo_people_search.py)."
        )

    before = len(df)
    df = df[df["apollo_person_id"].notna() & (df["apollo_person_id"].astype(str).str.strip() != "")]
    dropped = before - len(df)
    if dropped:
        print(f"  ⚠  Dropped {dropped} rows with empty apollo_person_id")

    print(f"  Loaded {len(df)} contacts from 'output/{BULK_ENRICH_INPUT_FILE}'")
    return df.reset_index(drop=True)


def bulk_match_batch(person_ids):
    """POST one batch (max 10 IDs) to /people/bulk_match."""
    details = [{"id": pid} for pid in person_ids]

    response = requests.post(
        f"{BASE_URL}/people/bulk_match",
        params={
            "reveal_personal_emails": str(REVEAL_PERSONAL_EMAILS).lower(),
            "reveal_phone_number":    str(REVEAL_PHONE_NUMBER).lower(),
        },
        json={"details": details},
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


def parse_phones(match):
    """Parse phone_numbers array into typed buckets."""
    phones  = match.get("phone_numbers") or []
    buckets = {
        "Work Direct Phone": "",
        "Home Phone":        "",
        "Mobile Phone":      "",
        "Corporate Phone":   "",
        "Other Phone":       "",
    }
    type_map = {
        "work_direct": "Work Direct Phone",
        "home":        "Home Phone",
        "mobile":      "Mobile Phone",
        "corporate":   "Corporate Phone",
        "other":       "Other Phone",
    }
    for p in phones:
        raw    = p.get("sanitized_number") or p.get("raw_number", "")
        ptype  = (p.get("type") or "other").lower().replace(" ", "_")
        bucket = type_map.get(ptype, "Other Phone")
        if raw and not buckets[bucket]:   # keep first occurrence per type
            buckets[bucket] = raw
    return buckets


def parse_emails(match):
    """
    Parse contact_emails array for primary / secondary / tertiary.
    Falls back to top-level email + email_status if array is absent.
    """
    contact_emails = match.get("contact_emails") or []
    contact_emails = sorted(contact_emails, key=lambda x: x.get("position", 99))

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
    """
    Flatten one Apollo bulk_match result into a flat dict
    keyed exactly by DESIRED_COLUMNS names.
    """
    def safe(obj, *keys):
        for k in keys:
            if not isinstance(obj, dict):
                return ""
            obj = obj.get(k)
            if obj is None:
                return ""
        return obj if obj is not None else ""

    org  = match.get("organization") or {}
    acct = match.get("account") or {}

    keywords = ", ".join(org.get("keywords") or [])
    technologies = ", ".join(
        [t.get("name", "") for t in (org.get("technologies") or [])]
    ) if isinstance(org.get("technologies"), list) else ""

    intent_signals = match.get("intent_signals") or []
    intent_1_topic = intent_signals[0].get("topic", "") if len(intent_signals) > 0 else ""
    intent_1_score = intent_signals[0].get("score", "") if len(intent_signals) > 0 else ""
    intent_2_topic = intent_signals[1].get("topic", "") if len(intent_signals) > 1 else ""
    intent_2_score = intent_signals[1].get("score", "") if len(intent_signals) > 1 else ""

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
        "Primary Intent Topic":             intent_1_topic,
        "Primary Intent Score":             intent_1_score,
        "Secondary Intent Topic":           intent_2_topic,
        "Secondary Intent Score":           intent_2_score,
        # CRM / activity — always empty
        "Contact Owner":                    "",
        "Do Not Call":                      "",
        "Stage":                            "",
        "Lists":                            "",
        "Last Contacted":                   "",
        "Account Owner":                    "",
        "Subsidiary of":                    "",
        "Subsidiary of (Organization ID)":  "",
        "Email Sent":                       "",
        "Email Open":                       "",
        "Email Bounced":                    "",
        "Replied":                          "",
        "Demoed":                           "",
        "Qualify Contact":                  "",
    }

    row.update(parse_emails(match))
    row.update(parse_phones(match))
    return row


def empty_enriched_row(apollo_person_id=""):
    """
    Return a blank enriched row (all DESIRED_COLUMNS = "").
    Apollo Contact Id is pre-filled with the Step 1 person ID
    so unmatched rows are still identifiable in the output.
    """
    row = {col: "" for col in DESIRED_COLUMNS}
    row["Apollo Contact Id"] = apollo_person_id
    return row


def enforce_column_order(df):
    """Reindex to DESIRED_COLUMNS; add missing cols as empty, drop extras."""
    for col in DESIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[DESIRED_COLUMNS]


def save_output(enrich_map, input_df):
    """
    Build three sheets preserving input row order throughout:

      Enriched  — ALL input rows in original order.
                  Matched rows have full enriched data.
                  Unmatched rows have empty enriched columns
                  but retain Apollo Contact Id so nothing is lost.
                  Row count == input row count always.

      Matched   — Subset of Enriched where Email is not empty.
      Unmatched — Subset of Enriched where Apollo returned no match.
    """
    enriched_rows  = []
    unmatched_pids = []

    for _, in_row in input_df.iterrows():
        pid = str(in_row.get("apollo_person_id", "")).strip()

        if pid in enrich_map:
            enriched_rows.append(enrich_map[pid])
        else:
            # Keep the row but fill enriched columns with empty strings
            enriched_rows.append(empty_enriched_row(apollo_person_id=pid))
            unmatched_pids.append(pid)

    all_df       = pd.DataFrame(enriched_rows)
    all_df       = enforce_column_order(all_df)

    has_email    = all_df["Email"].notna() & (all_df["Email"].astype(str).str.strip() != "")
    matched_df   = all_df[has_email].reset_index(drop=True)
    unmatched_df = all_df[~has_email].reset_index(drop=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        all_df.to_excel(writer,       index=False, sheet_name="Enriched")
        matched_df.to_excel(writer,   index=False, sheet_name="Matched")
        unmatched_df.to_excel(writer, index=False, sheet_name="Unmatched")

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

    out_name = os.path.basename(OUTPUT_FILE)
    print(f"\n  ✅  Saved → 'output/{out_name}'")
    print(f"       ├─ Enriched (all rows)  : {len(all_df)}  ← matches your input count")
    print(f"       ├─ Matched  (has email) : {len(matched_df)}")
    print(f"       └─ Unmatched            : {len(unmatched_df)}")
    if unmatched_pids:
        print(f"\n  Unmatched apollo_person_ids ({len(unmatched_pids)}):")
        for pid in unmatched_pids:
            print(f"    • {pid}")


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
    print(f"  Batch size            : {BATCH_SIZE} (Apollo max per call)")
    print(f"  Reveal personal email : {REVEAL_PERSONAL_EMAILS}")
    print(f"  Reveal phone number   : {REVEAL_PHONE_NUMBER}")
    print(f"  Output columns        : {len(DESIRED_COLUMNS)}")
    print("═" * 62 + "\n")

    input_df = load_input_file()
    total    = len(input_df)

    person_ids_all = input_df["apollo_person_id"].astype(str).str.strip().tolist()
    batches        = [person_ids_all[i:i+BATCH_SIZE] for i in range(0, total, BATCH_SIZE)]
    total_batches  = len(batches)
    total_credits  = 0

    print(f"  Processing {total} contacts in {total_batches} batch(es) of {BATCH_SIZE}...\n")

    # enrich_map: apollo_person_id → flattened enriched dict
    enrich_map = {}

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
    save_output(enrich_map, input_df)

    out_name = os.path.basename(OUTPUT_FILE)
    print()
    print(f"  Total credits consumed this run : {total_credits}")
    print()
    print("  ─── WORKFLOW COMPLETE ───────────────────────────────────")
    print(f"  Open 'output/{out_name}'")
    print("  'Enriched'  = ALL input rows (matched + unmatched)")
    print("  'Matched'   = contacts where Apollo returned an email")
    print("  'Unmatched' = contacts Apollo could not find")
    print("─" * 62 + "\n")


if __name__ == "__main__":
    main()
