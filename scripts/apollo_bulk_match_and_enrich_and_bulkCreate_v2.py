#!/usr/bin/env python3
"""
═══════════════════════════════════════════════════════════════
  Apollo Contact Scraper — STEP 2: Bulk People Enrichment
  Uses: POST /api/v1/people/bulk_match

  INPUT : Filtered Excel/CSV from Step 1
          (keep only the rows you want, save as new file)
  OUTPUT: output/apollo_enriched_DDMMMYYYY_HHMMSS.xlsx
          Excel file with enriched data — email, phone, title,
          LinkedIn, employment history, org details, etc.

  LIMIT : 10 people per API call (Apollo hard limit)
          Script batches automatically.
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

# API_KEY = "l4AkVpSXwV9fICkebUMAXw"       # Your Apollo master API key
API_KEY = "vWTBtYd1P9IpMLV-wghAzw"       # Your Apollo master API key

TIMESTAMP = datetime.now().strftime("%d%b%Y_%H%M%S").lower()   # e.g. 27apr2026_124055

# Input file = your manually filtered Excel/CSV from Step 1
# Only column needed from Step 1: apollo_person_id
BULK_ENRICH_INPUT_FILE = "apollo_people_search_20260423_144441.xlsx"

# Output file — timestamped so it is never overwritten
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"apollo_enriched_{TIMESTAMP}.xlsx")

# Apollo hard limit: max 10 people per bulk_match call
BATCH_SIZE = 10

# Seconds to wait between batches to avoid rate-limiting
REQUEST_DELAY = 1.2

# Set True to reveal personal emails (uses extra credits)
REVEAL_PERSONAL_EMAILS = True

# Set True to reveal phone numbers (uses extra credits)
REVEAL_PHONE_NUMBER = False


# ╔══════════════════════════════════════════════════════════════╗
# ║                     HELPER FUNCTIONS                        ║
# ╚══════════════════════════════════════════════════════════════╝

BASE_URL = "https://api.apollo.io/api/v1"


def load_input_file():
    """
    Load the filtered Step 1 output (Excel or CSV).
    Must contain column: apollo_person_id
    All other Step 1 columns are preserved in the final output.
    """
    path = BULK_ENRICH_INPUT_FILE
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Input file not found: '{path}'\n"
            f"Set BULK_ENRICH_INPUT_FILE to the path of your filtered Step 1 output."
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

    # Drop rows with empty apollo_person_id
    before = len(df)
    df = df[df["apollo_person_id"].notna() & (df["apollo_person_id"].astype(str).str.strip() != "")]
    dropped = before - len(df)
    if dropped:
        print(f"  ⚠  Dropped {dropped} rows with empty apollo_person_id")

    print(f"  Loaded {len(df)} contacts from '{path}'")
    return df.reset_index(drop=True)


def bulk_match_batch(person_ids):
    """
    POST one batch (max 10 IDs) to /people/bulk_match.
    Each detail object only needs the 'id' field (apollo_person_id).
    Returns list of matched person objects (may be fewer than input if no match).
    """
    details = [{"id": pid} for pid in person_ids]

    response = requests.post(
        f"{BASE_URL}/people/bulk_match",
        params={
            "reveal_personal_emails": str(REVEAL_PERSONAL_EMAILS).lower(),
            "reveal_phone_number":    str(REVEAL_PHONE_NUMBER).lower(),
        },
        json={"details": details},
        headers={
            "Cache-Control":  "no-cache",
            "Content-Type":   "application/json",
            "Accept":         "application/json",
            "x-api-key":      API_KEY,
        },
        timeout=60,
    )

    if response.status_code == 429:
        wait = int(response.headers.get("Retry-After", 30))
        print(f"\n    ⚠  Rate limited — waiting {wait}s...")
        time.sleep(wait)
        return bulk_match_batch(person_ids)   # retry

    response.raise_for_status()
    data = response.json()
    return data.get("matches", []), data


def flatten_match(match):
    """
    Flatten one Apollo match object from bulk_match response
    into a flat dict for the output Excel.
    Captures all useful fields from the sample response.
    """
    def safe(obj, *keys):
        """Safe nested get."""
        for k in keys:
            if not isinstance(obj, dict):
                return ""
            obj = obj.get(k)
            if obj is None:
                return ""
        return obj if obj is not None else ""

    org  = match.get("organization") or {}
    acct = match.get("account") or {}

    # Employment history — current role
    history = match.get("employment_history") or []
    current_roles = [h for h in history if h.get("current")]
    current = current_roles[0] if current_roles else {}

    return {
        # ── Identity ──────────────────────────────────────────────
        "apollo_person_id":         safe(match, "id"),
        "first_name":               safe(match, "first_name"),
        "last_name":                safe(match, "last_name"),
        "full_name":                safe(match, "name"),
        "title":                    safe(match, "title"),
        "headline":                 safe(match, "headline"),
        "seniority":                safe(match, "seniority"),
        "departments":              ", ".join(match.get("departments") or []),
        "subdepartments":           ", ".join(match.get("subdepartments") or []),
        "functions":                ", ".join(match.get("functions") or []),

        # ── Contact details ───────────────────────────────────────
        "email":                    safe(match, "email"),
        "email_status":             safe(match, "email_status"),
        "linkedin_url":             safe(match, "linkedin_url"),
        "twitter_url":              safe(match, "twitter_url"),
        "github_url":               safe(match, "github_url"),
        "facebook_url":             safe(match, "facebook_url"),
        "photo_url":                safe(match, "photo_url"),

        # ── Location ──────────────────────────────────────────────
        "city":                     safe(match, "city"),
        "state":                    safe(match, "state"),
        "country":                  safe(match, "country"),

        # ── Current employment (from history) ─────────────────────
        "current_company":          safe(current, "organization_name"),
        "current_title":            safe(current, "title"),
        "current_role_start":       safe(current, "start_date"),

        # ── Organisation (Apollo DB) ───────────────────────────────
        "org_id":                   safe(org, "id"),
        "org_name":                 safe(org, "name"),
        "org_website":              safe(org, "website_url"),
        "org_linkedin":             safe(org, "linkedin_url"),
        "org_twitter":              safe(org, "twitter_url"),
        "org_industry":             safe(org, "industry"),
        "org_employees":            safe(org, "estimated_num_employees"),
        "org_founded":              safe(org, "founded_year"),
        "org_country":              safe(org, "country"),
        "org_city":                 safe(org, "city"),
        "org_state":                safe(org, "state"),
        "org_phone":                safe(org, "phone"),

        # ── Apollo CRM Account ─────────────────────────────────────
        "apollo_account_id":        safe(acct, "id"),
        "account_name":             safe(acct, "name"),
        "account_domain":           safe(acct, "domain"),
        "account_phone":            safe(acct, "phone"),

        # ── Engagement signals ─────────────────────────────────────
        "is_likely_to_engage":      safe(match, "is_likely_to_engage"),
        "revealed_for_team":        safe(match, "revealed_for_current_team"),
    }


def save_output(enriched_df, input_df):
    """
    Write output Excel with three sheets:
      1. Enriched  — Step 1 columns + all enriched columns merged
      2. Matched   — only enriched rows (contact found in Apollo)
      3. Unmatched — input rows where Apollo returned no match
    """
    # Build a lookup: apollo_person_id → enriched row
    enrich_map = {}
    for _, row in enriched_df.iterrows():
        pid = str(row.get("apollo_person_id", "")).strip()
        if pid:
            enrich_map[pid] = row

    # Merge enriched columns back into input_df
    enrich_cols = [c for c in enriched_df.columns if c != "apollo_person_id"]
    merged_rows = []
    unmatched_rows = []

    for _, in_row in input_df.iterrows():
        pid = str(in_row.get("apollo_person_id", "")).strip()
        base = in_row.to_dict()

        if pid in enrich_map:
            enriched = enrich_map[pid].to_dict()
            # Prefix enriched cols with 'enriched_' if name clashes with input
            for col in enrich_cols:
                key = f"enriched_{col}" if col in base and col != "apollo_person_id" else col
                base[key] = enriched.get(col, "")
            merged_rows.append(base)
        else:
            for col in enrich_cols:
                key = f"enriched_{col}" if col in base and col != "apollo_person_id" else col
                base[key] = ""
            unmatched_rows.append(base)

    merged_df    = pd.DataFrame(merged_rows)
    unmatched_df = pd.DataFrame(unmatched_rows) if unmatched_rows else pd.DataFrame(columns=merged_df.columns)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        # Sheet 1 — All contacts (matched + unmatched)
        merged_df.to_excel(writer, index=False, sheet_name="Enriched")

        # Sheet 2 — Only successfully enriched
        enriched_only = merged_df[merged_df["email"].notna() & (merged_df["email"].astype(str).str.strip() != "")]
        enriched_only.to_excel(writer, index=False, sheet_name="Matched")

        # Sheet 3 — Unmatched / no data returned
        unmatched_df.to_excel(writer, index=False, sheet_name="Unmatched")

        # Auto-fit column widths
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 55)

    print(f"\n  ✅  Saved → '{OUTPUT_FILE}'")
    print(f"       ├─ Total rows   : {len(merged_df)}")
    print(f"       ├─ Matched      : {len(enriched_only)}")
    print(f"       └─ Unmatched    : {len(unmatched_df)}")


# ╔══════════════════════════════════════════════════════════════╗
# ║                        MAIN RUNNER                         ║
# ╚══════════════════════════════════════════════════════════════╝

def main():
    print("\n" + "═" * 62)
    print("  Apollo Bulk People Enrichment — Step 2 of 2")
    print("  Endpoint: POST /api/v1/people/bulk_match")
    print("═" * 62)
    print(f"  Input file            : {BULK_ENRICH_INPUT_FILE}")
    print(f"  Output file           : {OUTPUT_FILE}")
    print(f"  Batch size            : {BATCH_SIZE} (Apollo max per call)")
    print(f"  Reveal personal email : {REVEAL_PERSONAL_EMAILS}")
    print(f"  Reveal phone number   : {REVEAL_PHONE_NUMBER}")
    print("═" * 62 + "\n")

    # ── Load input ────────────────────────────────────────────────
    input_df = load_input_file()
    total    = len(input_df)

    # ── Split into batches of 10 ──────────────────────────────────
    person_ids_all = input_df["apollo_person_id"].astype(str).str.strip().tolist()
    batches        = [person_ids_all[i:i+BATCH_SIZE] for i in range(0, total, BATCH_SIZE)]
    total_batches  = len(batches)
    total_credits  = 0

    print(f"  Processing {total} contacts in {total_batches} batch(es) of {BATCH_SIZE}...\n")

    all_matches = []   # list of flattened match dicts

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
            all_matches.append(flatten_match(m))

        if b_idx < total_batches:
            time.sleep(REQUEST_DELAY)

    # ── Build results DataFrame ───────────────────────────────────
    if not all_matches:
        print("\n  ⚠  No matches returned. Check API key and input file.")
        return

    enriched_df = pd.DataFrame(all_matches)

    # ── Save output ───────────────────────────────────────────────
    print("\n" + "─" * 62)
    save_output(enriched_df, input_df)

    print()
    print(f"  Total API credits consumed this run : {total_credits}")
    print()
    print("  ─── WORKFLOW COMPLETE ───────────────────────────────────")
    print(f"  Open '{OUTPUT_FILE}'")
    print("  'Enriched'  sheet = all contacts with enriched data")
    print("  'Matched'   sheet = contacts with email found")
    print("  'Unmatched' sheet = contacts Apollo could not match")
    print("─" * 62 + "\n")


if __name__ == "__main__":
    main()
