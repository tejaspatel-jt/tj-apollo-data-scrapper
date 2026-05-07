#!/usr/bin/env python3
"""
═══════════════════════════════════════════════════════════════
  Apollo Contact Scraper — STEP 2: Bulk Create Contacts
  Uses: POST /api/v1/contacts/bulk_create

  INPUT : Filtered Excel file from Step 1
          (keep only the rows you want, save as new file)
  OUTPUT: Excel file with enriched contact data from Apollo CRM
          — includes email, email_status, phone, apollo contact ID
═══════════════════════════════════════════════════════════════
"""

import requests
import pandas as pd
import time
import os
from datetime import datetime


# ╔══════════════════════════════════════════════════════════════╗
# ║                     CONFIGURATION                            ║
# ╚══════════════════════════════════════════════════════════════╝

API_KEY = "l4AkVpSXwV9fICkebUMAXw"       # Your Apollo master API key

# Input file = your manually filtered Excel output from Step 1
# Keep only the rows you want to enrich, save as a new .xlsx or .csv
BULK_CREATE_INPUT_FILE = "apollo_people_search_20260423_144441.xlsx"

# Output file — timestamped so it is never overwritten
OUTPUT_FILE = f"apollo_enriched_contacts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# contacts/bulk_create accepts max 100 per request
BATCH_SIZE = 100

# Seconds to wait between batches (avoid rate-limiting)
REQUEST_DELAY = 1.5

# When True, Apollo checks for existing contacts and returns them
# in 'existing_contacts' instead of creating duplicates
RUN_DEDUPE = True


# ╔══════════════════════════════════════════════════════════════╗
# ║                     HELPER FUNCTIONS                        ║
# ╚══════════════════════════════════════════════════════════════╝

BASE_URL = "https://api.apollo.io/api/v1"


def load_input_file():
    """
    Load the filtered Step 1 output.
    Accepts .xlsx or .csv.
    Expected columns (all produced by Step 1):
      first_name, last_name, title, organization_name,
      website_url, linkedin_url, apollo_person_id,
      searched_company, searched_domain  (kept for reference)
    """
    path = BULK_CREATE_INPUT_FILE
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Input file not found: '{path}'\n"
            f"Set BULK_CREATE_INPUT_FILE to the path of your filtered Step 1 output."
        )

    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(path)
    elif ext == ".csv":
        df = pd.read_csv(path)
    else:
        raise ValueError(f"Unsupported file type: {ext}  (use .xlsx or .csv)")

    df.columns = [c.strip().lower() for c in df.columns]
    print(f"  Loaded {len(df)} rows from '{path}'")
    return df


def row_to_contact_payload(row):
    """
    Map one row (from Step 1 output) to the contact object
    expected by contacts/bulk_create.

    Apollo bulk_create fields used:
      first_name, last_name, title, organization_name,
      linkedin_url, website_url

    apollo_person_id is intentionally NOT sent to bulk_create
    (it is not a valid field there); it is kept in our output
    as a cross-reference back to Step 1.
    """
    def val(col):
        v = row.get(col, "")
        return str(v).strip() if pd.notna(v) and str(v).strip() not in ("", "nan") else ""

    payload = {}
    for field in ("first_name", "last_name", "title",
                  "organization_name", "linkedin_url", "website_url"):
        v = val(field)
        if v:
            payload[field] = v

    return payload


def bulk_create_batch(contacts_payload):
    """
    POST one batch (≤100 contacts) to /contacts/bulk_create.
    Returns the full JSON response dict.
    """
    response = requests.post(
        f"{BASE_URL}/contacts/bulk_create",
        json={
            "contacts": contacts_payload,
            "run_dedupe": RUN_DEDUPE,
        },
        headers={
            "Cache-Control": "no-cache",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "x-api-key": API_KEY,
        },
        timeout=60,
    )

    if response.status_code == 429:
        wait = int(response.headers.get("Retry-After", 30))
        print(f"    ⚠  Rate limited — waiting {wait}s...")
        time.sleep(wait)
        return bulk_create_batch(contacts_payload)     # retry once

    response.raise_for_status()
    return response.json()


def flatten_contact(c, status_label):
    """
    Flatten one Apollo contact object returned by bulk_create.
    status_label: 'created' or 'existing'
    """
    def safe(key):
        v = c.get(key, "")
        return v if v is not None else ""

    # Extract primary email
    emails = c.get("contact_emails") or []
    primary_email = ""
    primary_email_status = ""
    if emails:
        primary = next((e for e in emails if e.get("position") == 0), emails[0])
        primary_email = primary.get("email", "")
        primary_email_status = primary.get("email_status", "")
    else:
        # Fallback to top-level email field
        primary_email = safe("email")
        primary_email_status = safe("email_status")

    # Extract primary phone
    phones = c.get("phone_numbers") or []
    primary_phone = ""
    if phones:
        primary_phone = (next((p for p in phones if p.get("position") == 0), phones[0])
                         .get("raw_number", ""))

    return {
        "apollo_contact_id":    safe("id"),
        "apollo_person_id":     safe("person_id"),    # cross-ref to Step 1
        "status":               status_label,         # created | existing
        "first_name":           safe("first_name"),
        "last_name":            safe("last_name"),
        "full_name":            safe("name"),
        "title":                safe("title"),
        "seniority":            safe("seniority"),
        "email":                primary_email,
        "email_status":         primary_email_status,
        "phone":                primary_phone,
        "linkedin_url":         safe("linkedin_url"),
        "city":                 safe("city"),
        "state":                safe("state"),
        "country":              safe("country"),
        "organization_name":    safe("account") and (c.get("account") or {}).get("name", ""),
        "website_url":          safe("account") and (c.get("account") or {}).get("website_url", ""),
        "apollo_account_id":    safe("account") and (c.get("account") or {}).get("id", ""),
        "created_at":           safe("created_at"),
    }


def save_output(results_df, input_df):
    """
    Write three sheets to the output Excel file:
      1. Enriched       — merged result (input cols + enriched cols)
      2. Created        — only newly created contacts
      3. Existing       — contacts already in Apollo CRM
    """
    # Merge enriched data back with original input rows (by position)
    # input_df rows preserved in order; enriched rows aligned by index
    combined = input_df.copy().reset_index(drop=True)
    enriched = results_df.reset_index(drop=True)

    # Add enriched columns to the right of input columns
    for col in enriched.columns:
        if col not in combined.columns:
            combined[col] = enriched[col] if col in enriched.columns else ""
        else:
            combined[f"enriched_{col}"] = enriched[col] if col in enriched.columns else ""

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        # Sheet 1 — full enriched output
        combined.to_excel(writer, index=False, sheet_name="Enriched")

        # Sheet 2 — only newly created
        created = results_df[results_df["status"] == "created"].reset_index(drop=True)
        created.to_excel(writer, index=False, sheet_name="Created")

        # Sheet 3 — already existing in CRM
        existing = results_df[results_df["status"] == "existing"].reset_index(drop=True)
        existing.to_excel(writer, index=False, sheet_name="Existing")

        # Auto-fit columns on all sheets
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 55)

    print(f"\n  ✅  Saved {len(combined)} enriched contacts → '{OUTPUT_FILE}'")
    print(f"       ├─ Created  : {len(created)}")
    print(f"       └─ Existing : {len(existing)}")


# ╔══════════════════════════════════════════════════════════════╗
# ║                        MAIN RUNNER                         ║
# ╚══════════════════════════════════════════════════════════════╝

def main():
    print("\n" + "═" * 62)
    print("  Apollo Bulk Create Contacts — Step 2 of 2")
    print("═" * 62)
    print(f"  Input file  : {BULK_CREATE_INPUT_FILE}")
    print(f"  Output file : {OUTPUT_FILE}")
    print(f"  Batch size  : {BATCH_SIZE}  |  Dedupe: {RUN_DEDUPE}")
    print("═" * 62 + "\n")

    # ── Load input ────────────────────────────────────────────────
    input_df = load_input_file()

    # ── Build batches ─────────────────────────────────────────────
    total = len(input_df)
    batches = [input_df.iloc[i:i+BATCH_SIZE] for i in range(0, total, BATCH_SIZE)]
    print(f"  Processing {total} contacts in {len(batches)} batch(es)...\n")

    all_results = []   # flat list of flattened contact dicts, aligned to input rows

    for b_idx, batch_df in enumerate(batches, 1):
        batch_df = batch_df.reset_index(drop=True)
        contacts_payload = [row_to_contact_payload(row) for _, row in batch_df.iterrows()]

        print(f"  Batch {b_idx:>3}/{len(batches)}  ({len(contacts_payload)} contacts)...", end=" ", flush=True)

        try:
            resp = bulk_create_batch(contacts_payload)
        except requests.exceptions.HTTPError as e:
            print(f"\n    ✗  HTTP Error: {e}  — skipping batch")
            # Fill with empty rows so index alignment is preserved
            for _ in range(len(contacts_payload)):
                all_results.append({"status": "error"})
            continue
        except Exception as e:
            print(f"\n    ✗  Error: {e}  — skipping batch")
            for _ in range(len(contacts_payload)):
                all_results.append({"status": "error"})
            continue

        new_contacts      = resp.get("contacts", [])          # newly created
        existing_contacts = resp.get("existing_contacts", []) # already in CRM

        # Apollo returns new and existing contacts unordered;
        # re-align them to our input order using position index.
        # The API mirrors back entries in the same order as sent —
        # new in 'contacts', existing in 'existing_contacts'.
        # We build a unified flat list in input order.
        created_flat  = [flatten_contact(c, "created")  for c in new_contacts]
        existing_flat = [flatten_contact(c, "existing") for c in existing_contacts]

        # Apollo splits results: new → contacts[], existing → existing_contacts[]
        # Both lists correspond positionally to the input order (created first,
        # then existing), so we just concatenate and pad to batch length.
        combined_flat = created_flat + existing_flat

        # Pad any missing entries (API may skip invalid rows)
        while len(combined_flat) < len(contacts_payload):
            combined_flat.append({"status": "no_response"})

        all_results.extend(combined_flat)

        created_count  = len(created_flat)
        existing_count = len(existing_flat)
        print(f"✓  created={created_count}  existing={existing_count}")

        if b_idx < len(batches):
            time.sleep(REQUEST_DELAY)

    if not all_results:
        print("\n  ⚠  No results returned from API. Check your API key and input file.")
        return

    results_df = pd.DataFrame(all_results)

    # ── Save output ───────────────────────────────────────────────
    print("\n" + "─" * 62)
    save_output(results_df, input_df)

    print()
    print("  ─── WORKFLOW COMPLETE ───────────────────────────────────")
    print(f"  Open '{OUTPUT_FILE}' to review enriched contacts.")
    print("  'Enriched' sheet = all contacts with email/phone appended")
    print("  'Created'  sheet = new contacts added to your Apollo CRM")
    print("  'Existing' sheet = contacts already present in Apollo CRM")
    print("─" * 62 + "\n")


if __name__ == "__main__":
    main()
