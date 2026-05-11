#!/usr/bin/env python3
"""
═══════════════════════════════════════════════════════════════
  Ignite Master Merger — Step 4 of 4
  (Python Port of mergeSourceSheetsIntoMaster Google Script)

  PURPOSE:
    Merge a new source Excel file (output of Script 3) INTO
    the "Scrapped Master Data By Tejas - Apollo" master Excel
    sheet, deduplicating by First Name + Last Name + Email.

  INPUT:
    - Master Excel File  : "Scrapped Master Data By Tejas - Apollo.xlsx"
    - Source Excel File  : ignite_ready_contacts_DDMMMYYYY_HHMMSS.xlsx
                           (output of Script 3 / Campaign Formatter)

  OUTPUT:
    - ignite_master_merged_DDMMMYYYY_HHMMSS.xlsx
      (combined, deduplicated, final sheet)

  MERGE LOGIC (mirrors mergeSourceSheetsIntoMaster JS):
    ✓ Master rows loaded first (they take priority by default)
    ✓ Source rows checked against master using smart dedup:
        A. Exact match  → First Name + Last Name + Email (skip/overwrite)
        B. Name match   → Same name but different/missing email
                          → Keep the row with MORE filled columns
        C. No match     → Genuinely new person, appended
    ✓ Rows with ALL dedup columns empty → always kept (not skipped)
    ✓ Column aliases supported (e.g. "Company LinkedIn Url" → "Company LinkedIn")
    ✓ Output column order driven by config (not source file order)
    ✓ "No" serial column re-numbered from 1 in final output
    ✓ Master file is NEVER modified — output is always a NEW file

  FEATURES:
    ✓ Smart dedup (exact + name-only match with completeness scoring)
    ✓ Column alias resolution (case-insensitive)
    ✓ Configurable: keep first vs last occurrence
    ✓ Blank dedup key rows always kept
    ✓ Detailed console summary (added / skipped / name-merged)
═══════════════════════════════════════════════════════════════
"""

import pandas as pd
import os
from datetime import datetime

# ─── PATH SETUP ───────────────────────────────────────────────
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR  = os.path.join(SCRIPT_DIR, "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

TIMESTAMP = datetime.now().strftime("%d%b%Y_%H%M%S").lower()


# ╔══════════════════════════════════════════════════════════════╗
# ║                      CONFIGURATION                          ║
# ╚══════════════════════════════════════════════════════════════╝

# ── 1. INPUT FILES ────────────────────────────────────────────
# Master Excel file — this is your "Scrapped Master Data By Tejas - Apollo" sheet.
# It is NEVER modified. The merged result always goes to a new output file.
# MASTER_FILE = "Scrapped Master Data By Tejas - Apollo_sample.xlsx"
MASTER_FILE = "Scrapped Master Data By Tejas - Apollo.xlsx"

# The sheet tab name inside the master Excel file that holds the data.
# MASTER_FILE_SHEET_NAME = "CombinedData - Master"   # <- change if your tab has a different name
MASTER_FILE_SHEET_NAME = "Sheet1"   # <- change if your tab has a different name

# Source file — output of Script 3 (Campaign Formatter / ignite_ready_contacts_*.xlsx)
# SOURCE_FILE = "ignite_ready_contacts_latest.xlsx"   # <- change to your actual file name
SOURCE_FILE = "ignite_ready_contacts_11may2026_120822.xlsx"   # <- change to your actual file name

# The sheet tab name inside the source Excel file.
SOURCE_FILE_SHEET_NAME = "Sheet1"   # <- change if needed (usually Sheet1 or the only tab)

# ── 2. OUTPUT FILE ────────────────────────────────────────────
OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    f"ignite_master_merged_{TIMESTAMP}.xlsx"
)

# ── 3. COLUMN HEADERS & ALIASES ───────────────────────────────
#
# masterColumn : The column name that appears in the OUTPUT.
#                - If it exists in master  -> data pulled from master.
#                - If it does NOT exist    -> blank column created;
#                  data can still be pulled from source via aliases.
#
# aliases      : Alternative column names to look for in the SOURCE file.
#                First match wins. Comparison is case-insensitive.
#                If omitted, masterColumn name itself is tried.
#
# Example:
#   { "masterColumn": "Company LinkedIn",
#     "aliases": ["Company Linkedin Url", "Company LinkedIn URL"] }
#
COLUMN_CONFIG = [
    { "masterColumn": "No"                                                               },
    { "masterColumn": "Industry"                                                         },
    { "masterColumn": "Region"                                                           },
    { "masterColumn": "Country"                                                          },
    { "masterColumn": "Company Name",      "aliases": ["Company Name for Emails"]        },
    { "masterColumn": "Company LinkedIn",  "aliases": ["Company Linkedin Url",
                                                        "Company LinkedIn URL"]           },
    { "masterColumn": "Website"                                                          },
    { "masterColumn": "First Name"                                                       },
    { "masterColumn": "Last Name"                                                        },
    { "masterColumn": "Designation",       "aliases": ["Title"]                          },
    { "masterColumn": "Person LinkedIn",   "aliases": ["Person LinkedIn URL",
                                                        "Person Linkedin Url"]            },
    { "masterColumn": "Like"                                                             },
    { "masterColumn": "Comment"                                                          },
    { "masterColumn": "Connection Sent"                                                  },
    { "masterColumn": "M1"                                                               },
    { "masterColumn": "M2"                                                               },
    { "masterColumn": "M3"                                                               },
    { "masterColumn": "Email"                                                            },
    { "masterColumn": "Email Sent"                                                       },
    { "masterColumn": "Status"                                                           },
    { "masterColumn": "Notes"                                                            },

    # ── Add new columns below if needed ──
    # Example: a column that doesn't exist in master yet (will be blank):
    # { "masterColumn": "Pipeline Stage" },
    #
    # Example: a column with a different name in source:
    # { "masterColumn": "Assigned To", "aliases": ["Owner", "Assigned To"] },
]

# ── 4. DEDUPLICATION COLUMNS ──────────────────────────────────
# Rows are considered duplicates when ALL these columns match another row.
# Hard stop if any of these are missing from the master file.
DEDUPE_COLUMNS = ["First Name", "Last Name", "Email"]

# ── 5. KEEP UNMENTIONED MASTER COLUMNS ───────────────────────
# True  -> Master columns NOT listed in COLUMN_CONFIG are kept at the END of output.
# False -> Only columns defined in COLUMN_CONFIG appear in output.
KEEP_UNMENTIONED_COLUMNS = False

# ── 6. KEEP FIRST OR LAST OCCURRENCE ─────────────────────────
# True  -> First occurrence wins (master rows have priority — recommended).
# False -> Last occurrence wins (source rows overwrite master — use for fresher data).
KEEP_FIRST_OCCURRENCE = True

# ── 7. SKIP BLANK DEDUP KEYS ─────────────────────────────────
# True  -> Rows where ALL dedup columns are empty are always kept (never skipped).
# False -> Blank rows treated like any other row and deduplicated normally.
SKIP_BLANK_DEDUPE_KEYS = True

# ── 8. RE-NUMBER "No" COLUMN IN OUTPUT ───────────────────────
# True  -> "No" column is re-numbered 1, 2, 3, ... in the final merged file.
# False -> "No" values from master/source are kept as-is.
RENUMBER_NO_COLUMN = True


# ╔══════════════════════════════════════════════════════════════╗
# ║                    HELPER FUNCTIONS                          ║
# ╚══════════════════════════════════════════════════════════════╝

def _safe_str(val) -> str:
    """Return a clean string; treat NaN / None as empty string."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none") else s


def _build_match_keys(row: list, indices: list) -> dict:
    """
    Build two dedup keys for a row (mirrors _msm_getMatchKeys in JS):
      strictKey -> First Name + Last Name + Email  (all 3 must match)
      nameKey   -> First Name + Last Name only     (name-only match)
    Assumes DEDUPE_COLUMNS order: [First Name, Last Name, Email]
    """
    fn = _safe_str(row[indices[0]]).lower()
    ln = _safe_str(row[indices[1]]).lower()
    em = _safe_str(row[indices[2]]).lower()
    return {
        "strictKey": f"{fn}|||{ln}|||{em}",
        "nameKey":   f"{fn}|||{ln}",
    }


def _completeness_score(row: list) -> int:
    """Count how many cells in a row are non-empty (mirrors _getRowCompletenessScore)."""
    return sum(1 for cell in row if _safe_str(cell) != "")


def _resolve_col_index(headers_lower: list, master_col: str, aliases: list) -> int:
    """
    Find index of a column in a sheet's headers.
    Search order: masterColumn name first, then aliases (case-insensitive).
    Returns -1 if not found.
    """
    search_names = [master_col] + [a for a in aliases if a != master_col]
    for name in search_names:
        try:
            return headers_lower.index(name.strip().lower())
        except ValueError:
            continue
    return -1


# ╔══════════════════════════════════════════════════════════════╗
# ║                        CORE LOGIC                            ║
# ╚══════════════════════════════════════════════════════════════╝

def load_sheet(filepath: str, sheet_name: str, label: str) -> tuple:
    """
    Load an Excel sheet. Returns (headers_original, rows_as_lists).
    headers_original : list of strings as they appear in the file.
    rows_as_lists    : list of lists (one per data row, values as strings).
    """
    print(f"  Reading {label}: {os.path.basename(filepath)}  [sheet='{sheet_name}']")
    df = pd.read_excel(filepath, sheet_name=sheet_name, dtype=str)
    df.fillna("", inplace=True)
    headers = list(df.columns)
    rows    = df.values.tolist()
    print(f"  └─ Loaded {len(rows)} rows  ({len(headers)} columns)\n")
    return headers, rows


def build_output_col_structure(master_headers: list) -> dict:
    """
    Determine the final output column list based on COLUMN_CONFIG and master headers.
    Returns a dict with keys:
      output_cols       : list of final column names (in order)
      output_to_master  : list of master column indices (or -1 if not in master)
      dedupe_indices    : indices of dedup columns within output_cols
      unmentioned_cols  : master columns not in COLUMN_CONFIG (kept if flag is True)
    """
    master_headers_lower = [h.strip().lower() for h in master_headers]
    master_col_index     = {h: i for i, h in enumerate(master_headers_lower)}

    config_cols  = [c["masterColumn"] for c in COLUMN_CONFIG]
    config_set   = {c.lower() for c in config_cols}
    unmentioned  = [h for h in master_headers
                    if h.strip().lower() not in config_set and h.strip() != ""]

    output_cols = config_cols + (unmentioned if KEEP_UNMENTIONED_COLUMNS else [])

    output_to_master = [
        master_col_index.get(col.strip().lower(), -1)
        for col in output_cols
    ]

    output_col_index = {col: i for i, col in enumerate(output_cols)}
    dedupe_indices   = [output_col_index[c] for c in DEDUPE_COLUMNS]

    # Warn about new columns (not in master — will be blank unless source has them)
    new_cols = [c["masterColumn"] for c in COLUMN_CONFIG
                if c["masterColumn"].strip().lower() not in master_col_index]
    if new_cols:
        print(f"  ℹ️  New columns (not in master, blank unless found in source): "
              f"{', '.join(new_cols)}\n")

    return {
        "output_cols":      output_cols,
        "output_to_master": output_to_master,
        "dedupe_indices":   dedupe_indices,
        "unmentioned_cols": unmentioned,
        "output_col_index": output_col_index,
    }


def load_master_rows(master_rows: list, structure: dict) -> tuple:
    """
    Convert master sheet rows into output-column order.
    Returns (output_rows, key_to_idx_map).
    """
    output_to_master = structure["output_to_master"]
    dedupe_indices   = structure["dedupe_indices"]

    key_to_idx  = {}
    output_rows = []
    int_dups    = 0

    for master_row in master_rows:
        out_row = [
            _safe_str(master_row[m_idx]) if m_idx != -1 else ""
            for m_idx in output_to_master
        ]

        keys = _build_match_keys(out_row, dedupe_indices)

        if keys["strictKey"] in key_to_idx:
            int_dups += 1
            if not KEEP_FIRST_OCCURRENCE:
                output_rows[key_to_idx[keys["strictKey"]]] = out_row
        else:
            key_to_idx[keys["strictKey"]] = len(output_rows)
            output_rows.append(out_row)

    if int_dups:
        print(f"  ⚠️  Master internal duplicates detected: {int_dups}  "
              f"({'kept first' if KEEP_FIRST_OCCURRENCE else 'kept last'})\n")

    return output_rows, key_to_idx


def merge_source_rows(
    src_headers: list,
    src_rows:    list,
    structure:   dict,
    output_rows: list,
    key_to_idx:  dict,
) -> dict:
    """
    Merge source rows into output_rows using the smart dedup logic from JS.
    Returns a stats dict.
    """
    src_headers_lower = [h.strip().lower() for h in src_headers]
    output_cols       = structure["output_cols"]
    output_col_index  = structure["output_col_index"]
    dedupe_indices    = structure["dedupe_indices"]
    unmentioned_cols  = structure["unmentioned_cols"]

    n_out = len(output_cols)
    output_to_src = [-1] * n_out

    # Resolve each output column -> source column index
    for col_def in COLUMN_CONFIG:
        mc     = col_def["masterColumn"]
        out_i  = output_col_index.get(mc, -1)
        if out_i == -1:
            continue
        aliases = col_def.get("aliases", [])
        src_i   = _resolve_col_index(src_headers_lower, mc, aliases)
        output_to_src[out_i] = src_i

    # Resolve unmentioned master columns by direct name match
    if KEEP_UNMENTIONED_COLUMNS:
        for col in unmentioned_cols:
            out_i = output_col_index.get(col, -1)
            if out_i == -1:
                continue
            try:
                src_i = src_headers_lower.index(col.strip().lower())
                output_to_src[out_i] = src_i
            except ValueError:
                pass

    added        = 0
    skipped_dup  = 0
    name_merged  = 0

    for src_row in src_rows:
        out_row = [
            _safe_str(src_row[s_idx]) if s_idx != -1 else ""
            for s_idx in output_to_src
        ]

        keys = _build_match_keys(out_row, dedupe_indices)

        # Blank dedup key -> always keep
        is_blank_key = SKIP_BLANK_DEDUPE_KEYS and all(
            _safe_str(out_row[i]) == "" for i in dedupe_indices
        )
        if is_blank_key:
            output_rows.append(out_row)
            added += 1
            continue

        # A. Exact match (First + Last + Email)
        if keys["strictKey"] in key_to_idx:
            skipped_dup += 1
            if not KEEP_FIRST_OCCURRENCE:
                output_rows[key_to_idx[keys["strictKey"]]] = out_row
            continue

        # B. Name-only match (same name, different/missing email)
        name_match_idx  = -1
        existing_strict = ""
        for s_key, idx in key_to_idx.items():
            if s_key.startswith(keys["nameKey"] + "|||"):
                name_match_idx  = idx
                existing_strict = s_key
                break

        if name_match_idx != -1:
            existing_row = output_rows[name_match_idx]
            # Keep the row with more filled columns
            if _completeness_score(out_row) > _completeness_score(existing_row):
                output_rows[name_match_idx] = out_row
                del key_to_idx[existing_strict]
                key_to_idx[keys["strictKey"]] = name_match_idx
            name_merged += 1
            continue

        # C. Completely new person
        key_to_idx[keys["strictKey"]] = len(output_rows)
        output_rows.append(out_row)
        added += 1

    return {
        "added":       added,
        "skipped_dup": skipped_dup,
        "name_merged": name_merged,
    }


# ╔══════════════════════════════════════════════════════════════╗
# ║                          MAIN                                ║
# ╚══════════════════════════════════════════════════════════════╝

def main():

    print("\n" + "═" * 62)
    print("  Ignite Master Merger — Step 4 of 4")
    print("  mergeSourceSheetsIntoMaster  (Python Port)")
    print("═" * 62)
    print(f"  Master file   : {MASTER_FILE}")
    print(f"  Source file   : {SOURCE_FILE}")
    print(f"  Output file   : {os.path.basename(OUTPUT_FILE)}")
    print(f"  Dedup columns : {', '.join(DEDUPE_COLUMNS)}")
    print(f"  Occurrence    : {'First kept (master priority)' if KEEP_FIRST_OCCURRENCE else 'Last kept (source overwrites)'}")
    print("═" * 62 + "\n")

    # STEP 1 : Load & validate master
    master_path = os.path.join(OUTPUT_DIR, MASTER_FILE)
    if not os.path.exists(master_path):
        print(f"  ❌  Master file not found: {master_path}")
        print("      -> Place the file in the 'output/' folder and re-run.\n")
        return

    master_headers, master_rows = load_sheet(master_path, MASTER_FILE_SHEET_NAME, "Master")

    master_headers_lower = [h.strip().lower() for h in master_headers]
    missing_dedupe = [c for c in DEDUPE_COLUMNS
                      if c.strip().lower() not in master_headers_lower]
    if missing_dedupe:
        print(f"  ❌  Dedup column(s) not found in master sheet: {missing_dedupe}")
        print("      -> Check DEDUPE_COLUMNS in config and re-run.\n")
        return

    # STEP 2 : Build output column structure
    structure   = build_output_col_structure(master_headers)
    output_cols = structure["output_cols"]
    print(f"  Output will have {len(output_cols)} columns.\n")

    # STEP 3 : Load master rows into output buffer
    print("  Loading master rows...")
    output_rows, key_to_idx = load_master_rows(master_rows, structure)
    print(f"  └─ {len(output_rows)} unique master rows loaded.\n")

    # STEP 4 : Load & merge source file
    source_path = os.path.join(OUTPUT_DIR, SOURCE_FILE)
    if not os.path.exists(source_path):
        print(f"  ❌  Source file not found: {source_path}")
        print("      -> Place the file in the 'output/' folder and re-run.\n")
        return

    src_headers, src_rows = load_sheet(source_path, SOURCE_FILE_SHEET_NAME, "Source")

    print("  Merging source rows...")
    stats = merge_source_rows(src_headers, src_rows, structure, output_rows, key_to_idx)

    print(f"  └─ Added (new)     : {stats['added']}")
    print(f"  └─ Skipped (exact) : {stats['skipped_dup']}")
    print(f"  └─ Name-merged     : {stats['name_merged']}  "
          f"(same name, different email -> kept more complete row)\n")

    # STEP 5 : Re-number "No" column
    if RENUMBER_NO_COLUMN and "No" in output_cols:
        no_idx = output_cols.index("No")
        for i, row in enumerate(output_rows):
            row[no_idx] = i + 1

    # STEP 6 : Write output
    final_df = pd.DataFrame(output_rows, columns=output_cols)
    final_df.to_excel(OUTPUT_FILE, index=False)

    print("─" * 62)
    print(f"  ✅  Saved -> '{os.path.basename(OUTPUT_FILE)}'")
    print(f"       ├─ Master rows    : {len(master_rows)}")
    print(f"       ├─ Source rows    : {len(src_rows)}")
    print(f"       ├─ New rows added : {stats['added']}")
    print(f"       ├─ Duplicates     : {stats['skipped_dup']}")
    print(f"       ├─ Name-merged    : {stats['name_merged']}")
    print(f"       └─ Total output   : {len(output_rows)} unique rows")
    print("═" * 62 + "\n")


if __name__ == "__main__":
    main()
