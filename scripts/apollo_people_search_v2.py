#!/usr/bin/env python3
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

# ╔══════════════════════════════════════════════════════════════╗
# ║                     CONFIGURATION                           ║
# ╚══════════════════════════════════════════════════════════════╝

API_KEY = "l4AkVpSXwV9fICkebUMAXw" 
# ADDED: Master Database Path
MASTER_DB_PATH = os.path.join(OUTPUT_DIR, "master_lead_database.xlsx")

TIMESTAMP = datetime.now().strftime("%d%b%Y_%H%M%S").lower()
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"apollo_people_search_{TIMESTAMP}.xlsx")

REQUEST_DELAY = 1.5 

# Target Decision Makers
PERSON_TITLES = [
    "ceo", "founder", "co-founder", "partner", "owner", "managing director",
    "cto", "vp of engineering", "head of engineering", "engineering manager",
    "vp of product", "head of product", "product manager",
    "vp of marketing", "chief marketing officer", "head of marketing",
    "vp of sales", "chief revenue officer", "head of sales"
]

# ─── EXISTING HELPER FUNCTIONS ──────────────────────────────────────────────

def get_companies():
    csv_path = os.path.join(INPUTS_DIR, "companies_input.csv")
    if os.path.exists(csv_path):
        df = pd.read_csv(csv_path)
        return df.to_dict('records')
    return [{"name": "Apollo", "domain": "apollo.io"}]

def flatten_person(p, company_name, domain):
    """ Your original flattening logic """
    return {
        "apollo_person_id": p.get("id"),
        "First Name": p.get("first_name"),
        "Last Name": p.get("last_name"),
        "Name": p.get("name"),
        "Title": p.get("title"),
        "Company Name": company_name,
        "Company Domain": domain,
        "Email": p.get("email"),
        "Email Status": p.get("email_status"),
        "Person Linkedin Url": p.get("linkedin_url"),
        "City": p.get("city"),
        "State": p.get("state"),
        "Country": p.get("country"),
        # Added status flags for the Master DB
        "is_enriched": True if p.get("email") else False,
        "is_saved_to_crm": False # Default, updated in main
    }

# ─── NEW: SEARCH FUNCTIONS FOR PROSPECTS & SAVED CONTACTS ────────────────────

def search_apollo_api(endpoint, domain):
    """
    Handles both 'mixed_people/api_search' (New) 
    and 'contacts/search' (Already Saved)
    """
    all_results = []
    page = 1
    
    # API key for the results list is different for each endpoint
    data_key = "people" if "mixed_people" in endpoint else "contacts"
    
    while True:
        url = f"https://api.apollo.io/api/v1/{endpoint}"
        payload = {
            "api_key": API_KEY,
            "q_organization_domains_list": [domain],
            "person_titles": PERSON_TITLES,
            "page": page
        }
        
        try:
            response = requests.post(url, json=payload, timeout=20)
            if response.status_code != 200:
                print(f"         Error {response.status_code} on {endpoint}")
                break
                
            data = response.json()
            results = data.get(data_key, [])
            all_results.extend(results)
            
            # Pagination check
            total_pages = data.get("pagination", {}).get("total_pages", 1)
            if page >= total_pages or page >= 5: # Safety cap at 5 pages
                break
            page += 1
            time.sleep(REQUEST_DELAY)
        except Exception as e:
            print(f"         Request failed: {e}")
            break
            
    return all_results

# ─── MAIN EXECUTION ─────────────────────────────────────────────────────────

def main():
    companies = get_companies()
    all_rows = []

    print(f"\n🚀 Starting Search for {len(companies)} companies...")

    for idx, company in enumerate(companies, 1):
        name = company["name"]
        domain = company["domain"]
        print(f"[{idx}/{len(companies)}] Processing: {name}")

        # STEP 1A: Search for Already Saved People (Contacts API)
        saved_contacts = search_apollo_api("contacts/search", domain)
        for c in saved_contacts:
            row = flatten_person(c, name, domain)
            # Override ID if using contacts endpoint which might return 'person_id'
            row["apollo_person_id"] = c.get("person_id") or c.get("id")
            row["is_saved_to_crm"] = True
            all_rows.append(row)

        # STEP 1B: Search for New People (Mixed People API)
        new_people = search_apollo_api("mixed_people/api_search", domain)
        for p in new_people:
            row = flatten_person(p, name, domain)
            row["is_saved_to_crm"] = False
            all_rows.append(row)

    if not all_rows:
        print("No results found.")
        return

    # Create current search DataFrame
    current_df = pd.DataFrame(all_rows)
    current_df.drop_duplicates(subset=["apollo_person_id"], keep="first", inplace=True)

    # ─── UPDATE MASTER DATABASE ──────────────────────────────────────────────
    if os.path.exists(MASTER_DB_PATH):
        master_df = pd.read_excel(MASTER_DB_PATH)
        # Append new results to master, then drop duplicates keeping the NEWEST status
        updated_master = pd.concat([current_df, master_df]).drop_duplicates(
            subset=["apollo_person_id"], keep="first"
        )
    else:
        updated_master = current_df

    # Save the Centralized Database
    updated_master.to_excel(MASTER_DB_PATH, index=False)
    print(f"✅ Master Database Updated: {MASTER_DB_PATH}")

    # Save this specific run's output (to be used in Step 2)
    current_df.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Run Output Saved: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()