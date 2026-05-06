import requests
import pandas as pd
import os
import time
from datetime import datetime

# --- CONFIGURATION ---
API_KEY = "l4AkVpSXwV9fICkebUMAXw"
MASTER_DB = "master_lead_database.xlsx"
BASE_URL = "https://api.apollo.io/api/v1"

# Decision-maker titles (Keep your existing list here)
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


def load_master_db():
    """Loads or creates the central database."""
    if os.path.exists(MASTER_DB):
        return pd.read_excel(MASTER_DB)
    return pd.DataFrame(columns=[
        "apollo_person_id", "first_name", "last_name", "email", 
        "organization_name", "is_enriched", "is_saved_to_crm", "source_type"
    ])

def search_apollo(endpoint, domain, company_name):
    """Generic search function for both People and Contacts."""
    all_results = []
    page = 1
    
    # Map key names based on endpoint (people uses 'people', contacts uses 'contacts')
    data_key = "people" if "mixed_people" in endpoint else "contacts"
    
    while page <= 5: # Limit pages for safety during testing
        payload = {
            "q_organization_domains_list": [domain],
            "person_titles": PERSON_TITLES,
            "page": page,
            "per_page": 100
        }
        
        response = requests.post(
            f"{BASE_URL}/{endpoint}",
            json=payload,
            headers={"x-api-key": API_KEY, "Content-Type": "application/json"}
        )
        
        if response.status_code != 200:
            break
            
        data = response.json()
        results = data.get(data_key, [])
        all_results.extend(results)
        
        if page >= data.get("pagination", {}).get("total_pages", 1):
            break
        page += 1
        time.sleep(1.2)
        
    return all_results

def main():
    master_df = load_master_db()
    companies = [{"name": "Apollo", "domain": "apollo.io"}] # Example input
    
    new_data = []

    for co in companies:
        print(f"Searching {co['name']}...")
        
        # 1. Search for ALREADY SAVED Contacts
        existing_contacts = search_apollo("contacts/search", co['domain'], co['name'])
        for c in existing_contacts:
            new_data.append({
                "apollo_person_id": c.get("person_id") or c.get("id"),
                "first_name": c.get("first_name"),
                "last_name": c.get("last_name"),
                "email": c.get("email"),
                "organization_name": co['name'],
                "is_enriched": True if c.get("email") else False,
                "is_saved_to_crm": True,
                "source_type": "CRM_CONTACT"
            })

        # 2. Search for NEW Prospects
        new_prospects = search_apollo("mixed_people/api_search", co['domain'], co['name'])
        for p in new_prospects:
            new_data.append({
                "apollo_person_id": p.get("id"),
                "first_name": p.get("first_name"),
                "last_name": p.get("last_name"),
                "email": p.get("email"),
                "organization_name": co['name'],
                "is_enriched": False,
                "is_saved_to_crm": False,
                "source_type": "NEW_PROSPECT"
            })

    # Merge with Master DB and drop duplicates based on Apollo ID
    search_df = pd.DataFrame(new_data)
    updated_master = pd.concat([master_df, search_df]).drop_duplicates(subset=["apollo_person_id"], keep="first")
    
    # Save files
    updated_master.to_excel(MASTER_DB, index=False)
    
    # Create the regular output for Step 2 (Filtering)
    # We only want to send "NEW_PROSPECT" rows to the enrichment script
    step2_input = updated_master[updated_master["is_enriched"] == False]
    step2_input.to_excel("apollo_people_search_output.xlsx", index=False)
    
    print(f"Master Database updated. Total leads: {len(updated_master)}")
    print(f"Step 2 input file created with {len(step2_input)} leads requiring enrichment.")

if __name__ == "__main__":
    main()