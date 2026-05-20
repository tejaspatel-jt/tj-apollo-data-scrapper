# 🚀 Apollo → Ignite Workflow

## ✅ Step 0 — Add Companies

**Input File:**
- `companies_input.csv`
  
**Action:**
- Add 10–15 target companies
---
## ✅ Step 1 — Run People Search

**Run Script_1:** 
- `STEP_1_apollo_people_search.py`
  
**Output Files:**
- `apollo_people_search_{TIMESTAMP}.xlsx`
- `master_lead_database.xlsx`

---

## ✅ Step 1.5 — Filter Out Unwanted records based on titles 🔴

**Input File:**
- `titles to be removed.md`
    - Add your titles to be removed in this file.

**Run VB Script_1.5:** 
- Go to `apollo_people_search_{TIMESTAMP}.xlsx` > Automate > New Script > Create in Code Editor > Add this script `STEP_1point5_filter_search_results_by_titles` > Save > Run

---

## ✅ Step 2 — Run Enrichment + CRM Sync

**Configure Input In Script_2:**
- Update `BULK_ENRICH_INPUT_FILE` → `apollo_people_search_{TIMESTAMP}.xlsx` _(from Step 1)_

**Run Script_2:**
- `STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py`

**Output Files:**
- `apollo_enriched_{TIMESTAMP}.xlsx`
- Updated `master_lead_database.xlsx`

**Quick Verification:**
- `is_enriched = true`
- `is_crm_contact = true`
---

## ✅ Step 3 — Convert to Ignite Format

**Run Script_3:**
- `STEP_3_apollo_to_ignite_format.py`

**Output File:**
- `ignite_ready_contacts_{TIMESTAMP}.xlsx`
---

## ✅ Step 4 — Merge Into Apollo Ignite Master DB Sheet

**Configure Input In Script_4:**
- Update `SOURCE_FILE` → `ignite_ready_contacts_{TIMESTAMP}.xlsx` _(from Step 4)_

**Run Script_4:**
- `STEP_4_apollo_ignite_master_merger.py`

**Final Output File:** 🎉
- `ignite_master_merged_{TIMESTAMP}.xlsx` 
