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
    - Create new tab `titles` under `apollo_people_search_{TIMESTAMP}.xlsx`
    - Add your titles to be removed in this tab from md file.

**Run VB Script_1.5:** 
- Go to `apollo_people_search_{TIMESTAMP}.xlsx` > Automate > New Script > Create in Code Editor > Add this script `STEP_1point5_filter_search_results_by_titles` > Save > Run
  **OR**
- Run `TJ Ignite Filter Titles` Script Already Available there (Created + Tested + Used by Me Only 😏 huh ...😒)

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
    - If Unmapped Countries Found...
        1. Go to [tejasatjignect-Gemini-Thread](https://gemini.google.com/app/b0d2c34119340152) and Find Mapping
        2. Update `COUNTRY_REGION_MAP` under step_3's script & Re-Run🔁.

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

## Post-Credit Scenes
1. Contacts Movement
	- to One Drive
		- Copy Data from `ignite_master_merged_{TIMESTAMP}.xlsx` to One Drive Excel (Shared with Ignite Team ) - [Scrapped Master Data by Tejas - Apollo](https://jignecttechnologies-my.sharepoint.com/:x:/r/personal/piyush_jignect_tech/Documents/JigNect/Org%20Departments/Sales%20%26%20Marketing/Sales/Business%20Development/Business%20Channels/LinkedIn%20Outreach/Linken%20Ignite/Required%20Data/Scrapped%20Master%20Data%20By%20Tejas%20-%20Apollo.xlsx?d=w0c52fbac4c5244e4858d2d905cdf89e7&csf=1&web=1&e=OuIP6V)
    - Copy Data from `ignite_master_merged_{TIMESTAMP}.xlsx` to Local Excel Files
	    1. [ Scrapped Master Data By Tejas - Apollo-Utility.xlsx ]
	    2. [ ✅ Scrapped Master Data By Tejas - Apollo - 22jun2026.xlsx ]

2. Drafts Movement
	- Create New tab under [📌 ALL GOOD & MAYBE - Master - TEJAS_USE_ONLY](https://jignecttechnologies-my.sharepoint.com/:x:/r/personal/piyush_jignect_tech/Documents/JigNect/Org%20Departments/Sales%20%26%20Marketing/Sales/Business%20Development/Business%20Channels/LinkedIn%20Outreach/Linken%20Ignite/Required%20Data/%F0%9F%93%8C%20ALL%20GOOD%20%26%20MAYBE%20-%20Master%20-%20TEJAS_USE_ONLY.xlsx?d=w4dd5767871194bd8a39d047bf39d7c57&csf=1&web=1&e=DGJi5o)
	- Rename `Good Fit` & `Maybe Fit` tabs with today's date . Example : `Good Fit - 22jun2026`
		- Hide both of Them.
    - Copy Sequence Drafts from Excel [📌 ALL GOOD & MAYBE - Master - TEJAS_USE_ONLY](https://jignecttechnologies-my.sharepoint.com/:x:/r/personal/piyush_jignect_tech/Documents/JigNect/Org%20Departments/Sales%20%26%20Marketing/Sales/Business%20Development/Business%20Channels/LinkedIn%20Outreach/Linken%20Ignite/Required%20Data/%F0%9F%93%8C%20ALL%20GOOD%20%26%20MAYBE%20-%20Master%20-%20TEJAS_USE_ONLY.xlsx?d=w4dd5767871194bd8a39d047bf39d7c57&csf=1&web=1&e=DGJi5o) to [JigNect 2026 - LinkedIn - Outreach_Sequences_Master](https://jignecttechnologies-my.sharepoint.com/:x:/r/personal/piyush_jignect_tech/_layouts/15/Doc.aspx?sourcedoc=%7B91C2EF0C-8907-4010-9CD2-08372CF3AA5E%7D&file=JigNect%202026%20-%20LinkedIn%20-%20Outreach_Sequences_Master.xlsx&action=default&mobileredirect=true)

3. Copy All Files from `output` folder to `OLD_OUTPUT \ run_{CURRENT_DATE_FODER}`
   
4. Delete below generated files from `output` folder which are generated today [ FOR Future's output files in clean way ✨]
	1. apollo_people_search_{TIMESTAMP}.xlsx
	2. apollo_enriched_{TIMESTAMP}.xlsx
	3. ignite_ready_contacts_{TIMESTAMP}.xlsx
	4. ignite_master_merged_{TIMESTAMP}.xlsx
	   
5. Inform Ignite team about the New Data & Drafts Added in this [MS Teams Thread](https://teams.microsoft.com/l/message/19:e8f892342e1e489ea3b119f84fd7de12@thread.tacv2/1773989120162?tenantId=0dae2554-7767-4a66-a135-9744a257ab8f&groupId=4d978cf2-1b36-4847-97c7-a15090253ff5&parentMessageId=1773989120162&teamName=JigNect%20%7C%20Sales%20%26%20Marketing&channelName=Channel%20-%20LinkedIn%20-%20Ignite&createdTime=1773989120162&ngc=true)
   
   @⁠Channel⁠-⁠LinkedIn⁠-⁠Ignite, Here we go 🚀
- Processed 211 to 600 ( total 390 ) companies from 2025_Tier_1 ✅
    - Added Outreach Sequence Drafts for `110` Good Fit/Maybe Fit Accounts - [here](https://jignecttechnologies-my.sharepoint.com/:x:/r/personal/piyush_jignect_tech/_layouts/15/Doc.aspx?sourcedoc=%7B91C2EF0C-8907-4010-9CD2-08372CF3AA5E%7D&file=JigNect%202026%20-%20LinkedIn%20-%20Outreach_Sequences_Master.xlsx&action=default&mobileredirect=true "https://jignecttechnologies-my.sharepoint.com/:x:/r/personal/piyush_jignect_tech/_layouts/15/Doc.aspx?sourcedoc=%7B91C2EF0C-8907-4010-9CD2-08372CF3AA5E%7D&file=JigNect%202026%20-%20LinkedIn%20-%20Outreach_Sequences_Master.xlsx&action=default&mobileredirect=true")
        - Added `624` new contacts info - [here](https://jignecttechnologies-my.sharepoint.com/:x:/r/personal/piyush_jignect_tech/_layouts/15/Doc.aspx?sourcedoc=%7B0C52FBAC-4C52-44E4-858D-2D905CDF89E7%7D&file=Scrapped%20Master%20Data%20By%20Tejas%20-%20Apollo.xlsx&action=default&mobileredirect=true "https://jignecttechnologies-my.sharepoint.com/:x:/r/personal/piyush_jignect_tech/_layouts/15/Doc.aspx?sourcedoc=%7B0C52FBAC-4C52-44E4-858D-2D905CDF89E7%7D&file=Scrapped%20Master%20Data%20By%20Tejas%20-%20Apollo.xlsx&action=default&mobileredirect=true")
- Total : **7203** Contacts in our Master DB ✨
6. Finished 🎉, Have some Coffee 🍵 !