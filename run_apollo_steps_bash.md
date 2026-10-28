# Apollo Data Scrapper — Run Commands

All commands use absolute paths, so they work from any folder. Each block first `cd`s into the project root (same as your normal run), so any relative input/output paths inside the scripts still resolve correctly.

---

## Git Bash (MINGW64)

### Run all 4 steps (stops if any step fails)

```bash
cd "/c/Users/tejas/OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED/📌 W O R K 📌/Projects/apollo-data-scrapper" && \
C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_1_apollo_people_search.py" && \
C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py" && \
C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_3_apollo_to_ignite_format.py" && \
C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_4_apollo_ignite_master_merger.py"
```

### Individual steps

```bash
cd "/c/Users/tejas/OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED/📌 W O R K 📌/Projects/apollo-data-scrapper" && C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_1_apollo_people_search.py"
```

```bash
cd "/c/Users/tejas/OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED/📌 W O R K 📌/Projects/apollo-data-scrapper" && C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py"
```

```bash
cd "/c/Users/tejas/OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED/📌 W O R K 📌/Projects/apollo-data-scrapper" && C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_3_apollo_to_ignite_format.py"
```

```bash
cd "/c/Users/tejas/OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED/📌 W O R K 📌/Projects/apollo-data-scrapper" && C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_4_apollo_ignite_master_merger.py"
```

---

## PowerShell (Windows PowerShell 5.1 and PowerShell 7)

### Run all 4 steps (stops if any step fails)

```powershell
$env:PYTHONUTF8 = "1"
$py = "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe"
Set-Location "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper"
foreach ($s in @(
  "STEP_1_apollo_people_search.py",
  "STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py",
  "STEP_3_apollo_to_ignite_format.py",
  "STEP_4_apollo_ignite_master_merger.py"
)) {
  Write-Host "`n===== Running $s =====" -ForegroundColor Cyan
  & $py "scripts\$s"
  if ($LASTEXITCODE -ne 0) { Write-Host "FAILED: $s" -ForegroundColor Red; break }
}
```

### Individual steps

```powershell
Set-Location "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper"; & "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_1_apollo_people_search.py"
```

```powershell
Set-Location "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper"; & "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py"
```

```powershell
Set-Location "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper"; & "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_3_apollo_to_ignite_format.py"
```

```powershell
Set-Location "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper"; & "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_4_apollo_ignite_master_merger.py"
```

---

## Command Prompt (CMD)

`chcp 65001` is needed so CMD can read the 📌 emoji in the folder path.

### Run all 4 steps (stops if any step fails)

```bat
chcp 65001 >nul && set PYTHONUTF8=1 && cd /d "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_1_apollo_people_search.py" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_3_apollo_to_ignite_format.py" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_4_apollo_ignite_master_merger.py"
```

### Individual steps

```bat
chcp 65001 >nul && cd /d "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_1_apollo_people_search.py"
```

```bat
chcp 65001 >nul && cd /d "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py"
```

```bat
chcp 65001 >nul && cd /d "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_3_apollo_to_ignite_format.py"
```

```bat
chcp 65001 >nul && cd /d "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_4_apollo_ignite_master_merger.py"
```
