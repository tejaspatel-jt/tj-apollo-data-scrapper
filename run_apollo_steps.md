# Apollo Data Scrapper — Run Commands

Every command works from any folder. It first moves into the project root, then runs the script.

**Which block to use**

| Terminal prompt starts with | Use |
|---|---|
| `tejas@JTPC07 MINGW64` | Git Bash |
| `PS C:\...>` | PowerShell |
| `C:\...>` (no `PS`) | CMD |

---

## STEP 1 — Apollo People Search

**Git Bash**

```bash
cd "/c/Users/tejas/OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED/📌 W O R K 📌/Projects/apollo-data-scrapper" && C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_1_apollo_people_search.py"
```

**PowerShell**

```powershell
$env:PYTHONUTF8 = "1"; Set-Location "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper"; & "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_1_apollo_people_search.py"
```

**CMD**

```bat
chcp 65001 >nul && set PYTHONUTF8=1 && cd /d "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_1_apollo_people_search.py"
```

---

## STEP 2 — Bulk Match, Enrich & Bulk Create

**Git Bash**

```bash
cd "/c/Users/tejas/OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED/📌 W O R K 📌/Projects/apollo-data-scrapper" && C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py"
```

**PowerShell**

```powershell
$env:PYTHONUTF8 = "1"; Set-Location "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper"; & "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py"
```

**CMD**

```bat
chcp 65001 >nul && set PYTHONUTF8=1 && cd /d "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py"
```

---

## STEP 3 — Apollo to Ignite Format

**Git Bash**

```bash
cd "/c/Users/tejas/OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED/📌 W O R K 📌/Projects/apollo-data-scrapper" && C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_3_apollo_to_ignite_format.py"
```

**PowerShell**

```powershell
$env:PYTHONUTF8 = "1"; Set-Location "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper"; & "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_3_apollo_to_ignite_format.py"
```

**CMD**

```bat
chcp 65001 >nul && set PYTHONUTF8=1 && cd /d "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_3_apollo_to_ignite_format.py"
```

---

## STEP 4 — Apollo Ignite Master Merger

**Git Bash**

```bash
cd "/c/Users/tejas/OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED/📌 W O R K 📌/Projects/apollo-data-scrapper" && C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_4_apollo_ignite_master_merger.py"
```

**PowerShell**

```powershell
$env:PYTHONUTF8 = "1"; Set-Location "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper"; & "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_4_apollo_ignite_master_merger.py"
```

**CMD**

```bat
chcp 65001 >nul && set PYTHONUTF8=1 && cd /d "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_4_apollo_ignite_master_merger.py"
```

---

## RUN ALL 4 STEPS (stops if any step fails)

**Git Bash**

```bash
cd "/c/Users/tejas/OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED/📌 W O R K 📌/Projects/apollo-data-scrapper" && \
C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_1_apollo_people_search.py" && \
C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py" && \
C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_3_apollo_to_ignite_format.py" && \
C:/Users/tejas/AppData/Local/Programs/Python/Python314/python.exe "scripts/STEP_4_apollo_ignite_master_merger.py"
```

**PowerShell**

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

**CMD**

```bat
chcp 65001 >nul && set PYTHONUTF8=1 && cd /d "C:\Users\tejas\OneDrive - JIGNECT TECHNOLOGIES PRIVATE LIMITED\📌 W O R K 📌\Projects\apollo-data-scrapper" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_1_apollo_people_search.py" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_2_apollo_bulk_match_and_enrich_and_bulkCreate.py" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_3_apollo_to_ignite_format.py" && "C:\Users\tejas\AppData\Local\Programs\Python\Python314\python.exe" "scripts\STEP_4_apollo_ignite_master_merger.py"
```
