# IndiaPost Tracker Skill

You are executing the `/indiapost` skill. Follow these instructions precisely.

## What this skill does

Fetches India Post tracking data for articles listed in an Excel/CSV file and produces one or more outputs:

| Output | Description |
|--------|-------------|
| **report** | Colour-coded Excel sheet — last event, delivery outcome per article |
| **histories** | One PDF per article with full event timeline → `History_PDFs/` folder |
| **receipts** | One JSON file per article with booking + event data → `Receipts_JSON/` folder |
| **all** | All three of the above |

---

## Step 1 — Determine what the user wants

Parse the user's message (or `$ARGUMENTS`) for:

1. **Input file** — look for a path ending in `.xlsx`, `.xls`, or `.csv`. If none provided, check if they said "usual file" or similar and use `C:/Users/SaibaskarP/Downloads/sample-bulk-track-02-2026 (4).xlsx` as default. If still unclear, ask.
2. **Output type** — look for keywords: `report`, `history`/`histories`, `receipt`/`receipts`, `all`, `everything`. Default to `all` if not specified.

---

## Step 2 — Confirm and run

Tell the user:
- Which file you'll process
- Which outputs you'll generate

Then run the appropriate Python scripts from the project directory using Bash. Always activate the venv first:

```
cd "C:/Users/SaibaskarP/HLF_Reference_Letter_Generator"
source .venv/Scripts/activate
```

### For **report**:
```bash
python indiapost_report.py "<INPUT_FILE>"
```
The output Excel is saved next to the input file as `IndiaPost_Report_<timestamp>.xlsx`.

### For **histories**:
```bash
python indiapost_history_pdf.py "<INPUT_FILE>"
```
PDFs are saved to `History_PDFs/` next to the input file.

### For **receipts**:
```bash
python indiapost_receipts.py "<INPUT_FILE>"
```
Individual JSONs saved to `Receipts_JSON/<ArticleNumber>.json` and a combined `Receipts_JSON/receipts_all.json`, both next to the input file.

### For **all**:
Run all three sequentially.

---

## Step 3 — Report results

After each script completes, report:
- Number of articles processed
- Output file/folder path(s)
- Any errors encountered (show last 20 lines of traceback if script fails)
- Summary counts if available (Delivered / Returned / In Transit / No Data)

If a script fails, do NOT retry blindly. Read the error, diagnose the cause, fix the relevant script, and re-run.

---

## Error handling

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| `ModuleNotFoundError` | venv not activated | Activate venv first |
| `401 Unauthorized` | Token expired | Check `.env` credentials |
| `0 articles loaded` | Wrong column name in Excel | Run `python -c "import openpyxl; wb=openpyxl.load_workbook('FILE'); ws=wb.active; print([c.value for c in ws[1]])"` to inspect headers |
| PDF generation fails | fpdf2 not installed | `pip install fpdf2` |
| Empty JSON files | API returned null | Retry individually; log which IDs failed |

---

## Notes

- Credentials are in `.env` as `INDIAPOST_ID` and `INDIAPOST_SECRET`
- Production base URL: `https://app.indiapost.gov.in/beextcustomer/v1`
- Batch size is 20 articles per API call to avoid silent API failures
- The API returns **full event history** per article — useful for detecting return journey even before final delivery
- `History_PDFs/` and `Receipts_JSON/` are created next to the input file
