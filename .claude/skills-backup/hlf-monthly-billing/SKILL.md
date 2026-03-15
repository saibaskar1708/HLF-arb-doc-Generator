---
name: hlf-monthly-billing
description: >
  Monthly billing automation for Hinduja Leyland Finance Ltd. (HLF) arbitration cases.
  Use this skill whenever Saibaskar mentions HLF billing, monthly invoices, arbitration
  charges, generating state packages, or shares a billing Excel for Hinduja Leyland.
  The skill automates the full pipeline: reads the billing Excel, generates a covering
  letter per state (Sai & Sai Smart Solutions format), merges it with the GST bill PDF
  and a landscape Excel case detail sheet, and outputs a single combined PDF per state
  into an "HLF Invoice Packages" folder. Also detects states missing GST bills and
  auto-generates them. Trigger on phrases like: "generate HLF billing", "run monthly
  billing", "process this month's cases", "create invoice packages", "billing for April".
---

# HLF Monthly Billing — Skill Guide

This skill automates the complete monthly billing workflow for Hinduja Leyland Finance Ltd.
Each month Saibaskar uploads a fresh `billing-data.xlsx`, and this skill produces a
per-state PDF invoice package containing:

1. **Covering Letter** — Sai & Sai Smart Solutions letterhead, correct HLF state address,
   cases grouped by type/stage, GST calculation, bank details, signature space
2. **GST Bill** — the state's Tax Invoice PDF (provided by user, or auto-generated if missing)
3. **Case Detail Sheet** — landscape A4 Excel converted to PDF, no address column

---

## Step 1 — Gather inputs

Ask the user (if not already provided):

- **Billing Excel**: the `billing-data.xlsx` for this month (path or upload)
- **GST Bills folder**: path to the folder containing state PDFs (e.g. `JHARKHAND.pdf`,
  `RAJASTHAN.pdf`, etc.). Usually the same `GST Bills/` folder used last month, refreshed.
- **Bill date**: date to print on the covering letters (e.g. `14.04.2026`)
- **Billing month**: three-letter abbreviation — `APR`, `MAY`, `JUN`, etc.
- **Billing year**: four-digit year — `2026`

The output folder will be auto-named: `HLF Invoice Packages - {MONTH} {YEAR}/`
inside the same directory as the billing Excel, unless the user specifies otherwise.

---

## Step 2 — Check dependencies

Make sure these Python packages are installed before running:

```bash
pip install weasyprint pypdf num2words openpyxl pandas --break-system-packages -q
```

---

## Step 3 — Locate the script and state master

The skill ships with two files in `scripts/`:

- `build_packages.py` — the main build script (do not edit)
- `state_master.json` — all persistent HLF data: addresses, GST numbers, invoice prefixes

**Skill location:** The `.hlf-billing-skill/` folder lives in the user's `Bill 1` workspace.
If that folder has moved, search for `build_packages.py` in the workspace.

```
<workspace>/
└── .hlf-billing-skill/
    └── hlf-monthly-billing/
        └── scripts/
            ├── build_packages.py
            └── state_master.json
```

---

## Step 4 — Handle new or missing states

Before running, compare states in the billing Excel against `state_master.json`:

```python
import pandas as pd, json
df = pd.read_excel("<billing_excel>", sheet_name="Billing Data")
billing_states = set(df["State"].dropna().unique())
master_states = set(json.load(open("state_master.json"))["states"].keys())
missing = billing_states - master_states
print("States NOT in master:", missing)
```

If any states are missing from the master:
- Tell the user which states are new
- Ask for the HLF office address and GST number for each new state
- Add them to `state_master.json` before proceeding (follow the existing JSON structure)

If a state has no GST bill PDF in the GST bills folder:
- The script will auto-generate one from billing data and save it to the GST Bills folder
- Inform the user: "No GST bill found for [State] — generating automatically from billing data"

---

## Step 5 — Run the build

```bash
python3 <skill_scripts_dir>/build_packages.py \
  --billing-excel  "<path-to-billing-data.xlsx>" \
  --gst-bills-dir  "<path-to-gst-bills-folder>" \
  --output-dir     "<output-folder>" \
  --bill-date      "DD.MM.YYYY" \
  --billing-month  "MON" \
  --billing-year   YYYY
```

To process only specific states (useful for regenerating one state):
```bash
  --states-only "Jharkhand,Rajasthan"
```

To skip states with no GST bill instead of auto-generating:
```bash
  --skip-missing-gst
```

---

## Step 6 — Report results to user

After the script completes, tell the user:
- How many packages were generated
- The output folder path (with a `computer://` link)
- Any states where the GST bill was auto-generated
- Any warnings (unknown states, failed Excel→PDF conversions)

Example summary:
> 17 invoice packages saved to [HLF Invoice Packages - APR 2026](computer:///path/to/folder/).
> Note: Vidarbha GST bill was auto-generated from billing data (no PDF was found in GST Bills/).

---

## Step 7 — Updating state master data

If HLF changes an office address or GST number for a state, edit `state_master.json`
directly. Key fields per state:

| Field | Meaning |
|---|---|
| `inv_prefix` | Invoice number prefix: `ARB/HLF[gst-code]` |
| `state_suffix` | Short state code used in invoice numbers: `JH`, `RJ`, etc. |
| `hlf_address` | Array of address lines shown on covering letter and GST bill |
| `gstin` | HLF's GST number for this state |
| `state_code` | GST state code (numeric) |
| `gst_type` | `IGST` for inter-state, `CGST+SGST` for Tamil Nadu (intra-state) |
| `gst_file_pattern` | Filename stem of the GST bill PDF (without `.pdf`) |

---

## Step 8 — Adding a new state permanently

If next month's billing data contains a brand-new state:

1. Confirm the HLF office address and GST number with the user
2. Add the state to `state_master.json` following the existing pattern
3. Choose an appropriate `inv_prefix` (use HLF's GST state code), `state_suffix`, and `gst_file_pattern`
4. Re-run the build

---

## Important notes

- **Tamil Nadu** uses `CGST+SGST` (intra-state) instead of IGST — this is handled automatically
- **Vidarbha** shares Maharashtra's GST code (`27`) but has its own invoice suffix `VB`
- The covering letter date is always the date specified by the user, independent of GST bill dates
- The Excel case detail sheet deliberately omits the `Customer Address`, `Office`, and `GST No`
  columns for readability in landscape format
- All signature spaces are blank (60px) — Saibaskar signs physically after printing
