# HLF Reference Letter Generator

Generates Word reference letters for Hinduja Leyland Finance (HLF) arbitration
cases by reading case data from an Excel file and filling a Jinja2 Word template.

Two modes: **Web App** (drag-and-drop UI) or **CLI** (script / automation).

---

## Quick Start

### 1 — Install dependencies

```bash
pip install -r requirements.txt
```

### 2a — Web App (recommended)

```bash
python app.py
```

Open **http://localhost:5000**, upload your Excel file, configure the options,
and click **Generate**. When processing finishes you can download:

- **ZIP** — one individual `.docx` per case + the combined file inside
- **Combined `.docx`** — all letters merged into one Word document

### 2b — CLI

```bash
# Default (Lot 5 settings)
python run_cli.py

# Custom lot
python run_cli.py --excel "Lot 6 HLF Cases.xlsx" --ref-start 262 --lot Lot6
```

All options:

| Flag | Default | Description |
|---|---|---|
| `--excel` | `Lot 5 HLF Cases.xlsx` | Path to the Excel source |
| `--sheet` | `DataFormatted` | Sheet name |
| `--template` | `letter_templates/Reference_Letter_Template.docx` | Template path |
| `--output` | `Generated_Reference_Letters` | Output folder |
| `--ref-start` | `129` | First counter for `HLF/SNS/REF/2026/{n}` |
| `--lot` | `Lot5` | Label used in output filenames |

---

## Project Structure

```
hlf_generator/
├── app.py                          # Flask web app
├── run_cli.py                      # CLI entry point
├── generator.py                    # Core generation logic (shared)
├── requirements.txt
│
├── letter_templates/
│   └── Reference_Letter_Template.docx   # Jinja2 Word template
│
├── templates/                      # Flask HTML templates
│   ├── index.html                  # Upload form
│   └── status.html                 # Live progress + download
│
├── uploads/                        # Temp upload storage (auto-cleared)
└── outputs/                        # Generated ZIPs and combined docs
```

---

## Excel → Template Mapping

Sheet name: **DataFormatted**

| Excel Column | Template Variable | Notes |
|---|---|---|
| Contract No | `contract_no` | Also used in filename; blank rows are skipped |
| Reference Letter Date | `date` | Formatted `dd.MM.yyyy` |
| Contract Date | `agreement_date` | Formatted `dd.MM.yyyy` |
| Vehicle Details / Product Model | `asset_description` | Falls back to Product Model |
| Vehicle No + Engine No + Chassis No | `reg_chassis_phrase` | `"Reg. No. X \| Engine No. X \| Chassis No. X"` |
| Borrower Name | `borrower_name` | |
| Borrower Address | `borrower_address` | Multi-line |
| Co-Borrower Name | `co_borrower_name` | Empty if `HasCoBorrower = N` |
| Co-Borrower Address | `co_borrower_address` | Empty if `HasCoBorrower = N` |
| Guarantor Name | `guarantor_name` | Empty if `HasGuarantor = N` |
| Guaranator Address | `guarantor_address` | ⚠ Typo in header — matched exactly |
| LRN Ref No | `lrn_ref_no` | |
| LRN Date | `lrn_date` | Formatted `dd.MM.yyyy` |
| Claim Date | `claim_date` | Formatted `dd.MM.yyyy` |
| Claim Amount | `claim_amount` | |
| Claim Amount in Words | `claim_amount_words` | |
| HasCoBorrower | — | `Y` / `N` — controls Co-Borrower block |
| HasGuarantor | — | `Y` / `N` — controls Guarantor block |
| *(auto)* | `current_ref_no` | `HLF/SNS/REF/2026/{counter}` |

---

## Template (Jinja2 Placeholders)

```
{{ current_ref_no }}    {{ date }}             {{ contract_no }}
{{ agreement_date }}    {{ asset_description }}{{ reg_chassis_phrase }}
{{ borrower_name }}     {{ borrower_address }}

{% if co_borrower_name %}
{{ co_borrower_name }}  {{ co_borrower_address }}
{% endif %}

{% if guarantor_name %}
{{ guarantor_name }}    {{ guarantor_address }}
{% endif %}

{{ lrn_ref_no }}   {{ lrn_date }}   {{ claim_date }}
{{ claim_amount }} {{ claim_amount_words }}
```

---

## Adding a New Lot

**Web App:** Just upload the new Excel file and set a new Starting Reference Number.

**CLI:**
```bash
python run_cli.py \
  --excel "Lot 6 HLF Cases.xlsx" \
  --ref-start 262 \
  --lot Lot6
```

---

## Lot Reference Ranges

| Lot | File | REF range |
|---|---|---|
| 5 (active) | Lot 5 HLF Cases.xlsx | REF/2026/129 – 261 |
| 1–4 | Lot 1–4 HLF Cases.xlsx | Earlier batches |
