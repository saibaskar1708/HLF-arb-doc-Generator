# Notice of Hearing (NOH) — Requirements & Implementation Reference

> **Last updated:** February 2026
> **Status:** Implemented — `generate_noh.py` + Flask web interface (`app.py`)

---

## Source Data
- File: `Lot 4 HLF Cases.xlsx` (or any Excel supplied at runtime)
- Sheet: `DataFormatted` (configurable via web form)
- Current lot has **2 unique arbitrators**

---

## Page Layout — A4, Page 1 must fit on one page

| Setting | Value |
|---------|-------|
| Page size | A4 (21 cm × 29.7 cm) |
| Margins | Top 1.2 cm, Bottom 1.2 cm, Left 2.0 cm, Right 2.0 cm |
| Font | Times New Roman, 10 pt |
| Content width | 17 cm |

**Overflow handling (3 respondents):** When all three respondent slots are filled, spacing around "Vs" and the salutation is tightened (1 pt instead of 2 pt), and respondent entries are separated by a single blank line instead of two.

---

## Page 1 — Notice of Hearing

### Header (Letterhead)

Two-column borderless table (11 cm | 6 cm):

| Left (11 cm) | Right (6 cm) |
|---|---|
| **`{Arbitrator Name}`** (bold) | `{Arbitrator Address}` (left-aligned, wraps to 2–3 lines) |
| `ARBITRATOR` | |

- **Arbitrator Name** → `Arbitrator Name` (index 29)
- **Arbitrator Address** → `Arbitrator Address` (index 30) — narrow column forces natural 2–3 line wrap
- Horizontal rule drawn below header

---

### Section 1 — Tribunal & Case Identification (centered, bold)

```
BEFORE THE ARBITRATOR & ADVOCATE {Arbitrator Name}
CLAIM PETITION NO. {Case No}
```

- **Case No** → `Case No` (index 27)

---

### Section 2 — Parties

**Claimant block** (fixed text, left cell of 13.5 cm | 3.5 cm table):
```
1. M/S. Hinduja Leyland Finance Ltd.,
Rep. By Its Authorized Representative Ms. Sumana B - Corporate Legal
Having Its Corporate Office
No.27a, Developed Industrial Estate,
Guindy, Chennai-600 032.
```
- Right-aligned label (bottom-aligned cell): `.... Claimant`

---

**"Vs"** — centered

---

**Respondents** (same 13.5 cm | 3.5 cm borderless table):

| Party | Condition | Number |
|-------|-----------|--------|
| Borrower | Always present | `2.` |
| Co-Borrower | If `Co-Borrower Name` (index 18) is non-empty | `3.` |
| Guarantor | If `Guarantor Name` (index 20) is non-empty | `3.` or `4.` |

- Respondents always start from `2.` (Claimant is always `1.`)
- Right-aligned label (bottom-aligned cell): `.... Respondent` or `.... Respondents` — determined by count in code, **not** from the `Plural` Excel column (index 52)

**Address rules:**
- Each address is reflowed to a maximum of **3 lines** (~76 chars/line) via `clip_address()` to prevent page overflow
- If Co-Borrower or Guarantor address is a placeholder (`"Same as above"`, `"same"`, `"-"`, or blank), it is **replaced** with the Borrower's address automatically

**Field mappings:**
- Borrower Name/Address → index 16/17
- Co-Borrower Name/Address → index 18/19
- Guarantor Name/Address → index 20/21

---

### Section 3 — Salutation, Ref & Sub

```
Sir/Madam,

Ref: Letter dated {Arbitrator Appointment Date} from Sai & Sai Arbitration Centre – Appointment of
Arbitrator – Arbitration in the matter of dispute(s) between M/S. Hinduja Leyland Finance Ltd. vs.
{Honorific} {Borrower Name} in respect of Loan Account No. {Contract No} dated {Contract Date}.

                    Sub: Notice of Hearing
```

- **Arbitrator Appointment Date** → index 28
- **Honorific** → `Refer Borrower` (index 53) — e.g. `Mr.` / `Ms.`
- **Contract No** → index 0 · **Contract Date** → index 1
- Sub line is centered and bold

---

### Section 4 — Body Paragraph

```
Sai & Sai Arbitration Centre through its letter dated {Arbitrator Appointment Date} had nominated
and appointed me as Arbitrator to arbitrate on the disputes/claim arisen between {PARTY ORDINALS}
of you which was in furtherance to Letter dated {Reference Letter Date} from the first of you. I
hereby accept my appointment as Arbitrator through the letter dated {Arbitrator Acceptance Date}.
There are no circumstances exist that give rise to justifiable doubts as to my independence or
impartiality in resolving the dispute referred. Declaration under Section 12(1) of the Arbitration
and Conciliation Act, 1996 – as per Sixth Schedule is also annexed herewith. {CONDITIONAL LAST SENTENCE}
```

**Party ordinals** (built in code):

| Respondent count | Ordinal text |
|---|---|
| 1 (Borrower only) | `the 1st and 2nd` |
| 2 (+ Co-Borrower or Guarantor) | `the 1st, 2nd and 3rd` |
| 3 (Borrower + Co-Borrower + Guarantor) | `the 1st, 2nd, 3rd and 4th` |

**Conditional last sentence** (`Contract Status`, index 2):

| Value | Last sentence |
|-------|---------------|
| `L` or `R` | *"The Claimant has filed their Claim Statement along with the affidavit and the petition under the Arbitration and Conciliation Act 1996, which are appended herewith."* |
| Any other | *"The Claimant has filed their Claim Statement, which is appended herewith."* |

**Rule:** The phrase "Sole Arbitrator" must **never** appear — always use "Arbitrator".

- **Reference Letter Date** → index 26 · **Arbitrator Acceptance Date** → index 31

---

### Section 5 — Hearing Notice

```
Take notice that the above matter stands posted for hearing on {First Hearing Date} between
{Meeting Timings}. You have the option to appear either in person or through your authorized
representative at:
```

- **First Hearing Date** → index 37 · **Meeting Timings** → index 49

---

### Section 5a — Venue Address (static)

```
Sai and Sai Arbitration Centre,
No.2, Diwan Bahadur Shanmugam Street,
Kilpauk, Chennai- 600010,
e-mail: {Arbitrator Email}  Phone: +91 44 48557697.
```

- **Arbitrator Email** → index 50

---

### Section 6 — Video Conferencing Block

```
Alternatively, you may choose to attend the proceedings via video conferencing for which the
details are stated below

Google Meet Info  ← bold + underlined

Video call link: {Meeting Link}
```

- **Meeting Link** → index 48 — **mandatory**; row is skipped if missing

---

### Section 7 — Bottom Table (QR | Date | Signature)

Three-column borderless table (5.0 cm | 6.5 cm | 5.5 cm):

| Col 1 (left-aligned) | Col 2 (vertically centered, centered) | Col 3 (bottom-aligned, right-aligned) |
|---|---|---|
| "Or scan the QR Code below to join the meeting" | `Dated the {D}th day of {Month} {Year}` | `[signature image if provided]` |
| QR code image (3 cm, left-aligned) | | `{Arbitrator Name}` |
| | | `(ARBITRATOR)` |

- **NOH Date** → index 32 — **mandatory**; row is skipped if missing. Day gets ordinal suffix (1st, 2nd, 3rd, 9th, 11th, 21st, …)
- **QR code** generated from `Meeting Link` URL at ~3 cm width
- **Signature image** (optional): if uploaded via web interface, white/near-white background is made transparent (threshold 240), resized to 2.8 cm width, placed above name/ARBITRATOR text

---

## Page 2 — Disclosure Statement

### Title (centered, bold)
```
DISCLOSURE PROVIDED U/S.12 (1) READ WITH SIXTH SCHEDULE OF THE ARBITRATION AND CONCILIATION ACT, 1996
```

### Table (Table Grid, 3 cols: 1.5 cm | 9.0 cm | 6.5 cm)

| Sr. No. | Particulars | Details |
|---|---|---|
| 1 | Name of the Arbitrator | `{Arbitrator Name}` (index 29) |
| 2 | Contact Details | `{Arbitrator Address}` (index 30) |
| 3 | Prior experience (including experience with Arbitrations) | `{Arbitrator Experience}` (index 71) |
| 4 | Number of on-going arbitrations | *(blank)* |
| 5 | Circumstances disclosing any past or present relationship… | Fixed: *"No vested interest with any of the parties"* |
| 6 | Circumstances which are likely to affect the ability to devote sufficient time… | Fixed: *"No adverse circumstances to affect the ability to devote sufficient time in finishing the proceedings as stipulated."* |

### Page 2 Signature Block (bottom-right)

```
[signature image if provided — 2.8 cm wide]
{Arbitrator Name}
ARBITRATOR
```

Same signature image as Page 1 is reused — `make_sig_transparent()` is called again with the same path.

---

## Output Files

### Per-document filename
```
NOH_{safe_case_no}_{safe_contract_no}.docx
```
- Example: `NOH_HLF_SNS_ARB_2026_01_HLF_SNS_2026_001.docx`
- Case No leads the filename so files sort correctly in file explorer
- `/` and `\` replaced with `_`

### Batch outputs (web interface)
| File | Description |
|------|-------------|
| `HLF_NOH_{job_id[:8]}.zip` | ZIP of all individual `.docx` files |
| `HLF_NOH_Combined_{job_id[:8]}.docx` | All NOH docs merged into one file (page-break separated) |

Both are saved to the `NOH_Output/` folder and downloadable from the status page.

---

## Batch Processing & Mandatory Field Validation

- Processes all data rows in the configured sheet (skips fully empty rows)
- **Mandatory fields** — row is skipped (no document generated) if either is missing:
  - `NOH Date` (index 32)
  - `Meeting Link` (index 48)
- Skipped rows are reported in the web UI with Contract No, Borrower Name, and reason

---

## Web Interface (`app.py` + Flask)

### Endpoints

| Route | Method | Purpose |
|-------|--------|---------|
| `/` | GET | Upload form (index.html) |
| `/generate` | POST | Accept Excel + config, start background job, redirect to status |
| `/status/<job_id>` | GET | Status page (status.html) with live progress |
| `/api/status/<job_id>` | GET | JSON job status polled by browser every 1 s |
| `/api/scan-arbitrators` | POST | Accept Excel file, return unique arbitrator names (for UI auto-fill) |
| `/download/<job_id>/zip` | GET | Serve ZIP file |
| `/download/<job_id>/combined` | GET | Serve combined DOCX |

### NOH-specific form fields (index.html)

- **Document Type** dropdown: `reference_letter` / `noh`
- **Sheet Name**: default `DataFormatted`
- **Arbitrator Signatures** (dynamic slots, NOH only):
  - On Excel file selection, `scanArbitrators()` auto-POSTs to `/api/scan-arbitrators` and pre-fills name inputs
  - Each slot: Arbitrator Name (must match Excel exactly, case-insensitive) + Signature Image (PNG/JPG)
  - "+ Add arbitrator signature" button for additional slots

### Signature matching

Signatures are matched to documents by normalising both the uploaded name and the row's `Arbitrator Name` value to lowercase-stripped strings. The correct signature is affixed per document automatically.

---

## Required Libraries

```
pip install python-docx qrcode pillow openpyxl flask werkzeug
```

---

## Future Roadmap

- **Database backend**: Replace Excel input with a structured DB (PostgreSQL / SQLite) — rows pre-validated, no more manual column mapping
- **Hosted web app**: Deploy on a server with HTTPS
- **Authentication**: User login, role-based access (e.g. admin vs. read-only)
- **Audit trail**: Log which user generated which batch and when
- **Email delivery**: Optionally email generated NOH documents directly to respondents

---

## Notes

- "Sole Arbitrator" must **never** appear in any document — always "Arbitrator"
- `Plural` column (index 52) in Excel is ignored — singular/plural determined in code
- Page 1 must fit on a single A4 page; tight layout enforced via margins, font size, and dynamic spacing
- Signature images: white/near-white pixels (R ≥ 240, G ≥ 240, B ≥ 240) are made fully transparent before embedding
