# 📊 Monthly Shareholding of Directors & Sponsors
### United Commercial Bank PLC (UCB) — Regulatory Compliance Workbook

> 🎯 **Purpose in one line:** Every month, paste the latest share register into `rawdata` — the entire workbook recalculates automatically and produces ready-to-submit regulatory reports for BSEC, DSE, and CSE, plus individual director shareholding certificates.

---

## 🏦 File Information

| 🏷️ Attribute | 📋 Details |
|---|---|
| 📂 File Name | `Monthly_Shareholding_of_Directors_Sponsors.xlsm` |
| 🏛️ Issuer | United Commercial Bank PLC |
| 📁 Format | Excel Macro-Enabled Workbook (`.xlsm`) |
| 📡 Regulatory Bodies | BSEC · DSE · CSE |
| 💰 Total Paid-up Capital | BDT 14,765,484,750 |
| 🔢 Total Securities | 1,476,548,475 shares |

---

## 🧠 How the Whole System Works — Big Picture

This workbook follows a **single-input, multi-output** design. There is only **one place** where data is entered manually each month — the `rawdata` sheet. Everything else in the file is formulas that read from `rawdata` and produce outputs automatically.

```
                    ┌─────────────────────────────┐
                    │  👤 USER PASTES NEW DATA    │
                    │  into 'rawdata' each month  │
                    └──────────────┬──────────────┘
                                   │
                    ┌──────────────▼──────────────┐
                    │  📊 Categorywise Summary    │  ← Pivot table aggregates
                    │  (PivotTable engine)         │     all shares by category
                    └──────┬───────────────┬───────┘
                           │               │
           ┌───────────────▼──┐     ┌──────▼──────────────┐
           │ 📋 Page 2 & 3    │     │ 📄 For Web / raw   │
           │ (BSEC/DSE/CSE    │     │ data 2 (website     │
           │  official tables)│     │  disclosure format) │
           └───────────────┬──┘     └─────────────────────┘
                           │
           ┌───────────────▼──────────────────────────────┐
           │  📜 page 1 - BSEC / page 1 - DSE / page 1 -  │
           │  CSE  (cover letters, auto-dated, auto-ref'd) │
           └───────────────────────────────────────────────┘

           ┌──────────────────────────────────────────────┐
           │  👨‍👩‍👧 As per Family sheets                      │
           │  (director-wise breakdown, buy/sell detection)│
           └──────────────────────────────────────────────┘

           ┌──────────────────────────────────────────────┐
           │  📜 Shareholding Certificates 1 & 2          │
           │  (auto-generated text with BDT/USD value)    │
           └──────────────────────────────────────────────┘
```

---

## 🗂️ Sheet-by-Sheet Breakdown — What Each Sheet Does & How

### 1. 📥 `rawdata` — The Single Source of Truth

**What it is:** A structured Excel Table containing the full share register snapshot for the current month.

**How it works:**
- Each row = one BO (Beneficiary Owner) account with columns: `FOLIO`, `BOID`, `NAME`, `PAPER`, `ELECTRONIC`, `SUSPENSE`, `TOTAL`, `POSITION` (category tag), `PERCENTAGE`.
- A helper column `Combined` concatenates `BOID + FOLIO` to create a unique key for lookups — this is how other sheets find the right row without ambiguity, even when the same person has multiple BO accounts.
- The `POSITION` column (e.g. `Sponsor`, `Director`, `SPONSOR FAMILY`) is how the PivotTable in `Categorywise Summary` groups and counts shares.
- **The workflow:** When a new month starts, the user replaces the rows in this table with the fresh data from the share depository. Every formula in every other sheet refreshes instantly.

---

### 2. 🔍 `BOID Wise Search` — Quick Lookup Tool

**What it is:** A one-cell search panel where you type a BO ID and instantly see all share holdings for that person.

**How it works:**
- Type any BOID in cell B1 — a PivotTable filters `rawdata` and shows that person's Paper, Electronic, Suspense, and Total shares in seconds.
- Useful for quickly verifying a specific director or shareholder before submitting reports.

---

### 3. 🗃️ `raw data 2` — Category Percentage Summary (Simple View)

**What it is:** A compact summary of shareholding percentages by investor category, plus company-level metadata.

**How it works:**
- Lists the 15 investor categories (Sponsor, Directors, Local Investor, Foreign Institution, etc.) with their current percentage.
- Also stores the total paid-up capital in BDT and the share count formula (`=H6/10`) so any certificate or report can pull from one place.
- Acts as a quick reference for anyone who wants to see the headline numbers without opening the main pivot.

---

### 4. 📊 `Categorywise Summary` — The Central Calculation Engine

**What it is:** The mathematical heart of the workbook — a PivotTable over `rawdata` combined with BSEC-format category calculations.

**How it works — two parts:**

**Part A — PivotTable:** Automatically groups all rows in `rawdata` by the `POSITION` tag and sums up `PAPER`, `ELECTRONIC`, `SUSPENSE`, and `TOTAL` shares for each category (Sponsor, Directors, Sponsor Family, Govt, Local/Foreign Institution, etc.). Every other sheet uses `GETPIVOTDATA()` to pull live values from this pivot — meaning if `rawdata` changes, everything downstream updates in one refresh.

**Part B — BSEC 5-Category Mapping:** The raw categories above are then rolled up into the 5 regulatory buckets required by BSEC/DSE:
- **A** — Sponsors/Promoters & Directors → sum of Sponsor + Director pivot rows
- **B** — Govt → Government category
- **C** — Institute → Local Institution
- **D** — Foreign → Foreign Institution
- **E** — Public → everything remaining (computed as `1 - A - B - C - D`)

**Part C — Range-wise Distribution Table:** A separate table breaks the full shareholder list into share-count bands (1–500, 501–5000, … above 1,000,000) using `COUNTIFS` and `SUMIFS` against `rawdata`. This is the table that appears in the Page 2 & 3 regulatory submission under "distribution of securities."

---

### 5. 📅 `Monthwise Category Summary` — Historical Archive

**What it is:** A running log of every past month's category-wise share counts.

**How it works:**
- Each time the monthly report is finalised, the current month's data is copied-and-pasted-as-values into a new block of rows here (tagged with the month label e.g. `Dec'23`, `Jan'24`).
- This builds a permanent historical record so you can go back and see exactly what was reported to BSEC in any previous month without opening old files.
- Also used to verify that the current month's numbers are reasonable compared to last month.

---

### 6. 📄 `page 1 - BSEC` / `page 1 - DSE` / `page 1 - CSE` — Regulatory Cover Letters

**What they are:** Three identical-in-structure cover letter templates, each addressed to a different regulator.

**How they work:**
- The letter reference number is auto-generated: `=CONCATENATE("NO.UCB/HO/CS/F.27(V)/   /", YEAR(TODAY()))` — so it always carries the correct year.
- The date is `=TODAY()` — always current when you open the file.
- The body text cites the specific regulation: *"Pursuant to regulation 35(2) of the [Exchange] (Listing) Regulations 2015..."*
- The shareholding percentages and table in the lower section pull live from `Categorywise Summary` via `GETPIVOTDATA()`.
- **You never type in these sheets.** Just refresh `rawdata`, and all three letters are print-ready.

---

### 7. 📋 `Page 2 & 3` / `Page 2 & 3 (Previous)` — Official Director Shareholding Tables

**What they are:** The formal schedule of director/sponsor shareholding in the exact BSEC-prescribed format for monthly submission.

**How they work:**
- Each director row uses `=SUMIF([1]Sheet1!$A:$A, $B7, [1]Sheet1!$G:$G)` — this references the previous month's file (linked externally as `[1]Sheet1`) for the "Previous Month (B)" column automatically.
- The "Reporting Month (A)" column pulls from the live `rawdata` via `GETPIVOTDATA`.
- The "Change" column `= A - B` auto-calculates how many shares each director gained or lost.
- The "Reason for Change" column is filled manually (e.g. "5% Stock Dividend Added", "Purchased").
- `Page 2 & 3 (Previous)` is a saved snapshot of the prior month's version, kept for the cross-reference link above.

---

### 8. 📅 `date` — Shared Date Reference Cell

**What it is:** A tiny helper sheet with one purpose — to hold the reporting date as the single authoritative source.

**How it works:**
- Cell B3 is set to the report date (pulled from `'page 1 - CSE'!F27`).
- All other sheets that need the date — certificates, family sheets, cover letters — reference `date!B3` rather than calling `TODAY()` independently. This ensures every sheet shows the exact same date and avoids inconsistency if the file is opened on different days.
- Sub-cells extract the day, month name, and year separately for use in formatted text strings.

---

### 9. 🌐 `For Web` / `For Web (2)` — Website Disclosure Format

**What they are:** Clean, minimal summaries formatted for upload to the DSE/CSE public disclosure portal.

**How they work:**
- `For Web` pulls percentages directly from `page 1 - DSE` divided by 100 to get decimals.
- `For Web (2)` pulls from `Categorywise Summary` directly and has minor manual adjustments (e.g. `+0.24%` to C, `-0.24%` from E) to reconcile rounding.
- Both display the date dynamically: `=CONCATENATE("As on ", TEXT(date, "MMMM DD, YYYY"))`.
- These are copy-pasted into the exchange's online portal each month.

---

### 10. 🏆 `Top 20` — Largest 20 Shareholders

**What it is:** A ranked list of the 20 biggest shareholders by total shares held, required as part of the regulatory submission.

**How it works:**
- Director-representative companies (Volkart, Legendary Assets, Ardent, etc.) use `=SUMIF(rawdata[NAME], "COMPANY NAME", rawdata[TOTAL])` to aggregate holdings across multiple BO accounts.
- Individual large shareholders use `=VLOOKUP(BOID, rawdata, 6, FALSE)` to pull directly from `rawdata`.
- The list is semi-manual — the user re-ranks if the order changes — but all share counts update automatically.
- A secondary calculation block estimates post-5%-dividend projections: `=D4+(D4*0.05)`.

---

### 11. 👨‍👩‍👧 `As per Family - Faisal Vai` — Director Family Tracking (Current Month)

**What it is:** A detailed breakdown of each director's shareholding grouped by family, showing every individual BO account they or their family members hold.

**How it works:**
- Each row = one BO account of a director or their family member.
- The current month's shares are looked up live from `rawdata` using `INDEX/MATCH` on the `Combined` key (`BOID + FOLIO`): `=IFERROR(INDEX(rawdata[ELECTRONIC], MATCH($F3, rawdata[Combined], 0)), 0)`.
- The previous month's figure is stored manually in the `Share no of Previous Month` column.
- The **`Difference`** column = `Current - Previous` and a smart `Status` formula automatically classifies the change:
  - If difference ≈ 5% of previous month → labels it `"5% stock"` (bonus dividend)
  - If previous > current → `"Sold X Share"`
  - If current > previous → `"Bought X Share"`
  - If no change → `0`
- This makes it immediately visible whether any director or family member bought, sold, or just received a stock dividend.

---

### 12. 📋 `As per Family (Faisal Vai) New` / `As per Family (Faisal Vai)` — Newer & Older Versions

**What they are:** Two iterations of the family tracking sheet — the "New" version uses cross-file `VLOOKUP` to pull both current and previous months from separate source files, while the older version uses live `SUMIF` against the current `rawdata`.

**How the "New" version works differently:**
- `Sept'24` column: `=VLOOKUP(Folio, [2]Sheet1!$A:$G, 7, FALSE)` — pulls from the September file.
- `Aug'24` column: `=VLOOKUP(Folio, [3]Sheet1!$A:$G, 7, FALSE)` — pulls from the August file.
- This means you can keep two prior-month files open alongside this one and compare three months simultaneously without copying data in.

---

### 13. 📃 `Share Cer Summary` — Certificate Issuance Reference

**What it is:** A quick reference table listing which director is linked to which representative company and what certificate type they need.

**How it works:**
- Maps director names to their represented companies (e.g. Mr. Syed Kamruzzaman → Aramit Thai Aluminium Limited).
- Also contains a PivotTable showing total shares per director, used to populate the certificate sheets.
- Acts as a lookup guide when deciding which certificate template to use.

---

### 14. 📜 `Shareholding Certificate - 1` — Passport/Travel Certificate (with Company Rep)

**What it is:** An official certificate issued on UCB letterhead certifying a director's share value in both BDT and USD — used for visa applications and foreign travel.

**How it works:**
- Cell D1 = Director name; D2 = Representative company name; set these once per certificate.
- The certificate body is **one large formula** that builds the entire paragraph automatically:
  - Calculates total share value: `shares × closing price` in BDT
  - Converts to USD using today's exchange rate
  - Calls a custom VBA macro `NumberToTextBNMC()` to convert figures to words in Bangladesh Bank format (e.g. "One Crore Twenty Lac...")
  - Detects gender from title (`Mrs.`/`Ms.` → "she", otherwise "he") for grammatically correct pronouns
  - Auto-formats the DSE closing date from the `date` sheet
- The reference number `NO.UCB/HO/BOARD/F-30(VI)-6/.../YEAR` is generated by formula.
- **To issue:** Change D1 (name), D2 (company), update the share price and exchange rate cells, then print.

---

### 15. 📜 `Shareholding Certificate - 2` — Standard Director Certificate (No Travel)

**What it is:** A simpler certificate for directors who hold shares personally (not via a representative company), also including a note about historical dividend yield (~10% per annum over 10 years).

**How it works:**
- Same auto-text formula approach as Certificate 1, but the paragraph includes: *"received about 10% Dividend per annum on an average on the shares mentioned above for the last ten years."*
- No passport/travel section — used for domestic financial or professional purposes.

---

## 🔄 Monthly Workflow — Step by Step

```
Step 1 → 📥  Paste fresh share register data into 'rawdata'
             (Replace all rows; keep the table headers)

Step 2 → 🔄  Refresh all PivotTables
             (Data → Refresh All, or Ctrl+Alt+F5)

Step 3 → 📅  Set the reporting date in 'date' sheet cell B3

Step 4 → ✅  Check 'Categorywise Summary' — totals should
             match the depository's grand total

Step 5 → 📋  Update 'Previous Month' column in 'Page 2 & 3'
             (link to last month's file if not already linked)

Step 6 → 👨‍👩‍👧  Update 'Share no of Previous Month' in family sheet
             — the Status column auto-labels buy/sell/dividend

Step 7 → 🖨️  Print or export:
             - page 1 + Page 2 & 3 → BSEC, DSE, CSE submissions
             - For Web → paste into exchange portal
             - Certificates → issue to directors as needed

Step 8 → 💾  Copy current rawdata values into 'Monthwise
             Category Summary' for the permanent archive
```

---

## ⚙️ Key Excel Techniques Used

| 🛠️ Technique | 📍 Where Used | 🎯 Why |
|---|---|---|
| `GETPIVOTDATA()` | All output sheets | Pull live totals from the central PivotTable without hard-coding cell refs |
| `SUMIF()` on structured table | `rawdata`, Top 20 | Aggregate shares by name across multiple BO accounts |
| `INDEX/MATCH` on `Combined` key | Family sheets | Unique lookup by BOID+FOLIO to handle duplicate names |
| Cross-file `VLOOKUP` (`[2]Sheet1`) | New family sheet | Compare current vs. prior month files side by side |
| Smart `IF` chain for buy/sell/dividend | Family sheets | Auto-classify share changes without manual tagging |
| `CONCATENATE` + `TEXT()` for certificates | Certificate sheets | Build full letter paragraphs as a single formula |
| `NumberToTextBNMC()` VBA macro | Certificates | Convert numbers to Bengali/Bangladesh Bank word format |
| `TODAY()` + `YEAR()` in ref numbers | Cover letters | Auto-update reference numbers and dates each month |

---

## 🏛️ Regulatory Context

This file fulfils the monthly disclosure obligation under:

> **Regulation 35(2) of the DSE/CSE (Listing) Regulations 2015** — requiring all listed companies to submit a statement of shareholding position of sponsors, directors, foreigners, institutions, and shareholders holding 5% or more of securities to **BSEC**, **DSE**, and **CSE** every month.

---

## ⚙️ Technology

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)

- 📂 Format: **XLSM** (Excel Macro-Enabled) — VBA macro `NumberToTextBNMC()` required for certificates
- 🐍 Readable via `openpyxl` Python library
- 🏛️ Domain: **Capital Market Compliance / Listed Company Secretariat** (Bangladesh)

---

## 👤 Author

**Shafat** — Financial Data Analyst
> 📌 *Monthly regulatory shareholding disclosure system for UCB PLC Board Secretariat.*

---

## 🔒 License

Confidential corporate governance data. Intended for internal compliance use only.
