# 🏦 Shareholding Certificate of Directors & Sponsors
### United Commercial Bank PLC (UCB)

> 🎯 **Purpose in one line:** Every time a UCB Director needs to prove their shareholding — for a visa, a loan, or a board filing — this workbook looks up their current share position, calculates the market value in BDT and USD, and generates a print-ready official certificate automatically, without anyone typing a single number.

---

## 🏦 File Information

| 🏷️ Attribute | 📋 Details |
|---|---|
| 📂 File Name | `Shareholding_Certificate_of_Directors___Sponsors.xlsm` |
| 🏛️ Issuer | United Commercial Bank Limited (UCB) |
| 📁 Format | Excel Macro-Enabled Workbook (`.xlsm`) |
| 📅 Coverage Period | July 2019 — August 2023 |
| 💰 Total Paid-up Capital | BDT 11,595,437,190 |
| 🔢 Total Paid-up Shares | 1,159,543,719 |

---

## 🧠 How the Whole System Works — Big Picture

This workbook has two distinct jobs running in parallel:

**Job 1 — Historical Tracking:** Maintain a month-by-month record of every director's shareholding from 2019 onwards, automatically detecting whether changes are due to stock dividends, purchases, or sales.

**Job 2 — Certificate Generation:** Take the current share count for any director, multiply by the DSE closing price, convert to USD, spell the amount out in words, and produce a signed-ready certificate paragraph — entirely through formulas.

```
┌───────────────────────────────────────────────────────────────┐
│  📊 HISTORICAL TRACKING ENGINE                                │
│                                                               │
│  'As per Family' ──► Month-by-month share counts             │
│  (Jul 2019 → Aug 2023)  per director + family member         │
│       │                                                       │
│       ▼ Bonus events auto-detected                           │
│  5% bonus (2020) ──► shares × 1.05 → fraction deducted      │
│  10% bonus (2021) ──► shares × 1.10 → fraction deducted     │
│  5% bonus (2022) ──► shares × 1.05 → fraction deducted      │
│       │                                                       │
│       ▼                                                       │
│  'Own and Family' ──► Consolidated view per director family  │
│       │                                                       │
│       ▼                                                       │
│  'Detail Up 1st July 2023' ──► Latest snapshot for          │
│                                  certificate use             │
└────────────────────────────────┬──────────────────────────────┘
                                 │
                                 ▼
┌───────────────────────────────────────────────────────────────┐
│  📜 CERTIFICATE GENERATION ENGINE                             │
│                                                               │
│  'Calculation' ──► Monthly % per investor category           │
│  'Share Holding Position' ──► Group summary timeline         │
│  'Summary' ──► Per-director BO account totals                │
│       │                                                       │
│       ▼                                                       │
│  'certificate director' ──► Type 1: director with           │
│                               company representation         │
│  'certificate family'   ──► Type 2: personal holding        │
│  'certificate family    ──► Type 3: custom per-request      │
│   Custom'                                                     │
└───────────────────────────────────────────────────────────────┘
```

---

## 🗂️ Sheet-by-Sheet Breakdown — What Each Sheet Does & How

### 1. 📈 `Share Holding Position` — Group Summary Timeline

**What it is:** A wide timeline showing UCB's overall shareholder composition percentage for every reporting date from December 2019 to August 2023 (~45 snapshots).

**How it works:**
- Rows are the five investor categories: Sponsors/Directors (A), Government (B), Local Institute (C), Foreign Institute (D), Public (E).
- Each column is one reporting date. From the 5th column onward, values are formulas referencing the `Calculation` sheet: `=Calculation!J6` — meaning the `Calculation` sheet is the engine and this sheet just displays the result in timeline form.
- The first four columns (Dec 2019 – Feb 2020) are hardcoded values from before the formula system was built.
- Below the table: company metadata is stored here — name, paid-up capital (`BDT 11,595,437,190`), and total shares (`1,159,543,719`) — because other sheets pull these via cell reference rather than typing them repeatedly.

---

### 2. 👨‍👩‍👧 `As per Family` — The Core Historical Ledger

**What it is:** The most data-dense sheet — a complete month-by-month record of every share count for every director and their family members, with stock dividend adjustments calculated inline.

**How it works — three distinct phases:**

**Phase 1 — Monthly Snapshots (Jul 2019 → Aug 2021):**
- Each director group (e.g. "Mrs. Rukmila Zaman & Family") is a row block with sub-rows for each individual BO account.
- Columns march forward monthly. Each cell holds the share count at the start of that month.
- Share counts are static until a dividend event.

**Phase 2 — Stock Dividend Adjustment (three rounds):**
When a stock dividend is declared, the sheet calculates the new total automatically:

```
For each director:
  Gross Bonus  = Previous shares × 5% (or 10%)
  Less Fraction = fractional portion (e.g. 0.3 shares → rounded down)
  Actual Bonus  = INT(Gross Bonus) → whole shares only
  New Total    = Previous shares + Actual Bonus
```

- Three bonus events are embedded:
  - **5% Stock Dividend for 2020** (record date Sept 2021)
  - **10% Stock Dividend for 2021** (record date June 2022)
  - **5% Stock Dividend for 2022** (record date Aug 2023)
- After each bonus, a "Checking" column verifies `New Total = Old Total × (1 + rate)` — a built-in accuracy check.

**Phase 3 — Buy/Sell Detection:**
Between each monthly snapshot, a `Difference` column computes `Current Month − Previous Month`. If this difference is **not** explained by a bonus event, it means the director bought or sold shares that month. This difference is flagged for review.

---

### 3. 🔗 `Own and Family` — Consolidated Director View

**What it is:** A cleaner, flattened version of `As per Family` that shows the total position per director family group rather than individual BO accounts — used as the source for certificates and regulatory reports.

**How it works:**
- For each director, rows for their personal BO accounts and their family members' accounts are listed separately, then summed in a total row: `=H5+H4` (for a two-account family).
- The `Total Share for all BO Account` column sums across all BO accounts belonging to that family group.
- This consolidated number is what appears on certificates — "Mr. Anisuzzaman Chowdhury holds **41,323,559** shares" — not a per-BO breakdown.
- Unlike `As per Family` which has hundreds of columns going back to 2019, this sheet only carries the two most recent months (July 2023 and August 2023), keeping it fast to print and easy to audit.

---

### 4. 📋 `Detail Up 1st July, 2023` — Current Snapshot Table

**What it is:** A formal Excel Table (`Table1`) of every director and family member with their exact holding as of 1st July 2023 — the reference point for current certificates.

**How it works:**
- The `Designation` column uses `=IFERROR(VLOOKUP(Name, Designation!$A$2:$B$84, 2, FALSE), "Family")` — if the name is found in the `Designation` sheet, it shows their board title (e.g. "Chairman"); if not found, it defaults to "Family" (for non-director family members who hold shares).
- The `Share %` column = `Total Share / 1,476,548,475` — the denominator is the post-2022-bonus total paid-up shares, so percentages reflect the current diluted base.
- This table is the data source referenced by the `Summary` sheet and the certificate templates.

---

### 5. 🏷️ `Designation` — Board Title Registry

**What it is:** A two-column master lookup: Name → Board Designation.

**How it works:**
- Used exclusively as a VLOOKUP source by `Detail Up 1st July, 2023` and `Summary`.
- Contains all current and former directors, sponsors, independent directors, Managing Directors/CEOs, and Shariah Supervisory Committee members.
- When a director's role changes (e.g. from Director to Chairman), only this sheet needs updating — every certificate and report that pulls the title updates automatically.

---

### 6. 📊 `Summary` — Per-Director BO Account Aggregator

**What it is:** For each director, aggregates all their BO accounts into one total share count and market value.

**How it works:**
- Share total uses a dynamic `OFFSET + COUNTA` range: `=SUM(OFFSET($B$3, COUNTA($B$4:$B$23), 0, -COUNTA($B$4:$B$23), 1))` — this formula automatically expands or contracts as BO accounts are added or removed, without needing manual range adjustments.
- Market value = Total Shares × BDT 10 (face value) as a base; actual market value uses the closing price from the certificate sheet.
- This sheet feeds the certificate templates: when you open `Shareholding Certificate - 1` or `- 2`, the share count it displays is read from here.

---

### 7. 📜 `certificate director` — Certificate Template (With Company Representation)

**What it is:** The official UCB Company Secretariat certificate for directors who represent a company on the board (e.g. Mr. Aksed Ali Sarker representing Aromatic Properties Limited).

**How it works:**
- **Reference number** is auto-generated: `=CONCATENATE("NO.UCB/HO/BOARD/F-30 (VI)-6/   /", YEAR(H15))` — always carries the correct year.
- **Date** is `=TODAY()` — prints today's date whenever the file is opened.
- The certificate body paragraph is built by a single formula that pulls director name, company name, share count, BDT value, USD equivalent, exchange rate, and DSE closing date — all from reference cells — and assembles them into a grammatically complete paragraph ready for printing.
- **To issue:** Update the name and company cells at the top, confirm the share count, verify the closing price and exchange rate → print. No typing anywhere else.

---

### 8. 📜 `certificate family Custom` & `certificate family` — Personal Holding Certificates

**What they are:** Two versions of the certificate for directors who hold shares personally (not via a company representative structure).

**How they differ:**
- `certificate family` uses `=TODAY()-1` for the date — convention to show "as of previous business day" which is standard for financial certificates where the closing price is always from the prior day.
- `certificate family Custom` uses `=TODAY()` — for cases where the issuing date should match the request date exactly.
- Both reference the same share count and valuation cells, so only the date convention differs.

**How the body paragraph is constructed:**
- The formula detects the director's title prefix (`Mrs.` / `Ms.` → "she"; otherwise "he") to ensure grammatically correct pronouns — e.g. *"she is financially sound and solvent"* vs. *"he is financially sound and solvent"*.
- Share count, BDT market value, USD equivalent, and the DSE date are all formula-driven — the certificate text updates the moment the share price or exchange rate cells are refreshed.

---

### 9. 🔢 `Calculation` — Monthly Category Percentage Engine

**What it is:** The backend calculation sheet that computes the shareholding percentage for each of 15 investor categories across every reporting month.

**How it works:**
- Structured as repeating 5-column blocks, one block per month. Each block has: `SL`, `Category Name`, `Percentage`, `Shares`, `Total`.
- The 15 categories are: Sponsor, Directors, Sponsor Family, Sponsor Government, Foreign Sponsor, General Investor Local, General Investor NRB, General Investor Foreign, Local Institution, Foreign Institution, Mutual Fund, Other Funds, Government (Others), Associated Companies, Suspense.
- Values for the earlier months (2019–2021) are hardcoded from historical reports. Values from mid-2021 onward are formula-driven.
- `Share Holding Position` reads from this sheet via direct cell references (`=Calculation!J6`) — so updating one number here propagates to the group summary timeline automatically.

---

### 10. 📝 `Workings` — Raw BO Account Detail

**What it is:** The lowest-level data sheet — every BO account for every director, with the exact Electronic/Paper/Suspense breakdown.

**How it works:**
- For directors with multiple BO accounts (e.g. Mr. M. A. Sabur has 6 separate BO accounts across different depository participants), each account is a separate row.
- The `Total Share for all BO Account` column sums only the first BO row per director using a formula that counts sibling rows: `=H5+H4`.
- This sheet is the source that feeds `Own and Family` — when a new BO account is opened for a director or family member, a new row is added here and the consolidated total updates automatically.

---

## 🔄 How to Issue a New Certificate — Step by Step

```
Step 1 → 📋  Check 'Detail Up 1st July, 2023' (or latest
             snapshot) for the director's current share count

Step 2 → 💹  Update the share price cell with today's DSE
             closing price for UCB shares

Step 3 → 💱  Update the exchange rate cell (BDT per USD)

Step 4 → 🏷️  Open the correct certificate template:
             - Director representing a company → 'certificate director'
             - Personal holding, standard    → 'certificate family'
             - Personal holding, dated today  → 'certificate family Custom'

Step 5 → ✏️  Set the name cell (D1) to the director's name
             [For company-rep cert: also set D2 to company name]

Step 6 → 👁️  The certificate paragraph builds automatically —
             verify the share count, BDT value, and USD value

Step 7 → 🖨️  Print → sign → issue
```

---

## 🔄 How to Update for a New Month — Step by Step

```
Step 1 → 📥  Receive updated share position from CDBL
             (Central Depository Bangladesh Limited)

Step 2 → 📊  Add a new month column to 'As per Family'
             for each director family group

Step 3 → 🔍  Check the Difference column — any unexplained
             change triggers a buy/sell flag for verification

Step 4 → 🎁  If a stock dividend was declared this year:
             - Add bonus calculation rows
             - Deduct fractions
             - Carry forward new total to next month column

Step 5 → 🔗  Update 'Own and Family' with latest totals

Step 6 → 📋  Refresh 'Detail Up...' snapshot table

Step 7 → 📈  Add new percentage block to 'Calculation'
             → 'Share Holding Position' timeline auto-extends
```

---

## ⚙️ Key Excel Techniques Used

| 🛠️ Technique | 📍 Where Used | 🎯 Why |
|---|---|---|
| `VLOOKUP` on Designation table | Detail Up, Summary | Auto-fetch board titles without hardcoding |
| `OFFSET + COUNTA` dynamic range | Summary | BO account list self-expands when accounts are added |
| `IFERROR(..., "Family")` | Detail Up | Gracefully handle non-director family members |
| Fraction deduction formula | As per Family | `INT(shares × rate)` — only whole shares are issued |
| Bonus accuracy check column | As per Family | Built-in post-dividend verification before next entry |
| Gender-detection `IF` chain | Certificates | `IF(title="Mrs.", "she", "he")` for correct pronouns |
| Single-formula certificate body | Certificates | Entire paragraph built as one `CONCATENATE` expression |
| `=TODAY()-1` date convention | certificate family | Matches financial industry standard for closing price |
| Wide timeline layout | As per Family | All history visible in one sheet without file switching |

---

## 🏛️ Regulatory & Operational Context

Under **BSEC (Bangladesh Securities and Exchange Commission)** rules, directors of listed companies are required to:
1. Disclose their shareholding position to the exchange every month.
2. Notify the exchange within **3 working days** of any change in their holding.
3. Produce a shareholding certificate from the company secretariat when required for travel visas or financial transactions.

This workbook automates all three requirements — the tracking sheets provide the disclosure data, and the certificate templates produce the signed documents.

---

## ⚙️ Technology

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)

- 📂 Format: **XLSM** (Excel Macro-Enabled Workbook)
- 🐍 Readable via `openpyxl` Python library
- 🏛️ Domain: **Capital Market Compliance / Corporate Governance** (Bangladesh)

---

## 👤 Author

**Shafat** — Financial Data Analyst
> 📌 *Director shareholding tracking and certificate generation system for UCB Board Secretariat.*

---

## 🔒 License

Confidential corporate governance data. Intended for internal compliance use only.
