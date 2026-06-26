# 🏦 Shareholding Certificate of Directors & Sponsors
### United Commercial Bank PLC (UCB)

> 📋 A macro-enabled Excel workbook (`.xlsm`) tracking the **shareholding positions of UCB Directors, Sponsors, and their family members** — including month-by-month share movement, stock dividend adjustments, and official shareholding certificate generation.

---

## 📁 File Information

| 🏷️ Attribute | 📋 Details |
|---|---|
| 📂 File Name | `Shareholding_Certificate_of_Directors_Sponsors.xlsm` |
| 🏦 Issuer | United Commercial Bank Limited (UCB) |
| 📊 File Type | Excel Macro-Enabled Workbook (`.xlsm`) |
| 🗓️ Coverage Period | July 2019 — August 2023 |
| 💰 Total Paid-up Capital | BDT 11,595,437,190 |
| 🔢 Total Paid-up Shares | 1,159,543,719 |

---

## 👥 Board of Directors & Designation Registry

| 👤 Name | 🏷️ Designation |
|---|---|
| Mrs. Rukhmila Zaman | Former Director & Immediate Past Chairman |
| Mr. Bashir Ahmed | Vice-Chairman |
| Mr. Anisuzzaman Chowdhury | Director |
| Mr. Asifuzzaman Chowdhury | Director |
| Mrs. Roxana Zaman Chaudhury | Director |
| Mr. M. A. Sabur | Director |
| Mr. Bazal Ahmed | Director |
| Mrs. Masuma Parvin | Director |
| Mr. Syed Kamruzzaman | Director |
| Mr. Muhammed Shah Alam | Director |
| Mr. Kanak Kanti Sen | Director |
| Mr. Md. Aksed Ali Sarker | Director |
| Mr. Akhter Matin Chaudhury | Director |
| Mr. Syed Mohammed Nuruddin | Director |
| Mr. Touhid Shipar Rafiquzzaman | Independent Director |
| Dr. Aparup Chowdhury | Independent Director |
| Professor Dr. Iftekhar Uddin Chowdhury | Independent Director |
| Mr. Arif Quadri | Managing Director & CEO |
| Mr. Mohammed Shawkat Jamil | Managing Director & CEO |
| Mr. Hajee Yunus Ahmed | Sponsor |

---

## 📊 Shareholding Category Breakdown (as of Dec 2019)

| 🔠 Category | 🏷️ Holder Type | 📈 Percentage (%) |
|---|---|---|
| A | Sponsors / Promoters & Directors | **36.67%** |
| B | Government | 0.81% |
| C | Local Institutions | 18.54% |
| D | Foreign Institutions | 1.47% |
| E | General Public | 42.51% |
| | **Total** | **100%** |

---

## 🗂️ Workbook Structure — 11 Sheets

### 1. 📈 `Share Holding Position` — Group Summary Timeline
> Monthly shareholding percentage for each investor category from **December 2019 to August 2023** (~45 reporting periods).

**Tracks:** Sponsors/Directors, Government, Local Institutions, Foreign Institutions, Public — with percentage recalculated each month.

---

### 2. 👨‍👩‍👧 `As per Family` — Director Family Share Ledger
> The most detailed sheet — tracks **individual shareholding** of each Director and their family members month by month.

**Key Features:**
- 📅 Monthly snapshot from July 2019 to August 2023
- 📂 Shares split into: `Electronic`, `Paper`, `Suspense`
- 🔄 Buy/Sell difference tracked between each consecutive month
- 🎁 Stock dividend adjustments: **5% (2020)**, **10% (2021)**, **5% (2022)**
- ➗ Fraction share deduction handling after bonus allocation

**Director Family Groups tracked:**
- 👑 Mrs. Rukmila Zaman & Family
- 🤝 Mr. Anisuzzaman Chowdhury & Family
- *(additional director families)*

---

### 3. 🔗 `Own and Family` — Consolidated Director View
> Combines a director's **personal holdings** and **family holdings** into one consolidated view per reporting date.

**Notable Individuals:**
- Mr. Saifuzzaman Chowdhury — 2,068 shares (July 2023)
- Mrs. Rukhmila Zaman — 29,532,675 shares
- Mr. Anisuzzaman Chowdhury — 41,323,559 shares
- Mr. Asifuzzaman Chowdhury — 41,647,858 shares
- Mrs. Nur Nahar Zaman — 8,675,670 shares

---

### 4. 📋 `Detail Up 1st July, 2023` — Latest Position Table
> A structured data table (`Table1`) listing all directors and associated entities with their **exact share position as of 1st July 2023**.

**Key Columns:**
- 👤 `Name`, `Family Group`, `Designation` (auto-VLOOKUP)
- 🆔 `BO ID`
- 📦 Electronic / Paper / Suspense / Total Share
- 📊 `Share %` — calculated against total paid-up shares (1,476,548,475)

---

### 5. 🏷️ `Designation` — Director Title Registry
> Master lookup table mapping each director's name to their official board designation, used via `VLOOKUP` across the workbook.

---

### 6. 📊 `Summary` — Per-Director BO Account Summary
> Summarizes all BO accounts per director with total shares and calculated market value (at BDT 10 face value per share).

**Formula highlight:** Uses `OFFSET` + `COUNTA` dynamic range to auto-expand as new accounts are added.

---

### 7. 📜 `certificate director` — Director Share Certificate Template
> Auto-populating **official certificate template** for Director-held shares.

- 🔢 Certificate number auto-formatted: `NO.UCB/HO/BOARD/F-30(VI)-6/.../YEAR`
- 📅 Date auto-filled via `=TODAY()`
- Addressed "TO WHOM IT MAY CONCERN"

---

### 8. 📜 `certificate family Custom` — Custom Family Certificate
> Customizable version of the shareholding certificate for specific family members. Currently set to **Mr. Bashir Ahmed**.

---

### 9. 📜 `certificate family` — Standard Family Certificate
> Standard certificate template for family members, with date set to `=TODAY()-1` (previous business day convention).

---

### 10. 🔢 `Calculation` — Monthly Category % Engine
> Backend calculation sheet providing per-month shareholding percentages for **15 investor categories** across 40+ months.

**Categories include:**
- Sponsor, Directors, Sponsor Family, Sponsor Government
- Foreign Sponsor, Local Institution, Foreign Institution
- General Investor (Local / NRB / Foreign)
- Mutual Fund, Other Funds, Government (Others)
- Associated Companies, Suspense

---

### 11. 🔧 `Workings` — Director BO Account Detail Workings
> Raw working sheet listing every BO account per director with individual share counts — the source data feeding into `Detail Up 1st July, 2023`.

**Sample entries:**
| 👤 Director | 🔢 Total Shares (Jul 2023) |
|---|---|
| Mrs. Rukhmila Zaman | 28,126,358 |
| Mr. Bashir Ahmed | 28,284,349 |
| Mr. Anisuzzaman Chowdhury | 39,355,771 |
| Mr. M. A. Sabur | 30,540,602 |
| Mr. Bazal Ahmed | 28,181,894 |

---

## 🔄 Key Processes Automated in This Workbook

```
📅 Monthly Tracking     →  Share position updated every 1st of the month
🎁 Bonus Allocation     →  5% (2020), 10% (2021), 5% (2022) stock dividends
➗ Fraction Handling    →  Fractional bonus shares deducted automatically
📜 Certificate Gen.    →  Director & Family certificates auto-populated
🔍 VLOOKUP Designation →  Auto-fetches director title across all sheets
📐 Dynamic Ranges       →  OFFSET + COUNTA for auto-expanding BO accounts
```

---

## ⚙️ Technology & Tools

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)

- 📂 Format: **XLSM** (Excel Macro-Enabled Workbook)
- 🐍 Readable via `openpyxl` Python library
- 🏛️ Domain: **Capital Market Compliance / Corporate Governance** (Bangladesh)
- 📡 Regulatory context: DSE/BSEC shareholding disclosure requirements

---

## 👤 Author

**Shafat** — Financial Data Analyst
> 📌 *Shareholding certificate and tracking workbook for UCB Board of Directors & Sponsors.*

---

## 📜 License

🔒 This file contains confidential corporate governance and shareholder data. Intended for internal regulatory use only.
