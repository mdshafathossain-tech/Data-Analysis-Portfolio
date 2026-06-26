# 💰 Dividend Reconciliation Report — Year 2022

> 📊 A comprehensive financial reconciliation workbook tracking **5% Cash Dividend** distribution from retained earnings for the fiscal year 2022, processed via United Commercial Bank PLC. (UCB).

---

## 📁 File Information

| 🏷️ Attribute | 📋 Details |
|---|---|
| 📂 File Name | `reconciliation_2022.xlsb` |
| 📅 Reporting Year | **2022** |
| 🏦 Processing Bank | United Commercial Bank Ltd. (UCB) |
| 💵 Dividend Type | 5% Cash Dividend from Retained Earnings |
| 🗓️ Disbursement Start | July 05, 2023 |

---

## 💼 Financial Summary at a Glance

| 📊 Metric | 💲 Amount (BDT) |
|---|---|
| 💰 Gross Dividend | 703,118,321.50 |
| ➖ Total Deduction | 97,637,232.65 |
| ✅ Net Amount | 605,481,088.85 |
| 📜 Warrant Issued | 324,236,936.10 |
| ✔️ Warrant Paid | 307,830,780.75 |
| 🔄 BEFTN Disbursed | 219,382,411.40 |

---

## 🗂️ Workbook Structure

The workbook contains **12 sheets**, each serving a distinct role in the reconciliation process:

### 1. 📋 `Query.` — Master Shareholder Register
> Primary data sheet listing all shareholders with dividend details.

**Key Columns:**
- 🆔 `Folio` & `BOID` — Unique shareholder identifiers
- 👤 `Share Holder Name` — Registered shareholder name
- 📈 `Shares` — Number of shares held
- 💵 `Gross Dividend` — Total dividend before deduction
- ➖ `Deduction` — Tax / withholding deduction
- 💰 `Net Amount` — Final payable amount
- 🏦 `Bank Account No`, `Bank`, `Branch` — Disbursement bank details
- 🚚 `Delivery Type` — Warrant / BEFTN / Cheque
- 📌 `Payment Status` & `Return Status`

---

### 2. 🔁 `Transfer Mode` — BEFTN Reference Mapping
> Maps reference numbers to transfer systems (BEFTN) with digit validation.

**Key Columns:** `Ref No.`, `System`, `Digit`

---

### 3. 🧾 `Cheque Issue` — Physical Cheque Disbursement Log
> Tracks all physically issued cheques with clearance status.

**Key Columns:**
- 📅 `Date` — Cheque issue date
- 🔢 `Warrant No.` — Linked warrant reference
- 👤 `Name` — Payee name
- 💵 `Chq Amount` vs `Actual Amount` — Discrepancy check
- ✅ `Cheque Clearance` — Cleared / Pending status
- 🔍 `Check If Cheque Issued Double` — Duplicate safeguard

---

### 4. 🔍 `Query` — Detailed Reconciliation Tracker
> Extended version of the master register with full payment-mode breakdown and mismatch flags.

**Key Columns:**
- ✔️ `Warrant Paid`, `BEFTN Return`, `UCB Return`
- ⚠️ `Warrant Paid Mismatch`, `BEFTN Return Mismatch`
- 🏦 `Cheque Issue`, `Cheque Mismatch`
- 📌 `Unpaid Amount` — Outstanding balance per shareholder

---

### 5. ✅ `validation list` — Dropdown Values
> Simple Yes/No validation list used across data entry fields.

---

### 6. 🏦 `Bank Account` — UCB Bank Statement
> Raw bank account transaction log from UCB for cross-referencing disbursements.

**Key Columns:**
- 📅 `Trans. Date` — Transaction date
- 🔢 `Cheque No.` & `Ref.` — Bank reference details
- 📝 `Narration` — Transaction description
- 💸 `Debit` / `Credit` / `Balance`
- 🔁 `BEFTN No` — Electronic fund transfer ID

---

### 7. 🔙 `UCB Return` — UCB BEFTN Return Cases
> Shareholders whose BEFTN transfers were returned by UCB due to errors.

**Key Columns:**
- 👥 `Share Holder Category`
- ❌ Returned `Ref No.`, `Bank Account`, `Routing No`
- 💵 `Net Dividend` — Amount pending re-disbursement

---

### 8. 🔢 `Fraction Return` — Fractional BEFTN Returns
> Handles micro-amount BEFTN returns, typically due to invalid/unknown account numbers.

**Return Reasons Include:**
- ❗ `R04` — Invalid Account Number
- ❗ `R03` — No Account / Unable to Locate Account

---

### 9. 💳 `Charges` — Bank Service Charges
> Logs bank maintenance charges, VAT, and per-cheque processing fees deducted during disbursement.

**Notable Entries:**
- 🏦 Account Maintenance Charge: BDT 300
- 🧾 Value Added Tax (VAT): BDT 45
- 📋 Per-cheque clearing fees

---

### 10. 📊 `Summary` — Executive Financial Summary
> High-level reconciliation dashboard showing totals, paid vs unpaid breakdown, and variance analysis.

**Key Highlights:**
- 🟢 BEFTN Phases (P1–P4) disbursement tracking
- ⚖️ `Difference` between Gross Dividend and Total Paid
- 🏦 Final Bank Balance vs Net Amount reconciliation
- 📜 Warrant: Issued `115`, Cleared `103`, Kept with Bank `4,361`

---

### 11. 📝 `Notes` — BEFTN Outward Transaction Log
> Records all BEFTN outward transfers sent via UCB in batches.

| 📅 Date | 💸 Amount (BDT) | 🏷️ Batch |
|---|---|---|
| 18-07-2023 | 16,844,107.13 | BEFTN Phase 1 |
| 18-07-2023 | 69,489,328.28 | BEFTN Phase 2 |
| 18-07-2023 | 36,922,748.08 | BEFTN Phase 3 |
| 18-07-2023 | 72,629,154.33 | BEFTN Phase 4 |
| 23-07-2023 | 23,497,007.04 | BEFTN Phase 5 |

---

### 12. 🔢 `Fraction` — Fractional Share Dividend Settlement
> Handles fractional bonus/cash adjustments for Group B & C shareholders and ICB suspense accounts.

---

## 🔄 Payment Channels Used

```
📜 Dividend Warrant  →  Physical warrant issued to shareholders
🏦 BEFTN             →  Bangladesh Electronic Funds Transfer Network
🧾 Cheque            →  In-house cheque deposit (UCB branches)
```

---

## ⚙️ Technology & Tools

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)

- 📂 Format: **XLSB** (Excel Binary Workbook) for performance with large datasets
- 🐍 Readable via `pyxlsb` Python library
- 🏛️ Domain: **Capital Market / Corporate Dividend Management** (Bangladesh)

---

## 👤 Author

**Shafat** — Financial Data Analyst
> 📌 *Final version of the 2022 dividend reconciliation file.*

---

## 📜 License

This file contains confidential financial and shareholder data. Intended for internal use only. 🔒
