# 💰 Dividend Reconciliation — Year 2022
### United Commercial Bank Limited (UCB)

> 🎯 **Purpose in one line:** After UCB disburses its annual cash dividend to ~40,000+ shareholders through multiple payment channels, this workbook tracks every single payment, catches every mismatch, and produces a fully reconciled balance confirming that every taka sent out of the bank's account is accounted for.

---

## 🏦 File Information

| 🏷️ Attribute | 📋 Details |
|---|---|
| 📂 File Name | `reconciliation_2022.xlsb` |
| 🏛️ Issuer | United Commercial Bank Limited (UCB) |
| 📁 Format | Excel Binary Workbook (`.xlsb`) — optimised for large datasets |
| 💵 Dividend Type | 5% Cash Dividend from Retained Earnings |
| 📅 Disbursement Year | 2022 (processed July–September 2023) |
| 💰 Gross Dividend Pool | BDT 703,118,321.50 |
| 🔢 Net Amount Payable | BDT 605,481,088.85 |

---

## 🧠 How the Whole System Works — Big Picture

The core challenge this file solves: UCB disburses dividend through **four separate channels** (Warrant, BEFTN phases, UCB in-house cheque), each generating its own transaction records. Returns come back from different channels for different reasons. The workbook stitches all these streams together into one unified truth.

```
┌─────────────────────────────────────────────────────────┐
│           📋 DIVIDEND FILE (Source of Truth)            │
│   ~40,000 shareholders × Net Amount per shareholder     │
└───────────────────────┬─────────────────────────────────┘
                        │ feeds into
        ┌───────────────┼─────────────────┐
        ▼               ▼                 ▼
  📜 Warrant       🏦 BEFTN           🧾 UCB Cheque
  (physical        (Bangladesh         (in-house
   dividend         Electronic          bank deposit)
   warrant)         Fund Transfer)
        │               │                 │
        ▼               ▼                 ▼
   Paid ✅         Paid ✅            Paid ✅
   Unpaid ❌       Returned ❌        Cleared ✅
   Bank charge ⚠️  (invalid acct)    Bounced ❌
        │               │                 │
        └───────────────┼─────────────────┘
                        ▼
            ┌─────────────────────┐
            │  📊 Query Sheet     │  ← Master reconciliation
            │  (every shareholder │     per shareholder row
            │   accounted for)    │
            └──────────┬──────────┘
                       ▼
            ┌─────────────────────┐
            │  📋 Summary Sheet   │  ← Final balance check:
            │  Gross = Paid +     │    BDT 703M in = BDT 703M out
            │  Unpaid + Returns   │
            └─────────────────────┘
```

---

## 🗂️ Sheet-by-Sheet Breakdown — What Each Sheet Does & How

### 1. 📋 `Query.` — Master Shareholder Register (Payment Status)

**What it is:** The primary data sheet — one row per shareholder BO account, showing the current payment status of their dividend.

**How it works:**
- Each row carries: `Folio`, `BOID`, `Name`, `Shares`, `Gross Dividend`, `Deduction` (withholding tax), `Net Amount`, `Paid Amount`, `Unpaid Amount`, and `Payment Status`.
- The `Delivery Type` column tags each shareholder's payment method: `Warrant`, `BEFTN`, or blank (cheque).
- `Return Status` and `Delivery Status` flags indicate whether a payment came back and why.
- This sheet is the source that feeds `Query` (the full reconciliation sheet) — it reflects the raw dividend file as imported from the share depository, before reconciliation columns are added.
- **Key rule:** `Net Amount = Gross Dividend − Deduction`. Every row's deduction is exactly 10% of gross (standard withholding tax rate for listed company cash dividend under Bangladesh tax law).

---

### 2. 🔁 `Transfer Mode` — BEFTN Reference Mapper

**What it is:** A lookup table that maps each BEFTN reference number to its transfer system and digit-length validation flag.

**How it works:**
- Each BEFTN transaction gets a `Ref No.` (e.g. `23000001`) and a `System` label (`BEFTN`) and `Digit` count (`8`).
- The digit count is used elsewhere to validate that account numbers are the right length — a common source of BEFTN failures is account numbers with too few or too many digits.
- Other sheets cross-reference this table to confirm whether a given reference was sent via BEFTN and whether the account format was valid before submission.

---

### 3. 🧾 `Cheque Issue` — Physical Cheque Disbursement Log

**What it is:** A register of every physical cheque issued by UCB for dividend payment — mostly used for returned BEFTN payments that are then re-disbursed by cheque.

**How it works:**
- Each row = one issued cheque: date, warrant number, payee name, cheque amount, actual amount, disbursed amount.
- **Two difference columns** catch two separate types of error:
  - `Difference` = `Chq Amount − Actual Amount` → catches data entry errors (amount on cheque ≠ what the system recorded).
  - `Difference.` = `Actual Amount − Disbursed Amount` → catches fractional rounding issues (Bangladesh cheques round to the nearest paisa, so BDT 2,525.775 becomes 2,525.78, creating a BDT 0.005 difference).
- `Check If Cheque Issued Double` column flags `OK` or raises an alert if the same warrant number was issued a cheque twice — preventing duplicate payments.
- `Cheque Clearance` column marks each cheque as `Cleared` once the bank confirms it was deposited and honoured.

---

### 4. 🔍 `Query` — Full Reconciliation Master (All Payment Channels)

**What it is:** The most detailed sheet — extends every shareholder row with columns for every payment channel, mismatch flag, and unpaid balance. This is where the actual reconciliation logic lives.

**How it works — channel by channel:**

Each shareholder row has dedicated columns for every payment path:

| 🏷️ Column Group | 🎯 What it Checks |
|---|---|
| `Warrant Paid` | Amount confirmed paid via physical warrant |
| `Warrant Paid Mismatch` | Flag if warrant paid ≠ net amount |
| `Bank Charge Warrant` | UCB branch charges for in-house warrant deposits |
| `Warrant Unpaid` | Remaining unpaid after warrant processing |
| `BEFTN Return` | Amount returned by BEFTN (failed electronic transfer) |
| `BEFTN Return Mismatch` | Flag if returned amount ≠ expected |
| `BEFTN 2nd Phase` | Re-sent via a second BEFTN batch after corrections |
| `UCB Return` | Amount returned specifically from UCB's own BEFTN channel |
| `Cheque Issue` | Cheque issued for this shareholder's returned payment |
| `Cheque Mismatch` / `Cheque Mismatch.` | Rounding/data entry variance on the cheque |
| `Unpaid Amount` | Final unpaid balance after all channels are accounted for |

- **The reconciliation logic:** For every shareholder, `Net Amount = Paid Amount + Unpaid Amount`. If a payment was returned via BEFTN, the `BEFTN Return` column captures it, and the `Unpaid Amount` goes back up. When a cheque is issued and cleared for that return, `Cheque Issue` captures it and `Unpaid Amount` drops back to zero.
- `Payment Status` (`TRUE`/`FALSE`) is the final verdict — `TRUE` means fully paid, `FALSE` means still outstanding.

---

### 5. ✅ `validation list` — Dropdown Control

**What it is:** A two-cell helper sheet storing `yes` / `no` — the source for data validation dropdowns used in status columns across the workbook.

**How it works:** Excel data validation dropdowns reference this named list rather than hardcoding options. This ensures consistent, typo-free status flags throughout all sheets.

---

### 6. 🏦 `Bank Account` — UCB Bank Statement Import

**What it is:** The raw bank statement from UCB's dividend disbursement account — every debit and credit that touched the account during the entire dividend payment process.

**How it works:**
- Each row = one bank transaction: date, cheque/ref number, narration, debit, credit, running balance.
- Two helper columns parse the BEFTN reference number out of the narration text: `BEFTN No with Space` and `BEFTN No` (trimmed) — because the raw narration string embeds the account number deep inside a long description like *"In-House Cheque Deposit 2300049 From Account ... To Account 0773210000000384"*.
- `Ref. LEN` validates that the extracted reference is exactly 7 digits — BEFTN account numbers must be 7 digits; anything else flags a potential mismatch.
- `Repeat of BEFTN No` flags if the same BEFTN number appears more than once in the statement — catching any duplicate transaction.
- This sheet is compared against `Query` to confirm that every debit recorded in the bank statement matches a corresponding paid row in the shareholder ledger. The difference between the bank's closing balance and the sum of all payments is the reconciliation gap being closed.

---

### 7. 🔙 `UCB Return` — UCB BEFTN Return Register

**What it is:** A dedicated register of all BEFTN transfers that UCB's own banking system returned — meaning the money was sent electronically but bounced back before reaching the shareholder.

**How it works:**
- Each row = one failed transfer with: shareholder category, reference number, folio/BOID, name, shares, gross/net dividend, bank account, bank name, branch, and routing number.
- Return reasons are implicitly captured by the mismatch — an account number that doesn't exist in the routing network, a dormant account, or a routing number mismatch.
- The amounts here feed directly into the `UCB Return` column of `Query` — when a return is recorded here, the `Query` sheet picks it up and marks that shareholder's `Unpaid Amount` as restored, triggering the cheque issuance process.

---

### 8. 🔢 `Fraction Return` — Micro-Amount BEFTN Returns

**What it is:** A sub-register for returned BEFTN payments involving fractional (very small) amounts — typically from shareholders who own odd numbers of shares resulting in fractional dividend amounts.

**How it works:**
- Each row represents a tiny returned transfer (amounts like BDT 4.92, 3.07, 10.45) with a BEFTN ID, narration, and the exact return reason code:
  - `R04` — Invalid Account Number
  - `R03` — No Account / Unable to Locate Account
- These small returns are handled separately because fractional amounts are too small to re-issue individually via cheque. They are typically pooled and processed together in a batch.
- The total of all fraction returns (`BDT 36,679.61`) is cross-checked against the `Summary` sheet's `BEFTN Return` total to ensure no micro-amounts are lost.

---

### 9. 💳 `Charges` — Bank Service Charges Register

**What it is:** A log of all bank charges, VAT, and cheque processing fees debited from the dividend disbursement account during the payment process.

**How it works:**
- The first entries are account maintenance charges (`BDT 300`) and VAT (`BDT 45`) — standard monthly banking fees that appear in the bank statement.
- Subsequent entries are per-cheque clearing fees ranging from `BDT 34.50` to `BDT 1,150` depending on cheque value and branch.
- These charges are **not** part of the dividend payable to shareholders — they are UCB's own costs. The `Summary` sheet separates them out so the reconciliation is not distorted by these amounts.
- Any charge in this sheet that doesn't match a corresponding debit in `Bank Account` is a reconciliation gap.

---

### 10. 📊 `Summary` — Executive Reconciliation Dashboard

**What it is:** The final proof sheet — the "balance sheet" of the entire dividend operation. This is what management and auditors look at.

**How it works — four reconciliation checks:**

**Check 1 — Overall Balance:**
```
Gross Dividend (BDT 703,118,321.50)
= Total Paid (BDT 678,998,457.95)
+ Outstanding Unpaid (BDT 24,119,863.55)
✅ Difference: BDT 37.96 (rounding) → Bank Balance confirms
```

**Check 2 — Payment Channel Split:**
| 📦 Channel | 💰 Amount (BDT) |
|---|---|
| Warrant Issued | 324,236,936.10 |
| Warrant Paid | 307,830,780.75 |
| Warrant Unpaid | 16,406,155.35 |
| BEFTN Total | 219,382,411.40 |
| BEFTN Paid | 214,167,524.18 |
| BEFTN Return | 5,214,874.51 |

**Check 3 — BEFTN Phase Breakdown:**
Shows each BEFTN batch sent on specific dates, with the exact amount sent and confirmed received — any gap between sent and confirmed = pending return.

**Check 4 — Warrant Count:**
- Warrants Issued: 115
- Warrants Cleared: 103
- Kept with Bank: 4,361
- Outstanding: tracked individually in `Query`

---

### 11. 📝 `Notes` — BEFTN Outward Transaction Log

**What it is:** A clean record of every BEFTN outward batch transfer sent from UCB, organised by date and batch reference.

**How it works:**
- Each row = one BEFTN batch: date, reference, external reference, BEFTN number, narration, and debit amount.
- Six main BEFTN batches were sent (July 18–23 and September 17, 2023) — the dates correspond to when UCB's treasury actually pushed the money out to the BEFTN network.
- These rows are cross-matched against the `Bank Account` sheet — every debit here must appear as a debit in the bank statement. Any row in `Notes` without a matching bank statement entry is an unrecorded transaction.

---

### 12. 🔢 `fraction` — Fractional Bonus Share Settlement

**What it is:** Handles a special case — shareholders in "Group B & C" who received fractional bonus shares as part of the dividend, which cannot be issued as partial shares and must be settled in cash.

**How it works:**
- Fractional shareholders (e.g. `FRACTIONAL BONUS AC`, `ICB SUSPENSE FOR FRACTION`) are assigned warrant/reference numbers.
- Their gross dividend, deduction, and net amount are calculated normally.
- The total fractional gross dividend (`BDT 76,497`) flows into the main `Summary` totals to ensure the overall grand total still balances.

---

## 🔄 How the Reconciliation Is Done — Step by Step

```
Step 1 → 📥  Import dividend file from share depository
             → paste into 'Query.' as the master register

Step 2 → 📤  Issue warrants, submit BEFTN batches, process
             in-house cheques → record in 'Bank Account'

Step 3 → 🔁  BEFTN returns arrive (R03/R04 errors)
             → log in 'UCB Return' and 'Fraction Return'
             → 'Query' auto-updates each shareholder's
               Unpaid Amount upward

Step 4 → 🧾  Issue physical cheques for returned BEFTN
             → log in 'Cheque Issue'
             → 'Query' captures cheque amounts and marks
               mismatch if rounding difference exists

Step 5 → 💳  Log all bank charges in 'Charges'
             → keep them separate from dividend amounts

Step 6 → 📊  Open 'Summary' — all four balance checks
             must show ✅ zero difference (or <BDT 1
             rounding gap)

Step 7 → 📋  Any row in 'Query' with Payment Status = FALSE
             and Unpaid Amount > 0 remains open
             → escalate for manual follow-up
```

---

## ⚙️ Key Excel Techniques Used

| 🛠️ Technique | 📍 Where Used | 🎯 Why |
|---|---|---|
| Structured Table references `rawdata[[col]]` | All sheets | Self-documenting formulas that auto-expand with new rows |
| `SUMIF` on Folio/BOID | Query, Summary | Aggregate per-shareholder payments across multiple entries |
| Dual difference columns | Cheque Issue | Separately catches data-entry errors vs. rounding errors |
| Return reason codes (R03/R04) | Fraction Return | Categorise BEFTN failures for targeted correction |
| Running balance column | Bank Account | Continuously verify bank balance matches expected |
| BEFTN digit-length validation | Bank Account | Catch malformed account numbers before submission |
| Cross-sheet grand total checks | Summary | Four independent reconciliation proofs for auditability |

---

## 🏛️ Regulatory & Operational Context

Under **Bangladesh Securities and Exchange Commission (BSEC)** rules, a listed company must disburse its declared cash dividend within **30 working days** of the AGM and report any unclaimed dividend to BSEC. This workbook provides the audit trail proving that:

1. Every shareholder's net entitlement was correctly calculated (after 10% withholding tax).
2. Every payment attempt was made via the required channels.
3. Every returned payment was re-attempted and documented.
4. The total disbursed + outstanding = gross dividend declared. ✅

---

## ⚙️ Technology

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)

- 📂 Format: **XLSB** (Excel Binary) — handles 40,000+ shareholder rows efficiently
- 🐍 Readable via `pyxlsb` Python library
- 🏛️ Domain: **Capital Market Operations / Dividend Administration** (Bangladesh)

---

## 👤 Author

**Shafat** — Financial Data Analyst
> 📌 *Final reconciliation workbook for UCB's 2022 cash dividend disbursement.*

---

## 🔒 License

Confidential financial operations data. Intended for internal audit and compliance use only.
