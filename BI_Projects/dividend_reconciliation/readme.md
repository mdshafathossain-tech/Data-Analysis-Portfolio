# 💰 Dividend Reconciliation 2024 — Power BI
### United Commercial Bank Limited (UCB) — Power BI Migration from Excel (.xlsb)

> 🎯 **Purpose in one line:** The same dividend reconciliation that was previously managed in a 12-sheet Excel binary workbook is now rebuilt entirely in Power BI — separating each payment channel into its own source file, joining them live through Power Query and DAX, and presenting a single real-time dashboard where every mismatch surfaces automatically as a KPI card.

---

## 🔄 Why This Version Exists — Excel → Power BI Migration

The [original reconciliation](https://github.com/mdshafathossain-tech/Data-Analysis-Portfolio/tree/main/Excel_Projects/Dividend_Reconciliation) was a single `.xlsb` file containing all payment data, return records, bank statements, and summary logic within 12 sheets. While functional, it had limitations:

| ⚠️ Excel (.xlsb) Limitation | ✅ Power BI Solution |
|---|---|
| All data in one binary file — hard to update individual channels | Each channel is a separate `.xlsx` source file — update one without touching others |
| Mismatch detection via manual formula columns | Mismatch columns are DAX calculated columns — auto-refresh on data change |
| Summary totals in a static sheet | Live KPI cards that respond to BOID/ID slicers instantly |
| No drill-down per shareholder | Two slicers (BOID + ID) filter every visual simultaneously |
| Warrant payment matching required VLOOKUP chains | `COUNTROWS(FILTER(...))` + `SUMX` DAX pattern matches cheque numbers directly to bank statement rows |

---

## 📁 Project Files

| 📂 File | 🛠️ Tool | 📋 Purpose |
|---|---|---|
| `reconciliation_v1.pbix` | Power BI Desktop | Full reconciliation dashboard |
| `dividend_list_2024.xlsx` | Excel | Master dividend register (`Query` + `Cheque` sheets) |
| `beftn.xlsx` | Excel | BEFTN transfer records |
| `beftn_return.xlsx` | Excel | Returned BEFTN payments |
| `ucb.xlsx` | Excel | UCB in-house transfer records |
| `ucb_paid.xlsx` | Excel | UCB confirmed payment records |
| `warrant.xlsx` | Excel | Physical dividend warrant register |
| `bank_account.xlsx` | Excel | Raw UCB bank statement |
| `all_bo_ref.xlsx` | Excel | Complete BO reference master (Folio + BOID + ID) |

---

## 🧠 How the Whole System Works — Big Picture

Unlike a typical Power BI report where one fact table feeds multiple visuals, this model uses `Query` as a **central reconciliation spine** — every payment channel table is joined to it, and DAX columns on `Query` compute each shareholder's payment status per channel automatically.

```
┌──────────────────────────────────────────────────────────────────┐
│  📥  SOURCE FILES (9 Excel workbooks)                            │
│                                                                  │
│  dividend_list_2024.xlsx ──► Query sheet  (master dividend list) │
│                         └──► Cheque sheet (cheque issuance log)  │
│  beftn.xlsx             ──► BEFTN sent records                   │
│  beftn_return.xlsx      ──► BEFTN returned records               │
│  ucb.xlsx               ──► UCB in-house transfer records        │
│  ucb_paid.xlsx          ──► UCB confirmed paid records           │
│  warrant.xlsx           ──► Physical warrant register            │
│  bank_account.xlsx      ──► Raw bank statement (all debits)      │
│  all_bo_ref.xlsx        ──► Folio/BOID/ID reference master       │
└───────────────────────┬──────────────────────────────────────────┘
                        │  Power Query joins all into model
                        ▼
┌──────────────────────────────────────────────────────────────────┐
│  🔀  POWER QUERY — Table Preparation                             │
│                                                                  │
│  Query table gets 3 LEFT OUTER JOINs applied in sequence:        │
│  Query ←── beftn_return  (brings in Beftn Return amount)         │
│  Query ←── ucb_paid      (brings in UCB Paid amount)             │
│  Query ←── Cheque        (brings in Cheque Issued amount)        │
│                                                                  │
│  NULLs from unmatched joins → replaced with 0 (not blank)       │
└───────────────────────┬──────────────────────────────────────────┘
                        │
                        ▼
┌──────────────────────────────────────────────────────────────────┐
│  🧮  DAX LAYER — Calculated Columns on Query & warrant/Cheque    │
│                                                                  │
│  Delivery Mode       ──► classifies each shareholder's channel   │
│  Beftn Paid          ──► Net Dividend if BEFTN sent & not returned│
│  Beftn Return Mismatch──► gap between sent & returned amount     │
│  UCB Paid Mismatch   ──► gap between UCB paid & net dividend     │
│  Warrant Paid        ──► matches warrant cheque# to bank stmt    │
│  Warrant Paid Mismatch──► gap between warrant paid & entitlement │
│  Cheque Paid         ──► matches cheque# to bank statement debits│
│  Cheque Bank Charge  ──► isolates bank fees from payment amount  │
└───────────────────────┬──────────────────────────────────────────┘
                        │
                        ▼
┌──────────────────────────────────────────────────────────────────┐
│  📊  DAX MEASURES — Summary KPIs                                 │
│                                                                  │
│  Total Paid    = Warrant + BEFTN + UCB + Govt + Cheque           │
│  Total Unpaid  = Net Dividend − Total Paid                       │
│  Bank Balance  = SUM(Credit) − SUM(Debit)                       │
│  Total Mismatch = Total Unpaid − Bank Balance  ← target: 0      │
│  Bank Charges  = filtered debit SUM (4 charge keywords)         │
│  Tax           = SUM(Deduction)                                  │
│  Govt Dividend = fixed lookup for ID "24000226"                  │
└───────────────────────┬──────────────────────────────────────────┘
                        │
                        ▼
┌──────────────────────────────────────────────────────────────────┐
│  📋  DASHBOARD — Page 1                                          │
│                                                                  │
│  12 KPI Cards  ──► all mismatch & balance figures at a glance   │
│  2 Slicers     ──► BOID or ID drill-down to any shareholder     │
│  1 Detail Table──► per-shareholder payment breakdown             │
│  1 Channel Table──► delivery mode vs net dividend vs cheque     │
└──────────────────────────────────────────────────────────────────┘
```

---

## 🗄️ Data Model — 9 Tables & Their Roles

### 📋 `Query` — The Reconciliation Spine (Fact Table)

Every shareholder is one row. This table holds the entitlement figures and, after Power Query joins, carries the payment amounts from each channel as additional columns.

| 🏷️ Column | 📐 Type | 📋 Source |
|---|---|---|
| `ID` | text | Unique shareholder ID — join key for all channel tables |
| `Folio` | text | Folio number from share depository |
| `BOID` | text | Beneficiary Owner ID |
| `Category` | text | Shareholder category |
| `Shares` | int | Number of shares held |
| `Gross Dividnd` | decimal | Total dividend before deduction |
| `Tax (%)` | decimal | Withholding tax rate |
| `Deduction` | decimal | Tax amount |
| `Net Dividend` | decimal | Final payable amount (Gross − Deduction) |
| `Investment/Margin` | text | Flag for special account types |
| `Beftn Return` | decimal | **Joined** from `beftn_return` |
| `UCB Paid` | decimal | **Joined** from `ucb_paid` |
| `Cheque Issued` | decimal | **Joined** from `Cheque` |
| `Delivery Mode` | text | **DAX** — channel classification |
| `Beftn Paid` | decimal | **DAX** — net if BEFTN sent and not returned |
| `Beftn Return Mismatch` | decimal | **DAX** — gap on partial BEFTN returns |
| `UCB Paid Mismatch` | decimal | **DAX** — gap between UCB paid and entitlement |

---

### 🏦 `bank_account` — Ground Truth

The raw bank statement is the ultimate authority — every debit must map to a shareholder payment. Warrant and cheque payment confirmation both depend entirely on matching `Cheque#.` values in this table.

| 🏷️ Column | 📋 Purpose |
|---|---|
| `Trans. Date` | Transaction date |
| `Cheque#.` | **Join key** to `warrant[Warrant]` and `Cheque[Cheque]` |
| `Trans. Details` | Full description — scanned by `CONTAINSSTRING` for bank charge keywords |
| `Debit` | Payments out |
| `Credit` | Receipts in |
| `Balance` | Running account balance |

---

### 📜 `warrant` — Physical Dividend Warrant Register

Holds warrant cheque numbers for shareholders receiving physical warrants. The `Warrant Paid`, `warrant Bank Charge`, and `Warrant Paid Mismatch` columns are all DAX calculated — computed by matching `warrant[Warrant]` to `bank_account[Cheque#.]`.

### 🔁 `beftn` — BEFTN Sent Records
Shareholders submitted for electronic transfer. 3 columns: `ID`, `ReceiverName`, `Amount`.

### 🔙 `beftn_return` — BEFTN Return Register
Bounced-back BEFTN payments. 5 columns including `ReceivingBank` and `ReceivingBranch` — identifying which bank rejected the transfer.

### 🏦 `ucb` + ✅ `ucb_paid` — UCB In-House Transfers
`ucb` = submitted for internal UCB transfer. `ucb_paid` = confirmed credited to shareholder. The difference between these two tables surfaces as `UCB Paid Mismatch`.

### 🧾 `Cheque` — Post-Return Cheque Log
Physical cheques issued for BEFTN-returned payments. Loaded from a second sheet within `dividend_list_2024.xlsx`.

### 🗂️ `all_bo_ref` — BO Reference Master
Full Folio/BOID/ID lookup table from the depository. Bridges the M:M relationship between `Query` and the BO reference space.

---

## 🔗 Relationships

| From | Key | To | Key | Cardinality | Filter |
|---|---|---|---|---|---|
| `Query` | `ID` | `beftn` | `ID` | M:1 | Single |
| `Query` | `ID` | `ucb` | `ID` | M:1 | Single |
| `Query` | `ID` | `warrant` | `ID` | M:1 | Both |
| `Query` | `ID` | `ucb_paid` | `ID` | M:1 | Both |
| `Query` | `ID` | `beftn_return` | `ID` | **Inactive** | Single |
| `Query` | `ID` | `all_bo_ref` | `ID` | M:M | Both |
| `Cheque` | `ID` | `Query` | `ID` | M:M | Both |
| `beftn_return` | `ID` | `beftn` | `ID` | 1:1 | Both |
| `bank_account` | `Cheque#.` | `warrant` | `Warrant` | M:1 | Both |

> ⚠️ `Query → beftn_return` is **inactive** because return data is already embedded in `Query` via Power Query LEFT OUTER JOIN at load time. The inactive relationship remains as a fallback for direct cross-filtering.

---

## 🔀 Power Query (M Code) — All 9 Tables

### 📋 `Query` — Master with 3 Embedded LEFT OUTER JOINs

The most complex M script — loads the master list then joins three return/paid tables sequentially. Nulls replaced with 0 after each join to prevent DAX blank issues.

```m
let
    Source = Excel.Workbook(
        File.Contents("...\dividend_list_2024.xlsx"), null, true),
    Query_Sheet        = Source{[Item="Query", Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Query_Sheet, [PromoteAllScalars=true]),
    #"Changed Type"    = Table.TransformColumnTypes(#"Promoted Headers", {
        {"Category", type text}, {"ID", type text}, {"Folio", type text},
        {"BOID", type text}, {"Shares", Int64.Type}, {"Value of Shares", Int64.Type},
        {"Gross Dividnd", type number}, {"Tax(%)", type number},
        {"Deduction", type number}, {"Net Dividend", type number},
        {"Investment/Margin", type any}
    }),

    // ── JOIN 1: BEFTN return amounts ──────────────────────────────
    #"Merged Queries"        = Table.NestedJoin(#"Changed Type", {"ID"}, beftn_return, {"ID"}, "beftn_return", JoinKind.LeftOuter),
    #"Expanded beftn_return" = Table.ExpandTableColumn(#"Merged Queries", "beftn_return", {"Amount"}, {"Amount"}),
    #"Renamed Columns"       = Table.RenameColumns(#"Expanded beftn_return", {{"Amount", "Beftn Return"}}),
    #"Replaced Value"        = Table.ReplaceValue(#"Renamed Columns", null, 0, Replacer.ReplaceValue, {"Beftn Return"}),

    // ── JOIN 2: UCB confirmed paid amounts ────────────────────────
    #"Merged Queries1"    = Table.NestedJoin(#"Replaced Value", {"ID"}, ucb_paid, {"ID"}, "ucb_paid", JoinKind.LeftOuter),
    #"Expanded ucb_paid"  = Table.ExpandTableColumn(#"Merged Queries1", "ucb_paid", {"Amount"}, {"Amount"}),
    #"Renamed Columns1"   = Table.RenameColumns(#"Expanded ucb_paid", {{"Amount", "UCB Paid"}}),
    #"Replaced Value1"    = Table.ReplaceValue(#"Renamed Columns1", null, 0, Replacer.ReplaceValue, {"UCB Paid"}),

    // ── JOIN 3: Cheque issued amounts ─────────────────────────────
    #"Merged Queries2"  = Table.NestedJoin(#"Replaced Value1", {"ID"}, Cheque, {"ID"}, "Cheque", JoinKind.LeftOuter),
    #"Expanded Cheque"  = Table.ExpandTableColumn(#"Merged Queries2", "Cheque", {"Amount"}, {"Amount"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Expanded Cheque", {{"Amount", "Cheque Issued"}}),
    #"Replaced Value2"  = Table.ReplaceValue(#"Renamed Columns3", null, 0, Replacer.ReplaceValue, {"Cheque Issued"})
in
    #"Replaced Value2"
```

### 🏦 `bank_account`
```m
let
    Source = Excel.Workbook(File.Contents("...\bank_account.xlsx"), null, true),
    bank_account_Sheet = Source{[Item="bank_account", Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(bank_account_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"Trans. Date",    type date},   {"Cheque#.",      type text},
        {"Ref.",           type text},   {"Narration",     type text},
        {"Trans. Details", type text},   {"Debit",         type number},
        {"Credit",         type number}, {"Balance",       type number}
    })
in  #"Changed Type"
```

### 📜 `warrant`
```m
let
    Source = Excel.Workbook(File.Contents("...\warrant.xlsx"), null, true),
    Sheet1_Sheet = Source{[Item="Sheet1", Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"Folio", type text}, {"BOID", type text}, {"Name", type text},
        {"Warrant", type text}, {"Share", Int64.Type},
        {"Dividend", type number}, {"ID", type text}
    })
in  #"Changed Type"
```

### 🔁 `beftn`
```m
let
    Source = Excel.Workbook(File.Contents("...\beftn.xlsx"), null, true),
    beftn_Sheet = Source{[Item="beftn", Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(beftn_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"ID", type text}, {"ReceiverName", type text}, {"Amount", type number}
    })
in  #"Changed Type"
```

### 🔙 `beftn_return`
```m
let
    Source = Excel.Workbook(File.Contents("...\beftn_return.xlsx"), null, true),
    beftn_return_Sheet = Source{[Item="beftn_return", Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(beftn_return_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"ID", type text}, {"Amount", type number}, {"ReceiverName", type text},
        {"ReceivingBank", type text}, {"ReceivingBranch", type text}
    })
in  #"Changed Type"
```

### 🏦 `ucb` / ✅ `ucb_paid` / 🧾 `Cheque` / 🗂️ `all_bo_ref`
All follow the same pattern — `Excel.Workbook` → select named sheet → promote headers → enforce types. Type details for each:

```m
-- ucb:      {"ID", type text}, {"Share Holder Name", type text}, {"Amount", type number}
-- ucb_paid: {"ID", type text}, {"Folio/BOID", type text}, {"Share Holder Name", type text}, {"Amount", type number}
-- Cheque:   {"ID", type text}, {"Cheque", type text}, {"Amount", type number}
-- all_bo_ref: {"Folio", type text}, {"BOID", type text}, {"ID", type text}
```

---

## 🧮 DAX — All Measures & Calculated Columns

### 📊 Measures

#### `Total Paid`
> Five channels summed. The warrant and cheque components use measures (not column sums) because they need context-aware `FILTER` logic against the bank statement.
```dax
Total Paid =
[Warrant_Paid] + SUM(Query[Beftn Paid]) + SUM(Query[UCB Paid])
+ [Govt Dividend] + [Cheque_Paid]
```

#### `Total Unpaid`
```dax
Total Unpaid =
SUM(Query[Net Dividend]) - [Total Paid]
```

#### `Bank Balance`
```dax
Bank Balance =
SUM(bank_account[Credit]) - SUM(bank_account[Debit])
```

#### `Total Mismatch`
> **The reconciliation proof figure.** Must equal zero for the reconciliation to close. Non-zero means a gap exists between the shareholder register and the bank account.
```dax
Total Mismatch =
[Total Unpaid] - [Bank Balance]
```

#### `Warrant_Paid`
> Looks up each warrant cheque number against the bank statement. Subtracts bank charges to return the net dividend paid — not the gross debit including fees.
```dax
Warrant_Paid =
IF(
    COUNTROWS(FILTER(warrant, warrant[ID] IN VALUES(Query[ID]))) > 0,
    CALCULATE(
        SUM(warrant[Warrant Paid]) - SUM(warrant[warrant Bank Charge]),
        FILTER(warrant, warrant[ID] IN VALUES(Query[ID]))
    ),
    0
)
```

#### `Cheque_Paid`
> Same `COUNTROWS(FILTER(...))` + `CALCULATE` pattern as `Warrant_Paid` but applied to the `Cheque` table for post-return cheques.
```dax
Cheque_Paid =
IF(
    COUNTROWS(FILTER(Cheque, Cheque[ID] IN VALUES(Query[ID]))) > 0,
    CALCULATE(
        SUM(Cheque[Cheque Paid]) - SUM(Cheque[Cheque Bank Charge]),
        FILTER(Cheque, Cheque[ID] IN VALUES(Query[ID]))
    ),
    0
)
```

#### `Govt Dividend`
> Hard-coded lookup for the government shareholder — paid via challan (special payment method), not through any of the four standard channels.
```dax
Govt Dividend =
CALCULATE(MAX(Query[Net Dividend]), Query[ID] = "24000226")
```

#### `Bank Charges`
> Scans `Trans. Details` free text for four known bank fee descriptions — isolating charges from dividend payments in the total debit sum.
```dax
Bank Charges =
CALCULATE(
    SUM(bank_account[debit]),
    FILTER(
        bank_account,
        CONTAINSSTRING(bank_account[Trans. Details], "Online Inter Branch Charge") ||
        CONTAINSSTRING(bank_account[Trans. Details], "Value Added Tax")            ||
        CONTAINSSTRING(bank_account[Trans. Details], "Cheque Issuance")            ||
        CONTAINSSTRING(bank_account[Trans. Details], "Account Maintenance Charge")
    )
)
```

#### `Tax`
```dax
Tax = SUM(Query[Deduction])
```

---

### 🧩 DAX Calculated Columns

#### `Query[Delivery Mode]`
> Classifies each shareholder's payment channel by checking which channel table contains their ID. Priority order: BEFTN → UCB → Warrant → Government challan → KWB (catch-all).
```dax
Delivery Mode =
IF(COUNTROWS(FILTER(beftn,   beftn[ID]   = Query[ID])) > 0, "beftn",
IF(COUNTROWS(FILTER(ucb,     ucb[ID]     = Query[ID])) > 0, "UCB",
IF(COUNTROWS(FILTER(warrant, warrant[ID] = Query[ID])) > 0, "warrant",
IF(Query[ID] = "24000226", "challan", "kwb"))))
```

#### `Query[Beftn Paid]`
> Recognised as paid only when delivery mode is BEFTN AND no return exists. A zero-return means the money reached the shareholder.
```dax
Beftn Paid =
IF(AND([Delivery Mode] = "beftn", [Beftn Return] = 0), [Net Dividend], 0)
```

#### `Query[Beftn Return Mismatch]`
> Fires when a BEFTN payment partially bounced back — returned amount is less than the net dividend, leaving a gap. Investment/Margin accounts excluded (handled separately).
```dax
Beftn Return Mismatch =
IF(
    Query[Delivery Mode] = "beftn"
        && Query[Beftn Return] > 0
        && ISBLANK(Query[Investment/Margin]),
    Query[Net Dividend] - Query[Beftn Return],
    0
)
```

#### `Query[UCB Paid Mismatch]`
> Fires when UCB confirmed a payment but the confirmed amount differs from the net dividend entitlement.
```dax
UCB Paid Mismatch =
IF(
    Query[Delivery Mode] = "UCB" && Query[UCB Paid] > 0,
    Query[Net Dividend] - Query[UCB Paid],
    0
)
```

#### `warrant[Warrant Paid]`
> Row-level calculated column — scans `bank_account` for rows where `Cheque#.` matches this warrant's cheque number, then sums the debits found.
```dax
Warrant Paid =
IF(
    COUNTROWS(FILTER(bank_account, bank_account[cheque#.] = warrant[warrant])) > 0,
    SUMX(
        FILTER(bank_account, bank_account[Cheque#.] = warrant[warrant]),
        bank_account[debit]
    ),
    0
)
```

#### `warrant[warrant Bank Charge]`
> Same cheque-number match as above, but filtered further to only sum rows where `Trans. Details` contains a bank charge keyword — separating fees from the dividend payment.
```dax
warrant Bank Charge =
IF(
    COUNTROWS(FILTER(bank_account,
        bank_account[Cheque#.] = warrant[warrant] &&
        (CONTAINSSTRING(bank_account[Trans. Details], "Online Inter Branch Charge") ||
         CONTAINSSTRING(bank_account[Trans. Details], "Value Added Tax")            ||
         CONTAINSSTRING(bank_account[Trans. Details], "Cheque Issuance")            ||
         CONTAINSSTRING(bank_account[Trans. Details], "Account Maintenance Charge"))
    )) > 0,
    SUMX(FILTER(bank_account,
        bank_account[Cheque#.] = warrant[warrant] &&
        (CONTAINSSTRING(bank_account[Trans. Details], "Online Inter Branch Charge") ||
         CONTAINSSTRING(bank_account[Trans. Details], "Value Added Tax")            ||
         CONTAINSSTRING(bank_account[Trans. Details], "Cheque Issuance")            ||
         CONTAINSSTRING(bank_account[Trans. Details], "Account Maintenance Charge"))
    ), bank_account[debit]),
    0
)
```

#### `warrant[Warrant Paid Mismatch]`
> Uses `LOOKUPVALUE` to bring the shareholder's net dividend entitlement from `Query` into the `warrant` table context, then computes the gap against what was actually paid.
```dax
Warrant Paid Mismatch =
IF(
    warrant[Warrant_Paid] > 0,
    VAR matchingDividend =
        LOOKUPVALUE(Query[Net Dividend], Query[ID], warrant[ID])
    RETURN
        IF(NOT(ISBLANK(matchingDividend)), matchingDividend - warrant[Warrant_Paid], 0),
    0
)
```

#### `Cheque[Cheque Paid]` and `Cheque[Cheque Bank Charge]`
> Exact same `COUNTROWS(FILTER(...))` + `SUMX` pattern as the warrant equivalents — match `Cheque[Cheque]` to `bank_account[Cheque#.]`, sum debits, then isolate charge rows separately.

---

## 📊 Dashboard — Page 1

### 🎛️ Slicers
- **BOID slicer** — drill down to any individual shareholder by Beneficiary Owner ID
- **ID slicer** — drill down by internal shareholder ID

Both slicers filter all 12 KPI cards and both tables simultaneously.

### 💳 12 KPI Cards

| 🏷️ Card | 📐 Source | 🎯 What It Shows |
|---|---|---|
| Total Bank Debit | `SUM(bank_account[Debit])` | All money leaving the disbursement account |
| BEFTN Return Mismatch | `SUM(Query[Beftn Return Mismatch])` | Partial BEFTN return gaps |
| UCB Paid Mismatch | `SUM(Query[UCB Paid Mismatch])` | UCB confirmation vs entitlement gap |
| Warrant Paid Mismatch | `SUM(warrant[Warrant Paid Mismatch])` | Warrant payment vs entitlement gap |
| Net Dividend | `SUM(Query[Net Dividend])` | Total post-tax entitlement |
| Govt Dividend | `[Govt Dividend]` | Challan payment for government shareholder |
| Bank Charges | `[Bank Charges]` | Bank fees separated from payments |
| Total Paid | `[Total Paid]` | All channels combined |
| Total Unpaid | `[Total Unpaid]` | Outstanding balance |
| Bank Balance | `[Bank Balance]` | Net position of disbursement account |
| **Total Mismatch** | `[Total Mismatch]` | **Reconciliation gap — target: zero ✅** |
| Tax | `[Tax]` | Total withholding tax deducted |

### 📋 Tables

**Shareholder Detail Table** — one row per shareholder:
`ID · Folio · BOID · Net Dividend · Delivery Mode · Warrant_Paid · Beftn Paid · UCB Paid`

**Delivery Mode Summary Table** — grouped by channel:
`Delivery Mode · Net Dividend · Cheque`

---

## 🔄 Reconciliation Proof Logic

```
Reconciliation closes when:

  Total Paid + Total Unpaid  =  SUM(Net Dividend)   ← Shareholder side balances
  Total Unpaid               =  Bank Balance         ← Bank side matches
  ──────────────────────────────────────────────────
  Total Mismatch             =  0  ✅                ← Fully reconciled
```

If `Total Mismatch ≠ 0`, each mismatch card isolates the responsible channel:

| ⚠️ Card > 0 | 🔍 Investigate |
|---|---|
| BEFTN Return Mismatch | Partial return — shareholder received less than returned |
| UCB Paid Mismatch | UCB confirmation amount differs from entitlement |
| Warrant Paid Mismatch | Cheque cleared for wrong amount or bank charge included |
| `Total Unpaid > Bank Balance` | Payment not yet in bank statement — timing difference |

---

## 🔗 Relationship to Previous Excel Version

| 📋 Component | 📊 Excel (.xlsb) | 📊 Power BI (.pbix) |
|---|---|---|
| Data storage | All data in one binary file | 9 separate `.xlsx` source files |
| Return data | Manual paste into sheet | `beftn_return.xlsx` joined via Power Query |
| Channel classification | `Delivery Type` column, manually filled | `Delivery Mode` DAX calculated column — auto |
| Warrant matching | VLOOKUP chain across sheets | `COUNTROWS(FILTER)` + `SUMX` on bank statement |
| Bank charge isolation | Manual filter rows | `CONTAINSSTRING` DAX on `Trans. Details` |
| Mismatch detection | Formula columns per sheet | DAX columns auto-flagging every row on refresh |
| Drill-down | Manual filter on Query sheet | BOID + ID slicers — instant per-shareholder |
| Summary | Dedicated Summary sheet | 12 live KPI cards, always current |

---

## ⚙️ Technology

![Power BI](https://img.shields.io/badge/PowerBI-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![DAX](https://img.shields.io/badge/DAX-Measures_%26_Columns-blue?style=for-the-badge)
![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)

| 🔧 Tool | 🎯 Used For |
|---|---|
| **Power BI Desktop** | Data model, DAX, reconciliation dashboard |
| **Power Query (M)** | Loading 9 Excel files + 3 LEFT OUTER JOINs on `Query` |
| **DAX** | 8 measures + 9 calculated columns (columns + charge isolation) |
| **Excel (.xlsx)** | All 9 source data files |

---

## 👤 Author

**Md. Shafat Hossain** — Financial Data Analyst
> 📌 *Power BI rebuild of the UCB 2024 cash dividend reconciliation — migrated from Excel (.xlsb) to a live multi-source Power BI model.*

---

## 🔒 License

Confidential financial operations data. Intended for internal audit and compliance use only.
