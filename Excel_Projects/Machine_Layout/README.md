# 🧵 Jacquard Machine-Wise Layout & Production Tracker
### Garment / Knitting Factory — Daily Shift Operations Dashboard

> 🎯 **Purpose in one line:** Every shift, operators log what each Jacquard machine is running — and this workbook instantly shows supervisors which of the 200 machines are active, which buyer's order each is working on, how much was produced, and how much time was lost — all in one live floor-map view.

---

## 🏭 File Information

| 🏷️ Attribute | 📋 Details |
|---|---|
| 📂 File Name | `Jacquard_MC_Wise_Layout.xlsx` |
| 🏭 Domain | Garment / Textile Manufacturing (Jacquard Knitting) |
| ⚙️ Total Machines | 200 (M/C No. 1–100 on left floor, 101–200 on right floor) |
| 🕐 Shifts | 2 shifts per day × 12 hours each = 720 minutes per shift |
| 📦 Sample Buyer | Kiabi |
| 🧶 Sample Styles | 16Gesfi, Gpans |

---

## 🧠 How the Whole System Works — Big Picture

This workbook follows a strict **data-entry once, view everywhere** pipeline. Operators fill in only two sheets. Every other sheet — the combined view, the summary dashboard, and the physical floor map — updates automatically.

```
┌─────────────────────────────────────────────────────────┐
│  ✍️  OPERATOR DATA ENTRY (done each shift)               │
│                                                         │
│  '1st Shift Issue Sheet'  ──► MC 1–200, Shift 1        │
│  '2nd Shift Issue Sheet'  ──► MC 1–200, Shift 2        │
│                                                         │
│  Per machine: Buyer, Style, P.O, Colour, Size,         │
│               Production (pcs), Lost Time (min)        │
└───────────────────────┬─────────────────────────────────┘
                        │ all other sheets read from here
        ┌───────────────┼──────────────────┐
        ▼               ▼                  ▼
┌──────────────┐ ┌─────────────────┐ ┌──────────────────┐
│ (1st+2nd)    │ │  Summery Sheet  │ │   MC Layout      │
│ Shift        │ │  (KPI dashboard │ │   (physical       │
│ Combined     │ │   + style totals│ │   floor map of   │
│ (per-MC      │ │   across both   │ │   all 200 M/Cs   │
│  status +    │ │   shifts)       │ │   with live      │
│  style-wise  │ │                 │ │   status per box)│
│  counts)     │ │                 │ │                  │
└──────────────┘ └─────────────────┘ └──────────────────┘
```

---

## 🗂️ Sheet-by-Sheet Breakdown — What Each Sheet Does & How

### 1. 📝 `1st Shift Issue Sheet` — Shift 1 Data Entry (The Input)

**What it is:** The daily data entry form for the first 12-hour shift — one row per machine, 200 machines total (1–100 on the left side, 101–200 on the right side of the same sheet).

**How it works:**

Each machine row has these columns filled by the operator:

| 📋 Column | ✏️ Entered By | 🎯 Purpose |
|---|---|---|
| `M/C No` | Pre-filled | Machine number (1–200) |
| `Issue Date` | Operator | Date this machine was issued this order |
| `Working Hour (Min)` | Auto (`=12*60`) | Always 720 — one 12-hour shift in minutes |
| `Buyer` | Operator | Buyer name (e.g. Kiabi) |
| `Style` | Operator | Style code (e.g. 16Gesfi, Gpans) |
| `P.O` | Operator | Purchase Order number |
| `Colour` | Operator | Fabric/yarn colour |
| `Size` | Operator | Garment size being produced |
| `Production` | Operator | Pieces produced this shift on this machine |
| `Lost Time (Min)` | Operator | Minutes this machine was idle or stopped |
| `Effective Time (Min)` | **Auto formula** | Actual productive time |
| `Remarks` | Operator | Free text notes |

**The key formula — Effective Time:**
```
=IF(Lost Time > 0,  Working Hours − Lost Time,  720)
```
This means: if any downtime was entered, subtract it from the 720-minute shift. If no downtime was recorded, the machine is assumed to have run the full shift. This prevents blanks from inflating totals — a machine with no lost time logged is treated as fully productive, not as zero.

**The layout — two panels side by side:**
- Left panel: M/C 1–100
- Right panel: M/C 101–200
- Both panels have identical column structures, sitting in the same sheet to make one-screen comparison easy for supervisors.

**The total row (row 105):**
```
Total Production  = SUM(J5:J104)   ← all 100 left-side machines
Total Lost Time   = SUM(K5:K104)
Total Eff. Time   = SUM(L5:L104)
```
(Same totals mirror on the right-side panel.)

---

### 2. 📝 `2nd Shift Issue Sheet` — Shift 2 Data Entry (Identical Structure)

**What it is:** An exact copy of the 1st Shift Issue Sheet — same layout, same formulas — used for the second 12-hour shift.

**How it works:**
- Operators fill in the same fields for their shift: same machine numbers, but the Buyer, Style, Production, and Lost Time may differ because a machine can switch orders between shifts.
- The `=12*60` formula for Working Hours is hardcoded — it does not change between shifts since both are 12-hour shifts.
- All downstream sheets (Combined, Summary, MC Layout) pull from **both** shift sheets simultaneously, so the full day's picture is always live.

---

### 3. 🔀 `(1st+2nd) Shift Combined` — The Aggregation Engine

**What it is:** The calculation hub that aggregates both shift sheets into two views — a **style-wise machine count** at the top, and a **per-machine live status panel** below.

**How it works — two parts:**

**Part A — Style-Wise Machine Count (top section):**

For each Buyer + Style combination, this section answers: *"How many machines are currently running this style, and what is the total production from them?"*

```excel
M/C Count per Style (Left, 1st Shift):
=COUNTIF('1st Shift Issue Sheet'!F5:F104, "16Gesfi")

Production per Style (Left, 1st Shift):
=SUMIF('1st Shift Issue Sheet'!F5:F104, "16Gesfi",
       '1st Shift Issue Sheet'!J5:J104)
```

This is done for all four combinations:
- Left side × 1st Shift
- Right side × 1st Shift
- Left side × 2nd Shift
- Right side × 2nd Shift

A Grand Total row at the bottom uses `=IF(SUM(...)>0, SUM(...), 0)` — the `IF` guard prevents showing a negative zero if all machines are idle, keeping the display clean.

**Part B — Per-Machine Live Status (lower section):**

For every one of the 200 machines, this section shows:

| 📋 Column | 🔍 Formula Logic |
|---|---|
| `M/C` | Machine number (static) |
| `Qty` | `='1st Shift Issue Sheet'!J5` — live production pull |
| `Status` | `=IF(Qty > 0, "Running", "Not Running")` |
| `Count` | `=COUNT(Issue Date cell)` — 1 if date filled, 0 if blank |
| `Cmmnts` | `=IF(style_changes >= 2, "Alarm", "")` — fires if a machine has switched styles 2+ times in a shift |

The **Alarm flag** is the most operationally important feature: if a machine's style column has been updated 2 or more times (tracked via a hidden counter column `AG`), it raises `"Alarm"` — alerting the supervisor that this machine may be having issues or was improperly reassigned.

---

### 4. 📊 `Summery Sheet` — Daily KPI Dashboard

**What it is:** The management-level daily summary — combines all four shift/side combinations into final day totals.

**How it works — two sections:**

**Section A — Style-Wise Total (across all shifts and sides):**

Each style row sums from all four quadrants of the Combined sheet:
```excel
Total M/C for Style:
= Combined!E9 + Combined!J9 + Combined!O9 + Combined!T9
  (Left-1st)   (Right-1st)   (Left-2nd)   (Right-2nd)
```
This gives a single number: *"How many of the 200 machines ran Style X today across both shifts?"*

**Section B — Factory-Wide KPIs:**

| 📊 KPI | 🔢 Formula | 🎯 Meaning |
|---|---|---|
| Total M/C | `=200` | Total machines in the factory (static) |
| Total Working Hours (Min) | `=72000*4` | 200 machines × 720 min × 4 sides/shifts |
| Total Production | Sum of all 4 shift totals from both sheets | Total pieces produced today |
| Total Lost Time (Min) | Sum of all lost time columns | Total downtime across factory today |
| Total Effective Time (Min) | Sum of all effective time columns | Actual productive machine-minutes today |

**The Efficiency Formula (implicit):**
```
Efficiency % = Total Effective Time / Total Working Hours
             = (Total Working Hours − Total Lost Time) / Total Working Hours
```
This is not written as a single cell but all components are present — any manager can divide the two cells to get the day's OEE (Overall Equipment Effectiveness).

---

### 5. 🗺️ `MC Layout` — Physical Floor Map (The Visual Centrepiece)

**What it is:** A grid that replicates the actual physical layout of the factory floor — each cell block represents one machine in its real position, showing the style currently running on it and any alarm status.

**How it works:**

Each machine block occupies a small cell cluster in the grid, arranged in rows of 5 machines across. The pattern skips M/C 05, 10, 15... because every 5th position is a physical aisle on the factory floor — the layout in Excel mirrors the actual floor plan.

Each machine block shows two pieces of live data, both pulled from the Combined sheet:
```excel
Style label:   ='(1st+2nd) Shift Combined'!D30   ← buyer/style text
Alarm status:  ='(1st+2nd) Shift Combined'!G30   ← "Alarm" or blank
```

This means when a supervisor opens the `MC Layout` tab, they see the factory floor as it is right now — which machine is running which style, and which ones have raised an alarm — without navigating any other sheet.

**The `PATH` label** at the bottom of the layout marks the factory walkway — a visual reference line that matches the physical passage on the floor.

---

## 🔄 Daily Workflow — Step by Step

```
Start of Shift
──────────────
Step 1 → 📋  Operator opens '1st Shift Issue Sheet'
             (or '2nd Shift Issue Sheet' for second shift)

Step 2 → ✏️  For each running machine, fill in:
             Issue Date · Buyer · Style · P.O · Colour · Size

Step 3 → 🏭  As the shift runs, update Production (pcs)
             and Lost Time (min) for each machine

             ↓ Everything below is automatic ↓

Step 4 → ⚡  'Effective Time' auto-calculates per machine
             (720 min − Lost Time, or 720 if no downtime)

Step 5 → 📊  '(1st+2nd) Shift Combined' updates live:
             - Running / Not Running status per machine
             - Style-wise machine count and production totals
             - Alarm flag fires if any machine changed style 2x

Step 6 → 🗺️  'MC Layout' reflects the floor in real time —
             supervisors can glance at this to see the whole
             factory status without opening any other tab

End of Day
──────────
Step 7 → 📈  Open 'Summery Sheet' for the daily totals:
             total production · lost time · effective time
             across all 200 machines and both shifts
```

---

## ⚙️ Key Excel Techniques Used

| 🛠️ Technique | 📍 Where Used | 🎯 Why |
|---|---|---|
| `=12*60` hardcoded working time | Both shift sheets | Every shift is exactly 720 min — no manual entry, no errors |
| `=IF(Lost>0, 720−Lost, 720)` | Effective Time col | Blanks treated as full-run, not zero — prevents understating productivity |
| `COUNTIF` on Style column | Combined sheet | Count machines per style without helper columns |
| `SUMIF` on Style + Production | Combined sheet | Sum production per style in one formula |
| `=IF(Qty>0, "Running", "Not Running")` | Combined sheet | Live machine status flag — no VBA needed |
| `=IF(changes>=2, "Alarm", "")` | Combined sheet | Detects machines that switched jobs mid-shift |
| `=IF(SUM(...)>0, SUM(...), 0)` | Grand Total rows | Guards against displaying negative zero on empty days |
| Cross-sheet cell references | MC Layout | Floor map cells directly mirror Combined sheet values |
| Mirrored aisle gaps (skip every 5th) | MC Layout | Excel grid matches physical factory floor plan |
| Four-quadrant aggregation | Summery Sheet | Left+Right × Shift1+Shift2 summed to single day total |

---

## 🏭 Industry Context

**Jacquard machines** are programmable knitting/weaving machines that produce patterned fabric. In a garment factory, they are typically run in a large hall with numbered physical positions. Managing them requires:

- Knowing which buyer order each machine is working on at any moment
- Tracking production pace per machine to catch underperformers early
- Recording downtime (machine breakdown, yarn changeover, setup time) to calculate true efficiency
- Raising alerts when a machine is switched between orders too frequently (a sign of planning or quality issues)

This workbook replaces a manual paper-based floor tracking system with a live digital dashboard that requires only minimal operator input — enter the order details and production count, and all KPIs calculate themselves.

---

## ⚙️ Technology

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)

- 📂 Format: **XLSX** (Excel Workbook)
- 🐍 Readable via `openpyxl` Python library
- 🏭 Domain: **Garment Manufacturing / Industrial Production Tracking** (Textile sector)

---

## 👤 Author

**Shafat** — Data Analyst
> 📌 *Jacquard machine-wise daily production and layout tracking system for factory floor operations.*

---

## 🔒 License

Internal factory operations data. For production management use only.
