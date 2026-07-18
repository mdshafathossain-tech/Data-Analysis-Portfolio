# ☕ Daily Brew Coffee Ltd. — Business Intelligence Dashboard
### Power BI Case Study | End-to-End BI Analysis Project

> 🎯 **Purpose in one line:** Daily Brew Coffee Ltd. had transaction data from 15+ NYC stores but zero visibility into what it meant — this project turns that raw data into a 6-page interactive Power BI report that answers every strategic question leadership was asking, from peak-hour performance to customer loyalty to promotional impact.

---

## 🏢 Business Context

Daily Brew Coffee Ltd. is a fast-growing coffee chain operating across **New York City** with over 15 outlets. They sell a wide range of products — whole bean, espresso roasts, premium teas, and baked goods — and record every transaction. But despite having the data, leadership was making decisions based on gut feel, not facts.

**They didn't know:**
- Which stores were thriving and which were quietly underperforming
- Whether promotional items were boosting revenue or just cutting margin
- If loyal customers were returning or silently churning
- What time of day drove the most revenue — and whether staffing matched it

**This dashboard changes that.**

---

## 📁 Project Files

| 📂 File | 🛠️ Tool | 📋 Purpose |
|---|---|---|
| `BI_Analysis_-_Daily_Brew_Coffee_Ltd.pbix` | Power BI Desktop | 6-page interactive dashboard |
| `PowerBI_Dashboard_Case_Study_Assignment.docx` | Reference | Business case study & assignment brief |
| `PowerBI_Dataset_Dictionary_Detail.docx` | Reference | Full column-level data dictionary for all 7 tables |

---

## 🧠 How the Whole Project Works — Big Picture

The architecture follows a classic **star schema** — one central fact table surrounded by dimension lookup tables, all connected through a dedicated measures table.

```
                    ┌──────────────────────────────┐
                    │   🔢 _all_measures table       │
                    │   (all DAX measures live here) │
                    └──────────────┬───────────────┘
                                   │ referenced by all 6 pages
┌─────────────────┐                │
│  customer_lookup│──────┐         │
│  (demographics, │      │         ▼
│   loyalty cards)│      │  ┌─────────────────┐    ┌─────────────────┐
└─────────────────┘      └─►│                 │◄───│  store_lookup   │
                             │  sales_by_store │    │  (location,     │
┌─────────────────┐      ┌─►│  (FACT TABLE)   │    │   sqft, mgr)    │
│ employee_lookup │──────┘  │  transaction_id │    └─────────────────┘
│  (staff, role,  │         │  sale_date/time │
│   position)     │         │  store, staff,  │    ┌─────────────────┐
└─────────────────┘         │  customer,      │◄───│ product_lookup  │
                             │  product,       │    │  (category,     │
┌─────────────────┐         │  qty, price,    │    │   cost, retail, │
│    calendar     │──────►  │  promo_yn)      │    │   promo, new)   │
│  (date, month,  │         └────────┬────────┘    └─────────────────┘
│   quarter, year,│                  │
│   day type)     │         ┌────────▼────────┐
└─────────────────┘         │ food_inventory   │
                             │ (baked goods     │
┌─────────────────┐         │  stock tracking) │
│   time_band     │──────►  └─────────────────┘
│  (hour buckets  │
│   for heatmaps) │    ┌────────────┐   ┌──────────────────┐
└─────────────────┘    │ month_sort │   │ customer_summary  │
                        │ (sort order│   │ (aggregated view  │
                        │  for axis) │   │  of visit freq.)  │
                        └────────────┘   └──────────────────┘
```

**Two helper tables built specifically for visuals:**
- `time_band` — groups raw transaction hours into labelled buckets (e.g. "Morning Rush", "Afternoon Lull") for the peak-hour heatmap
- `month_sort` — provides a sort-order column so months display Jan→Dec on charts, not alphabetically

---

## 🗂️ Data Model — 7 Tables

### 📊 `sales_by_store` — The Fact Table (Centre of Everything)

Every row is one line item in one transaction. This is the table every measure aggregates.

| 🏷️ Column | 📋 What It Captures |
|---|---|
| `transaction_id` | Unique transaction identifier |
| `transaction_date` | Date of sale — joins to `calendar` |
| `transaction_time` | Exact time — used for hour-level heatmaps |
| `store_id` | Which store — joins to `store_lookup` |
| `staff_id` | Which employee handled it — joins to `employee_lookup` |
| `customer_id` | Which customer — joins to `customer_lookup` |
| `instore_yn` | Was it in-store or online? (Y/N) |
| `product_id` | What was sold — joins to `product_lookup` |
| `quantity_sold` | Units sold in this line item |
| `unit_price` | Price per unit at time of sale |
| `promo_item_yn` | Was this item on promotion? (Y/N) |

**Derived column added in Power BI:**
- `Hour` — extracted from `transaction_time` to power the peak-hour analysis
- `Customer_Occurrence_Key` — tracks how many times a customer has visited, enabling the visit frequency table

---

### 🏪 `store_lookup` — Store Dimension

Holds physical store metadata including `store_square_feet` — this enables the **Revenue per Square Foot** metric, which is the fairest way to compare stores of different sizes. A large store earning more absolute revenue may still be less efficient per sqft than a smaller one.

### 🛒 `product_lookup` — Product Dimension

Three flags make this table critical for promotional analysis:
- `promo_yn` — is this product currently on promotion?
- `new_product_yn` — is this a newly launched product?
- `tax_exempt_yn` — tax status affects margin calculations

Also contains `current_cost` vs `current_retail_price` — enabling gross margin analysis per product.

### 👤 `customer_lookup` — Customer Dimension

Contains `loyalty_card_number` — the bridge between anonymous transactions and identified customers. Customers without a loyalty card appear in sales data but cannot be tracked across visits, which is why the dashboard distinguishes **Identified vs Anonymous customers**.

### 👨‍🍳 `employee_lookup` — Staff Dimension

Contains `position` (e.g. Barista, Shift Manager) — enabling the staff performance table to show not just revenue per staff member but also their role context.

### 📅 `calendar` — Date Dimension

Includes a custom `Day Type` column — **Weekday vs Weekend** — which is one of the 6 global slicers on every page. This was one of the most operationally important dimensions: Daily Brew needed to know if staffing levels matched the demand pattern across day types.

### 🥐 `food_inventory` — Baked Goods Stock Table

Tracks `quantity_start_of_day` vs `quantity_sold` for baked goods — enabling waste and sell-through analysis for perishable items.

---

## 📊 Dashboard — 6 Pages in Detail

Every page shares the same **6 global slicers** in the header:
> `Promo Check` · `Store ID` · `Year` · `Month` · `Day Type`

This means every visual on every page responds to the same filters simultaneously — a supervisor can select "Store 3, Weekend, June" and the entire report recalculates for that context.

---

### 📄 Page 1 — Executive Overview

**Who uses it:** CEO, General Manager — big-picture health check in under 30 seconds.

**How it works:**

Two KPI card clusters sit at the top:

*Cluster 1 — Volume & Promo:*
- `Total Revenue` — SUM of all sales
- `Total Transactions` — COUNTROWS of fact table
- `AOV` (Average Order Value) — Total Revenue ÷ Total Transactions
- `Total Quantity Sold` — SUM of quantity_sold
- `Promo Revenue` — Revenue where promo_item_yn = "Y"
- `Promo % of Revenue` — Promo Revenue ÷ Total Revenue

*Cluster 2 — Growth Comparison:*
- `Revenue Prev Month` — previous month revenue using time intelligence
- `MoM Growth %` — (Current − Previous) ÷ Previous month revenue
- `Revenue LY` — same period last year using `SAMEPERIODLASTYEAR`
- `YoY Growth %` — year-on-year growth percentage

**The Revenue by Month line chart** plots `Revenue (Month)` as the primary line with `Total Transactions`, `Avg Order Value`, and `Total Quantity Sold` as secondary tooltip fields — hovering any data point shows the full context of that month's performance.

**Top Sold 5 Product Category bar chart** uses `product_lookup.product_category` against `Total Revenue` — giving immediate visibility into which product groups are driving the business.

---

### 📄 Page 2 — Store Performance

**Who uses it:** Operations Manager, Regional Director — comparing stores and spotting underperformers.

**How it works:**

KPI cards show:
- `Top Store Name` — dynamic label using `MAXX` + `RANKX` to always show the current #1 store given the active filter context
- `Top Store Revenue` — revenue of the top-ranked store
- `Average Revenue per Store` — total revenue ÷ distinct store count
- `Revenue per SqFt` — the efficiency metric: `Total Revenue ÷ SUM(store_square_feet)`
- `Transactions per Store` and `Revenue per Transaction`

**Store Revenue bar chart** — horizontal bars ranked by revenue, with tooltip showing `Total Transactions`, `Revenue/SqFt`, and `Avg Order Value` per store. This is the key visual for the "volume trap" question: a store might rank high on transactions but low on revenue per transaction.

**Store Wise Rev/SqFt column chart** — normalises revenue by store size. This often flips the ranking: a small neighbourhood store can outperform a large flagship on a per-sqft basis, which is a signal about location quality and customer density.

**Two peak-hour heatmaps:**
- `Heatmap of Peak Hour Transaction` — uses `time_band.hour_band` on the X axis, plotting transaction count across banded time periods
- `Heatmap of Peak Hour Revenue` — uses raw `sales_by_store.Hour` for higher granularity, showing which clock hours generate the most revenue

These two together answer: *"Are our busiest hours also our highest-revenue hours — or are we just getting lots of small orders at peak time?"*

---

### 📄 Page 3 — Product & Promotion Analysis

**Who uses it:** Product Manager, Marketing Team — understanding what sells and whether promotions work.

**How it works:**

KPI cards include:
- `% Revenue from New Products` — share of revenue from items where `new_product_yn = "Y"`
- `Average Items per Transaction` — basket size metric

**Promo vs Non-Promo Sales Trend column chart** — plots `Total Revenue` monthly with `Promo % of Revenue` as a secondary metric. The key question this answers: *when promo percentage goes up, does total revenue also go up — or does promo just cannibalise regular-price sales?* If total revenue is flat while promo % rises, promotions are reducing margin without growing volume.

**Product Category Contribution donut chart** — shows each `product_category`'s share of total revenue. Useful for spotting over-reliance on one category.

**Top Sold 5 Products bar chart** — ranks individual products (not categories) by total revenue. This is where hero products are identified.

**New vs Existing Products donut chart** — `product_lookup.new_product_yn` split of revenue. Tracks how well new product launches are penetrating the sales mix — a growing slice signals successful launches; a shrinking slice signals new products are being listed but not purchased.

---

### 📄 Page 4 — Customer & Staff Performance

**Who uses it:** HR, Store Managers — understanding who the valuable customers are and how staff are contributing.

**How it works:**

KPI cards show:
- `Total Customers (Genuine)` — DISTINCTCOUNT of customer_id where customer is identified (has loyalty card)
- `Returning Customer %` — % of customers who appear in more than one transaction period
- `Loyalty Revenue` — revenue attributable to loyalty card holders
- `Avg Revenue per Customer`

**Customer Segment donut chart** — `AnonymousCustomers` vs `IdentifiedCustomers` split. Anonymous customers are those who transact without a loyalty card — they're counted in revenue but can't be tracked or retargeted. The donut shows how much of the business is "dark" from a CRM perspective.

**Customer Growth Trend area chart** — plots `Total Customers` month-over-month to show whether the customer base is expanding or plateauing.

**Top Customers by Revenue table** — ranks customers by `Total Revenue` with their customer ID and loyalty card number. This is the VIP list — the customers worth personal outreach or exclusive offers.

**Customers Visit Frequency table** — uses `sales_by_store.Customer_Occurrence_Key` — a calculated column that tags each customer's nth visit. This reveals the distribution of visit frequency: what % of customers visit once, twice, 5+ times?

**Staff Productivity table** — shows each employee's `Total Revenue`, `Total Transactions`, and `Staff AOV` (average order value per staff member). AOV per staff is the most nuanced metric here — a barista with fewer transactions but higher AOV may be upselling more effectively than a high-volume colleague.

**Staff by Revenue bar chart** — the visual companion to the table, with `position` in the tooltip so managers can see if the revenue ranking correlates with seniority.

---

### 📄 Page 5 — Insights

**What it is:** A dedicated text-card page written inside Power BI summarising the key findings discovered through the data — not more charts, but the *meaning* behind the charts.

**How it works:**
- 6 text box visuals present written insights in bullet-card format
- No data fields — this is deliberate narrative, written by the analyst after exploring all four dashboard pages
- Covers: revenue trends, best/worst performing stores, product performance, promotional effectiveness, customer behaviour patterns, and staff productivity anomalies

---

### 📄 Page 6 — Recommendations

**What it is:** The strategic action layer — translating each insight into a concrete business recommendation.

**How it works:**
- 6 text box visuals structured as: *"Insight → Recommended Action → Expected Outcome"*
- Addresses: shift scheduling vs peak-hour gaps, store-level best-practice sharing, promotional targeting strategy, loyalty programme expansion, new product rollout approach, and staff training priorities
- This page is what distinguishes a BI report from a BI *analysis* — the dashboard shows what is happening; this page says what to do about it

---

## 📐 DAX Measures Architecture

All measures are stored in a dedicated `_all_measures` table — isolated from the data tables so they don't pollute any dimension's column list. This is a Power BI best practice for maintainability.

**Measure categories built:**

| 🏷️ Category | 📊 Measures |
|---|---|
| 💰 Revenue | `Total Revenue`, `Revenue (Month)`, `Store Revenue`, `Product Revenue`, `Promo Revenue`, `Loyalty Revenue` |
| 📈 Growth | `Revenue Prev Month`, `MoM Growth %`, `Revenue LY`, `YoY Growth %` |
| 🛒 Volume | `Total Transactions`, `Total Quantity Sold`, `Average Items per Transaction` |
| 💳 Basket | `AOV`, `Avg Order Value`, `Staff AOV`, `Revenue per Transaction` |
| 🏪 Store | `Top Store Name`, `Top Store Revenue`, `Average Revenue per Store`, `Revenue per SqFt`, `Revenue/SqFt`, `Transactions per Store` |
| 👥 Customer | `Total Customers (Genuine)`, `Returning Customer %`, `Avg Revenue per Customer`, `AnonymousCustomers`, `IdentifiedCustomers` |
| 🎯 Promo | `Promo % of Revenue`, `% Revenue from New Products` |

**Key DAX patterns used:**
- `CALCULATE` + `FILTER` — for promo and new-product sub-totals
- `SAMEPERIODLASTYEAR` — for YoY comparisons
- `DATEADD` — for MoM prior-period reference
- `DIVIDE(..., 0)` — safe division to handle stores/months with zero transactions
- `RANKX` + `MAXX` — for dynamic "Top Store" label that updates with every filter change
- `DISTINCTCOUNT` — for unique customer and store counts
- `SELECTEDVALUE` / `HASONEVALUE` — for context-aware card titles

---

## ⚙️ Key Power BI Design Decisions

| 🛠️ Decision | 🎯 Why |
|---|---|
| Isolated `_all_measures` table | Keeps DAX measures separate from data — cleaner field list, easier maintenance |
| `time_band` helper table | Groups raw hours into human-readable bands for non-technical stakeholders |
| `month_sort` helper table | Prevents months from sorting alphabetically (Apr, Aug, Dec...) on chart axes |
| `customer_summary` calculated table | Pre-aggregates customer visit counts so the visit frequency table doesn't time out on large data |
| 6 global slicers on every page | Consistent filter context — any selection applies everywhere, no page-level confusion |
| `Revenue/SqFt` as a separate metric | Normalises store size so a 500 sqft store can fairly be compared to a 2,000 sqft store |
| Anonymous vs Identified customer split | Quantifies the "dark" portion of the customer base — a key input for loyalty programme ROI |
| Separate Insights + Recommendations pages | Separates data (what happened) from analysis (what it means) and strategy (what to do) |

---

## 🛠️ Tech Stack

![Power BI](https://img.shields.io/badge/PowerBI-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![DAX](https://img.shields.io/badge/DAX-Measures-blue?style=for-the-badge)
![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)

| 🔧 Tool | 🎯 Used For |
|---|---|
| **Power BI Desktop** | Full dashboard — data model, DAX, visuals, navigation |
| **DAX** | All calculated measures and helper columns |
| **Power Query (M)** | Data transformation, type enforcement, helper table creation |
| **Excel / CSV** | Source data tables (imported into Power BI) |

---

## 📋 Assignment Questions — How Each Was Answered

| ❓ Question | 📊 Dashboard Answer |
|---|---|
| 📈 Monthly performance trend? | Revenue by Month line chart — Page 1 |
| 🏪 Which stores earn the most? | Store Revenue bar chart + Rev/SqFt — Page 2 |
| ☕ Top-selling products? | Top Sold 5 Products bar chart — Page 3 |
| 🎯 Are promo items boosting sales? | Promo vs Non-Promo Trend + Promo % KPI — Page 3 |
| 👥 Most valuable & loyal customers? | Top Customers table + Returning Customer % — Page 4 |
| 🕐 Peak sales hours? | Peak Hour Heatmaps (by transaction & revenue) — Page 2 |
| 👨‍🍳 Staff sales performance? | Staff Productivity table + Staff by Revenue chart — Page 4 |
| 🛒 Average basket size? | AOV + Average Items per Transaction KPIs — Pages 1 & 3 |
| 💡 Key insights & trends? | Dedicated Insights page — Page 5 |
| 🚀 Strategic recommendations? | Dedicated Recommendations page — Page 6 |

---

## 👤 Author

**Md. Shafat Hossain** — BI Analyst (Junior)
> 📌 *Power BI case study submission — Daily Brew Coffee Ltd. Business Intelligence Dashboard.*

---

## 🏫 Course Context

This project was submitted as a **Power BI Dashboard Case Study Assignment** for a Business Intelligence course. The brief simulated a real-world junior BI analyst role at a NYC coffee chain, requiring a 3–4 page interactive report with an insight and recommendation layer. The final submission delivers **6 pages** — exceeding the brief with dedicated Insights and Recommendations pages, a star schema data model with 11 tables/objects, and a full suite of time-intelligence DAX measures.

---

## 🔒 License

Academic project. Dataset is fictional and used for educational purposes only.
