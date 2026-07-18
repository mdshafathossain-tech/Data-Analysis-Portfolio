# 🛒 Omni Retail Pvt. Ltd. — Strategic Sales Analysis (FY 2024)
### End-to-End Data Analytics Project | SQL · Power BI · Excel · Python

> 🎯 **Purpose in one line:** A full-stack data analysis of Omni Retail's 2024 operations — from raw CSV to a cleaned database, 26 SQL views, a 5-page interactive Power BI dashboard, and a voiced presentation — uncovering the revenue drivers, return-rate risks, and customer retention patterns that shape the company's next strategic move.

---

## 📌 Executive Summary of Key Findings

| 📊 Metric | 💡 Result |
|---|---|
| 💰 Total Revenue (FY 2024) | **$497,663.79** |
| 🏆 Top Region | West — $143,823.03 |
| ⚠️ Critical Risk | Electronics return rate — **15.63%** |
| 📅 Peak Quarter | Q3 (highest total revenue) |
| 🛍️ Weekend vs. Weekday | Weekend shoppers spend **~$80 more** per transaction |
| 👥 Cohort Tracked | **186 unique customers** across retention cohorts |

---

## 📁 Project Files

| 📂 File | 🛠️ Tool | 📋 Purpose |
|---|---|---|
| `Omni_retail_Pvt_Ltd.sql` | PostgreSQL | Database schema + 26 analytical views |
| `Omni_retail_Pvt_Ltd.pbix` | Power BI | 5-page interactive dashboard |
| `Macro.xlsm` | Excel + VBA | Data cleaning automation with one-click macro button |
| `Presentation.mp4` | Video | Voiced walkthrough of findings (English, social-media ready) |
| `Question_Set.pdf` | Reference | 14 analytical questions across 3 difficulty levels |

---

## 🧠 How the Whole Project Works — Big Picture

This project follows a strict **raw data → clean data → analysis → visualisation → insight** pipeline. Each layer builds on the one before it.

```
┌─────────────────────────────────────────────────────────────┐
│  📥  RAW DATA                                               │
│  ecommerce_customer_behavior_dataset.csv                    │
│  (transactions: sale_id, customer, product, store,         │
│   region, quantity, price, payment, return flag)           │
└──────────────────────┬──────────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────────┐
│  🧹  STEP 1 — DATA CLEANING (Excel + VBA Macro)            │
│  Macro.xlsm                                                 │
│  - Removes duplicates & blank rows                         │
│  - Standardises date formats and column alignment          │
│  - Assigned to a button: one click = clean dataset         │
└──────────────────────┬──────────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────────┐
│  🗄️  STEP 2 — DATABASE LOADING (PostgreSQL)                 │
│  Omni_retail_Pvt_Ltd.sql                                    │
│  - Creates `sql_sales_project` database                    │
│  - Defines `sales_data` table with strict data types       │
│  - Loads cleaned CSV into the table                        │
└──────────────────────┬──────────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────────┐
│  🔍  STEP 3 — SQL ANALYSIS (26 Views)                       │
│  10 core business questions + 16 ad-hoc deep-dives         │
│  All saved as reusable CREATE OR REPLACE VIEW objects      │
└──────────────────────┬──────────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────────┐
│  📊  STEP 4 — POWER BI DASHBOARD (5 pages)                 │
│  Connected directly to PostgreSQL views                    │
│  Sales Trends · Store Performance · Category Diagnostics  │
│  Cohort Analysis · Payment Behaviour                       │
└──────────────────────┬──────────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────────┐
│  🎤  STEP 5 — VOICED PRESENTATION (MP4)                    │
│  Strategic recommendations recorded in English             │
│  Prepared for social media publication                     │
└─────────────────────────────────────────────────────────────┘
```

---

## 🗄️ Database Schema

The entire analysis runs on one central table: `sales_data`.

```sql
CREATE TABLE sales_data (
    sale_id          VARCHAR(10)     PRIMARY KEY,
    sale_date        DATE            NOT NULL,
    customer_id      VARCHAR(10)     NOT NULL,
    product_id       VARCHAR(10)     NOT NULL,
    product_category VARCHAR(50)     NOT NULL,
    store_id         VARCHAR(10)     NOT NULL,
    region           VARCHAR(20)     NOT NULL,   -- North / South / East / West
    quantity_sold    INTEGER         NOT NULL,
    unit_price       DECIMAL(10,2)   NOT NULL,
    payment_method   VARCHAR(30)     NOT NULL,
    returned         VARCHAR(3)      NOT NULL,   -- 'Yes' or 'No'
    total_sale_amount DECIMAL(12,2)  NOT NULL
);
```

**Why these design choices:**
- `returned` is stored as `VARCHAR('Yes'/'No')` instead of a boolean — matches the source CSV format and makes `CASE WHEN returned = 'Yes'` filters human-readable.
- `total_sale_amount` has 12-digit precision — wide enough for aggregate SUM operations without overflow.
- `sale_id` as `VARCHAR` rather than integer — preserves leading-zero prefixed IDs from the source system.

---

## 🔍 SQL Analysis — All 26 Views Explained

### 📦 Core Assignment Queries (10 Questions)

**1. 🌍 `highest_revenue_region`**
Groups all transactions by region, sums revenue, orders descending.
→ *Answers: Which region leads in total sales?*
→ **Finding: West Region — $143,823.03**

**2. 🏷️ `category_performance_avg`**
Calculates average sale value and transaction count per product category.
→ *Answers: Which category generates the most revenue per transaction on average?*

**3. 🔁 `category_return_rates`**
Uses `CASE WHEN returned = 'Yes' THEN 1 ELSE 0 END` inside `SUM()` to count returns, then divides by total transactions × 100 for the return rate percentage.
→ *Answers: Which categories have the worst return problem?*
→ **Finding: Electronics at 15.63% — the highest return rate**

**4. 📦 `top_5_products_by_quantity`**
Groups by `product_id` + `product_category`, sums quantity and revenue, limits to top 5.
→ *Answers: Which specific products are the volume leaders?*

**5. 🏪 `store_volume_vs_revenue` + `store_high_volume_low_revenue`**
The key metric here is `avg_rev_per_item = total_revenue / total_quantity_sold` — a store selling many cheap items scores low on this metric, identifying the "volume trap" stores that are busy but underperforming on value.
→ *Answers: Which store is working hardest for the least reward?*

**6. 💳 `revenue_by_payment_method`**
Uses a window function `SUM(...) OVER ()` to calculate each payment method's share of total revenue as a percentage — without needing a subquery.
→ *Answers: Does payment method correlate with spend level?*

**7. 👤 `top_spending_customers`**
Groups by `customer_id`, counts transactions and sums spend, orders by total spent descending.
→ *Answers: Who are the top 10 highest-value customers?*

**8. 📅 `sales_by_quarter` + `quarterly_performance_labels`**
Extracts quarter and year from `sale_date` using `EXTRACT(QUARTER FROM sale_date)`. The labels view maps 1→"First", 2→"Second" etc. for cleaner Power BI display.
→ *Answers: Which quarter peaks?*
→ **Finding: Q3 shows the highest revenue**

**9. 💲 `avg_unit_price_per_category`**
Returns average, min, and max unit price per category — gives full price range context, not just the average.
→ *Answers: How do price points differ across categories?*

**10. ♻️ `highest_return_categories`**
Identical logic to query 3 but ordered by return percentage descending with LIMIT 1 — extracts the single worst offender.
→ **Finding: Electronics confirmed as highest return category**

---

### 🔬 Ad-Hoc & Bonus Deep-Dive Queries (16 Views)

**💸 `revenue_lost_to_returns`** — "Money Pit" Analysis
Calculates `pct_revenue_lost` per category: how much of potential revenue is wiped out by returns. The formula `SUM(CASE WHEN returned='Yes' THEN total_sale_amount ELSE 0 END) / SUM(total_sale_amount)` gives the exact financial damage of the return problem per category.

**⚡ `store_efficiency`**
Computes `revenue_per_unit_sold` and `avg_items_per_transaction` per store. Separates high-revenue-per-unit stores (premium/efficient) from low-revenue-per-unit stores (discount/volume).

**💳 `payment_behavior_by_category`**
Cross-tabulates payment method against product category — reveals whether, for example, Electronics buyers prefer credit while Clothing buyers use cash. Used in the Power BI Category Diagnostics page.

**🗓️ `weekend_vs_weekday_sales`**
Uses `EXTRACT(DOW FROM sale_date) IN (0, 6)` to flag Saturday (6) and Sunday (0) as weekends. Compares average transaction value between the two day types.
→ **Finding: Weekend average transaction value is ~$80 higher than weekday**

**🔄 `order_cohort_sequence`** — The Foundation of Retention Analysis
This is the most technically complex view. It uses a CTE (`customer_first_purchase`) to find each customer's very first transaction date, then joins back to all their subsequent transactions to compute a `cohort_month` number:
```sql
cohort_month =
  (YEAR of transaction − YEAR of first purchase) × 12
  + (MONTH of transaction − MONTH of first purchase)
```
Month 0 = the customer's first-ever purchase. Month 1 = they came back one month later. This view is the base for the retention analysis.

**📊 `cohort_retention_rates`**
Built on top of `order_cohort_sequence`. Uses two CTEs:
- `cohort_counts` — counts unique customers in each cohort at Month 0 (their starting size)
- `retention_counts` — counts how many of those same customers returned in Month 1

Then LEFT JOINs them to compute the retention rate percentage for each cohort month.
→ **Tracked: 186 unique customers across cohorts**

**📈 `revenue_by_month`**
Uses `DATE_TRUNC('month', sale_date)::DATE` to group transactions into calendar months — powers the monthly revenue trend line in Power BI.

**🏪 `store_performance_summary` + `highest_avg_trans_value`**
Summary gives all stores ranked by total revenue. The second view isolates the single store with the highest average transaction value — useful for identifying the premium-customer store.

**💰 `overall_return_rate`**
A single-row aggregate view giving the factory-wide return rate across all categories and regions combined — the headline number for the executive summary.

**🛒 `top_10_products_by_quantity`**
Extended version of the top-5 view — used for the Power BI product ranking visual.

**🥇 `best_selling_category`** / **🔻 `worst_performing_returns`**
Two single-row views returning the #1 category by revenue and the #1 category by return rate respectively — used as KPI card data sources in Power BI.

**💳 `payment_method_counts`**
Uses `COUNT(sale_id) / SUM(COUNT(sale_id)) OVER ()` (window function) to compute each payment method's share of total transaction volume as a percentage.

---

## 📊 Power BI Dashboard — 5 Pages

The `.pbix` file connects directly to the PostgreSQL views above. Each page is built from a specific set of views.

| 📄 Page | 🔌 Source Views | 📈 Key Visuals |
|---|---|---|
| **1. Sales Trends** | `revenue_by_month`, `sales_by_quarter`, `weekend_vs_weekday_sales` | Monthly revenue line, quarterly bar, weekday/weekend comparison card |
| **2. Store Performance** | `store_performance_summary`, `store_efficiency`, `highest_avg_trans_value` | Store revenue rank, efficiency scatter, top-value store KPI card |
| **3. Category Diagnostics** | `category_return_rates`, `revenue_lost_to_returns`, `best_selling_category`, `worst_performing_returns` | Return rate bar chart, revenue lost waterfall, category revenue treemap |
| **4. Cohort Analysis** | `order_cohort_sequence`, `cohort_retention_rates` | Cohort retention heatmap, Month-0 vs Month-1 retention line |
| **5. Payment Behaviour** | `revenue_by_payment_method`, `payment_method_counts`, `payment_behavior_by_category` | Payment method donut, avg spend by method, category × method matrix |

---

## ❓ Questions Answered — All 14 Across 3 Levels

### 🟢 Level 1 — Basic Insights (10 Questions)

| # | Question | How Answered |
|---|---|---|
| Q1 | Mean, Median, Mode of customer Age | Descriptive stats in Excel / Python |
| Q2 | Variance, Std Dev, Z-score of Purchase Amount | `STDEV`, `VAR`, z-score formula |
| Q3 | Top 3 product categories by purchase count | `best_selling_category` + `category_performance_avg` |
| Q4 | How many return customers? | `top_spending_customers` + customer flag filter |
| Q5 | Average review score | `AVG(review_score)` aggregate |
| Q6 | Delivery time by subscription status (Free vs Premium) | `AVG(delivery_time)` grouped by subscription |
| Q7 | How many customers subscribed? | `COUNT` with subscription filter |
| Q8 | Device usage % (Mobile / Desktop / Tablet) | `payment_method_counts` pattern on device column |
| Q9 | Avg purchase amount: discount vs no discount | `AVG` with `CASE WHEN discount = 'Yes'` |
| Q10 | Most common payment method | `payment_method_counts` ORDER BY transaction_count DESC LIMIT 1 |

### 🟡 Level 2 — Intermediate Insights (5 Questions)

| # | Question | How Answered |
|---|---|---|
| Q1 | Avg review score of most common payment method users | Join of payment method filter + review score avg |
| Q2 | Correlation: time on website vs purchase amount | Pearson correlation coefficient; scatter plot in Power BI |
| Q3 | % of satisfied customers (rating 4–5) who are also return customers | Nested `COUNTIF` with dual condition |
| Q4 | Items purchased vs customer satisfaction relationship | Grouped avg satisfaction by items purchased bucket |
| Q5 | Location with 2nd highest average purchase amount | `RANK()` window function on location avg spend |

### 🔴 Level 3 — Critical Thinking (4 Questions)

| # | Question | Insight Delivered |
|---|---|---|
| Q1 | Factors driving return customer classification | Premium subscription, discount usage, higher avg spend, and specific payment methods are the strongest predictors |
| Q2 | Payment methods → satisfaction & return rates | `payment_behavior_by_category` cross-tab; certain methods correlate with higher return rates |
| Q3 | Location impact on purchase amount & delivery time | West region leads spend; delivery time variance is significant between regions |
| Q4 | Major insights summary | See Strategic Recommendations below |

---

## 🚀 5-Point Strategic Recommendations

Based on the data analysis, the following actions were proposed:

**1. 🔧 Fix the Electronics Return Crisis**
Electronics at 15.63% return rate is the highest-risk category. Recommended: enhanced product descriptions, pre-purchase Q&A, and post-purchase onboarding to reduce buyer's remorse.

**2. 🌍 Invest in Regional Underperformers**
West leads at $143K. The gap between West and the lowest-performing region represents direct revenue opportunity — targeted regional promotions or store-level restructuring.

**3. 💎 Protect High-Value Weekend Shoppers**
Weekend transactions average ~$80 more than weekday. These customers warrant premium loyalty incentives and weekend-exclusive offers.

**4. 🔄 Strengthen Month-1 Retention**
Cohort analysis shows significant customer drop-off after Month 0. A structured first-month re-engagement campaign (automated email, personalised offer) targeting the 186-customer tracked cohort can directly improve LTV.

**5. 🏪 Audit "Volume Trap" Stores**
Stores with high quantity sold but low `avg_rev_per_item` are absorbing operational cost without proportional revenue. Reassign these stores toward higher-margin product categories.

---

## ⚙️ Key SQL Techniques Used

| 🛠️ Technique | 📍 View | 🎯 Why |
|---|---|---|
| `CASE WHEN returned='Yes' THEN 1 ELSE 0 END` in `SUM()` | Return rate views | Count boolean-as-string flags without casting |
| `SUM(...) OVER ()` window function | `revenue_by_payment_method`, `payment_method_counts` | Compute % of total in a single pass — no subquery needed |
| CTE (`WITH ... AS`) | `order_cohort_sequence`, `cohort_retention_rates` | Multi-step logic without nested subqueries — readable and maintainable |
| `DATE_TRUNC('month', sale_date)::DATE` | `revenue_by_month`, cohort views | Snap all dates to month-start for clean grouping |
| `EXTRACT(DOW FROM sale_date) IN (0,6)` | `weekend_vs_weekday_sales` | Identify weekend days (0=Sunday, 6=Saturday in PostgreSQL) |
| `NULLIF(SUM(quantity), 0)` | `store_volume_vs_revenue` | Prevent division-by-zero on stores with no sales |
| `COALESCE(retained, 0)` | `cohort_retention_rates` | Replace NULL (no retention) with 0 for clean percentage calc |
| `CREATE OR REPLACE VIEW` | All 26 views | Reusable, version-safe query objects — Power BI connects directly |
| `ROUND(..., 2)` on all percentages | All rate views | Consistent 2-decimal display across dashboard |

---

## 🛠️ Tech Stack

![PostgreSQL](https://img.shields.io/badge/PostgreSQL-316192?style=for-the-badge&logo=postgresql&logoColor=white)
![Power BI](https://img.shields.io/badge/PowerBI-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)

| 🔧 Tool | 🎯 Used For |
|---|---|
| **PostgreSQL** | Database creation, schema design, 26 analytical views |
| **Power BI Desktop** | 5-page interactive dashboard connected to PostgreSQL |
| **Excel + VBA Macro** | Raw data cleaning automation |
| **Python** | Descriptive statistics (mean, median, mode, z-score) |
| **MP4 (Voiced PPT)** | Presentation recording for social media |

---

## 👤 Author

**Md. Shafat Hossain** — Data Analyst
> 📌 *Final project submission for the Data Analysis course — Omni Retail Pvt. Ltd. Strategic Analysis, FY 2024.*

---

## 🏫 Course Context

This project was submitted as a **Final Assignment** for a Data Analysis course. The instructor provided a retail transactions dataset and a 3-level question set (Basic → Intermediate → Critical Thinking). The submission goes beyond the required scope with cohort retention analysis, store efficiency audits, a voiced video presentation, and 5 strategic business recommendations.

---

## 🔒 License

Academic project. Dataset is fictional and used for educational purposes only.
