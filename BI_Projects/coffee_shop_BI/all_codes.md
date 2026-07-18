# 🧮 DAX Measures & Power Query Reference
## Daily Brew Coffee Ltd. — Power BI Code Documentation

> 📖 Complete reference of every measure, calculated column, calculated table, and Power Query (M) script used in `BI_Analysis_-_Daily_Brew_Coffee_Ltd.pbix` — extracted directly from the data model. Use this as a lookup guide while exploring the dashboard.

---

## 📑 Table of Contents

- [📐 Data Model Relationships](#-data-model-relationships)
- [🔢 DAX Measures — `_all_measures`](#-dax-measures----all_measures)
  - [💰 Revenue Measures](#-revenue-measures)
  - [📈 Time Intelligence (MoM & YoY)](#-time-intelligence-mom--yoy)
  - [🛒 Volume & Basket Measures](#-volume--basket-measures)
  - [🎯 Promotional Measures](#-promotional-measures)
  - [🏪 Store Performance Measures](#-store-performance-measures)
  - [🛍️ Product Measures](#️-product-measures)
  - [👥 Customer Measures](#-customer-measures)
  - [👨‍🍳 Staff Measures](#-staff-measures)
- [🧩 DAX Calculated Columns](#-dax-calculated-columns)
  - [📊 `sales_by_store` Calculated Columns](#-sales_by_store-calculated-columns)
- [⚡ Power Query (M Code) — All Tables](#-power-query-m-code---all-tables)
  - [📋 `sales_by_store`](#-sales_by_store)
  - [👤 `customer_lookup`](#-customer_lookup)
  - [🏪 `store_lookup`](#-store_lookup)
  - [🛒 `product_lookup`](#-product_lookup)
  - [👨‍🍳 `employee_lookup`](#-employee_lookup)
  - [🥐 `food_inventory`](#-food_inventory)
  - [📅 `calendar`](#-calendar)
  - [🕐 `time_band`](#-time_band)
  - [📆 `month_sort`](#-month_sort)
  - [🔢 `_all_measures` (empty anchor table)](#-_all_measures-empty-anchor-table)

---

## 📐 Data Model Relationships

All relationships are **Many-to-One (M:1)** — the fact table `sales_by_store` connects to every dimension.

| 🔗 From Table | 🔑 From Column | ➡️ To Table | 🔑 To Column | 🔄 Cross Filter |
|---|---|---|---|---|
| `sales_by_store` | `transaction_date` | `calendar` | `Date` | Both |
| `sales_by_store` | `customer_id` | `customer_lookup` | `customer_id` | Single |
| `sales_by_store` | `staff_id` | `employee_lookup` | `staff_id` | Single |
| `sales_by_store` | `store_id` | `store_lookup` | `store_id` | Single |
| `sales_by_store` | `product_id` | `product_lookup` | `product_id` | Single |
| `sales_by_store` | `hour_band` | `time_band` | `hour_band` | Both |
| `food_inventory` | `store_id` | `store_lookup` | `store_id` | Single |
| `food_inventory` | `product_id` | `product_lookup` | `product_id` | Single |
| `calendar` | `Short Month` | `month_sort` | `ShortMonth` | Both |

---

## 🔢 DAX Measures — `_all_measures`

> All measures live in a dedicated `_all_measures` table (an empty anchor table) — isolated from data tables to keep field lists clean.

---

### 💰 Revenue Measures

#### `Total Revenue`
> Core revenue measure — multiplies quantity by unit price per row, then sums. Does **not** rely on a pre-calculated column.
```dax
Total Revenue =
SUMX(
    'sales_by_store',
    'sales_by_store'[quantity_sold] * 'sales_by_store'[unit_price]
)
```

#### `Total Transactions`
> Counts distinct transactions using a composite key column (not raw `transaction_id`) to ensure uniqueness across all line items.
```dax
Total Transactions =
DISTINCTCOUNT('sales_by_store'[transaction_key])
```

#### `Avg Order Value`
> Average spend per transaction. Uses `DIVIDE` with a 0 fallback to avoid division errors when no transactions exist.
```dax
Avg Order Value =
DIVIDE([Total Revenue], [Total Transactions], 0)
```

#### `Revenue (Month)`
> Locks the calculation to the current month and year context — used in the monthly trend line chart so the visual always shows one bar per calendar month.
```dax
Revenue (Month) =
CALCULATE(
    [Total Revenue],
    VALUES('calendar'[Month Name]),
    VALUES('calendar'[Year])
)
```

#### `Store Revenue`
> Revenue scoped to individual store context. `ALLEXCEPT` removes all filters *except* the store_id — ensuring the measure respects store-level slicing without being affected by other dimensions.
```dax
Store Revenue =
CALCULATE([Total Revenue], ALLEXCEPT('store_lookup', 'store_lookup'[store_id]))
```

#### `Product Revenue`
> Revenue scoped per product. Same `ALLEXCEPT` pattern as Store Revenue, but scoped to `product_lookup`.
```dax
Product Revenue =
CALCULATE(
    [Total Revenue],
    ALLEXCEPT('product_lookup', 'product_lookup'[product])
)
```

#### `Loyalty Revenue`
> Revenue from customers who have a valid loyalty card — excludes known "anonymous" placeholder IDs (8600, 1000, 6000) that represent walk-in customers without loyalty accounts.
```dax
Loyalty Revenue =
CALCULATE(
    [Total Revenue],
    NOT 'customer_lookup'[loyalty_card_number]
        IN { "000-000-0003", "000-000-0005", "000-000-0005" }
)
```

---

### 📈 Time Intelligence (MoM & YoY)

#### `Revenue Prev Month`
> Prior month revenue using `DATEADD` to shift the date context back by 1 month.
```dax
Revenue Prev Month =
CALCULATE(
    [Total Revenue],
    DATEADD('Calendar'[Date], -1, MONTH)
)
```

#### `MoM Growth %`
> Month-over-Month growth rate. Uses `VAR` pattern to avoid recalculating the same measure twice. Returns `BLANK()` when there is no prior month data (e.g. the very first month in the dataset).
```dax
MoM Growth % =
VAR Prev = [Revenue Prev Month]
VAR Curr = [Total Revenue]
RETURN
IF(Prev = 0, BLANK(), (Curr - Prev) / Prev)
```

#### `Revenue LY`
> Same period last year revenue using the built-in `SAMEPERIODLASTYEAR` time intelligence function.
```dax
Revenue LY =
CALCULATE(
    [Total Revenue],
    SAMEPERIODLASTYEAR('Calendar'[Date])
)
```

#### `YoY Growth %`
> Year-over-Year growth rate. Same `VAR` safety pattern as MoM — returns `BLANK()` when no prior year data exists.
```dax
YoY Growth % =
VAR LY = [Revenue LY]
VAR Curr = [Total Revenue]
RETURN
IF(LY = 0, BLANK(), (Curr - LY) / LY)
```

---

### 🛒 Volume & Basket Measures

#### `Total Quantity Sold`
```dax
Total Quantity Sold =
SUM('sales_by_store'[quantity_sold])
```

#### `Avg Items/Transaction`
> Average basket size — total units divided by total distinct transactions.
```dax
Avg Items/Transaction =
DIVIDE(
    SUM('sales_by_store'[quantity_sold]),
    [Total Transactions],
    0
)
```

#### `Rev/Transaction`
```dax
Rev/Transaction =
DIVIDE([Total Revenue], [Total Transactions], 0)
```

---

### 🎯 Promotional Measures

#### `Promo Revenue`
> Revenue from transactions where the item was flagged as promotional.
```dax
Promo Revenue =
CALCULATE(
    [Total Revenue],
    'sales_by_store'[promo_item_yn] = "Y"
)
```

#### `Promo % of Revenue`
> Share of total revenue coming from promo items.
```dax
Promo % of Revenue =
DIVIDE([Promo Revenue], [Total Revenue], 0)
```

#### `Revenue per Transaction Promo`
> Average transaction value when the purchase includes at least one promo item.
```dax
Revenue per Transaction Promo =
VAR PromoTransRev =
    CALCULATE(
        [Total Revenue],
        'sales_by_store'[promo_item_yn] = "Y"
    )
VAR PromoTransCount =
    CALCULATE(
        DISTINCTCOUNT('sales_by_store'[transaction_id]),
        'sales_by_store'[promo_item_yn] = "Y"
    )
RETURN
DIVIDE(PromoTransRev, PromoTransCount, 0)
```

#### `Revenue per Transaction NonPromo`
> Average transaction value for non-promotional purchases — the baseline to compare against promo uplift.
```dax
Revenue per Transaction NonPromo =
VAR NPromoTransRev =
    CALCULATE(
        [Total Revenue],
        'sales_by_store'[promo_item_yn] = "N"
    )
VAR NPromoTransCount =
    CALCULATE(
        DISTINCTCOUNT('sales_by_store'[transaction_id]),
        'sales_by_store'[promo_item_yn] = "N"
    )
RETURN
DIVIDE(NPromoTransRev, NPromoTransCount, 0)
```

#### `Promo Uplift %`
> The percentage increase in transaction value that promotions generate compared to non-promo transactions. A positive value means promos increase basket size.
```dax
Promo Uplift % =
VAR P  = [Revenue per Transaction Promo]
VAR NP = [Revenue per Transaction NonPromo]
RETURN IF(NP = 0, BLANK(), (P - NP) / NP)
```

#### `New Product Revenue`
> Revenue from products flagged as newly launched.
```dax
New Product Revenue =
CALCULATE(
    [Total Revenue],
    'product_lookup'[new_product_yn] = "Y"
)
```

#### `% Rev from New Prod`
```dax
% Rev from New Prod =
DIVIDE([New Product Revenue], [Total Revenue], 0)
```

---

### 🏪 Store Performance Measures

#### `Store Count`
```dax
Store Count =
DISTINCTCOUNT('store_lookup'[store_id])
```

#### `Top Store Revenue`
> Dynamically finds the maximum store revenue given the current filter context — updates when slicers change.
```dax
Top Store Revenue =
MAXX(
    VALUES('store_lookup'[store_id]),
    [Store Revenue]
)
```

#### `Top Rev Gen Store`
> Returns the **name** (store_id) of the top-performing store. Uses `FILTER` to find the store whose revenue equals the maximum, then `MAX` to return its ID as a label.
```dax
Top Rev Gen Store =
CALCULATE(
    MAX('store_lookup'[store_id]),
    FILTER(
        VALUES('store_lookup'[store_id]),
        [Store Revenue] = [Top Store Revenue]
    )
)
```

#### `Avg Rev/Store`
```dax
Avg Rev/Store =
DIVIDE([Total Revenue], [Store Count], 0)
```

#### `Revenue/SqFt`
> Normalises revenue by store physical size — the fairest cross-store comparison metric.
```dax
Revenue/SqFt =
DIVIDE(
    [Total Revenue],
    SUM('store_lookup'[store_square_feet]),
    0
)
```

#### `Transactions/Store`
> Average number of transactions per store using `AVERAGEX` to iterate over each store individually.
```dax
Transactions/Store =
AVERAGEX(
    VALUES('store_lookup'[store_id]),
    CALCULATE([Total Transactions])
)
```

---

### 👥 Customer Measures

#### `Total Customers`
> Combines identified customers (with loyalty cards) and anonymous walk-ins (without). Anonymous customers are tracked via a unique occurrence key rather than customer_id since their IDs are placeholder values (8600, 1000, 6000).
```dax
Total Customers =
VAR IdentifiedCustomers =
    CALCULATE(
        DISTINCTCOUNT('sales_by_store'[customer_id]),
        NOT 'sales_by_store'[customer_id] IN { 8600, 1000, 6000 }
    )

VAR AnonymousCustomers =
    CALCULATE(
        DISTINCTCOUNT('sales_by_store'[Customer_Occurrence_Key]),
        'sales_by_store'[customer_id] IN { 8600, 1000, 6000 }
    )

RETURN
IdentifiedCustomers + AnonymousCustomers
```

#### `IdentifiedCustomers`
> Only customers with real loyalty card accounts.
```dax
IdentifiedCustomers =
    CALCULATE(
        DISTINCTCOUNT('sales_by_store'[customer_id]),
        NOT 'sales_by_store'[customer_id] IN { 8600, 1000, 6000 }
    )
```

#### `AnonymousCustomers`
> Walk-in customers without loyalty accounts — tracked by unique occurrence key.
```dax
AnonymousCustomers =
    CALCULATE(
        DISTINCTCOUNT('sales_by_store'[Customer_Occurrence_Key]),
        'sales_by_store'[customer_id] IN { 8600, 1000, 6000 }
    )
```

#### `Returning Customers`
> Counts customers who have made more than one transaction. Uses `SUMMARIZE` to build a per-customer transaction count, then `FILTER` to keep only those with more than 1 visit. Anonymous placeholder IDs excluded.
```dax
Returning Customers =
COUNTROWS(
    FILTER(
        SUMMARIZE(
            'sales_by_store',
            'sales_by_store'[customer_id],
            "TxnCount", [Total Transactions]
        ),
        [TxnCount] > 1
            && NOT 'sales_by_store'[customer_id] IN { 8600, 1000, 6000 }
    )
)
```

#### `Returning Customer %`
```dax
Returning Customer % =
DIVIDE([Returning Customers], [Total Customers], 0)
```

#### `Avg Rev/Customer`
```dax
Avg Rev/Customer =
DIVIDE([Total Revenue], [Total Customers], 0)
```

#### `Customer Revenue`
> Revenue scoped to individual customer context using `ALLEXCEPT`.
```dax
Customer Revenue =
CALCULATE(
    [Total Revenue],
    ALLEXCEPT('customer_lookup', 'customer_lookup'[customer_id])
)
```

#### `Customer Visit Count`
> Total distinct transactions per customer across all time. Used to classify customers as One-time / Occasional / Loyal.
```dax
Customer Visit Count =
CALCULATE(
    DISTINCTCOUNT('sales_by_store'[transaction_id]),
    ALLEXCEPT('sales_by_store', 'sales_by_store'[customer_id])
)
```

#### `Customer Segment`
> Classifies each customer into a behavioural segment based on their visit count.
```dax
Customer Segment =
SWITCH(
    TRUE(),
    [Customer Visit Count] = 1,      "One-time",
    [Customer Visit Count] <= 5,     "Occasional",
    "Loyal"
)
```

#### `Customer Lifetime Value`
> Monthly revenue rate per customer — total revenue divided by the number of months since their first purchase.
```dax
Customer Lifetime Value =
[Customer Revenue] /
DATEDIFF(
    MIN('customer_lookup'[customer_since]),
    MAX('calendar'[Date]),
    MONTH
)
```

#### `Loyalty Revenue %`
```dax
Loyalty Revenue % =
DIVIDE([Loyalty Revenue], [Total Revenue], 0)
```

---

### 👨‍🍳 Staff Measures

#### `Staff Revenue`
> Revenue attributed to each staff member using `ALLEXCEPT` to scope per employee.
```dax
Staff Revenue =
CALCULATE(
    [Total Revenue],
    ALLEXCEPT('employee_lookup', 'employee_lookup'[staff_id])
)
```

#### `Staff Transactions`
```dax
Staff Transactions =
CALCULATE(
    [Total Transactions],
    ALLEXCEPT('employee_lookup', 'employee_lookup'[staff_id])
)
```

#### `Staff AOV`
> Average order value per staff member — a measure of upselling effectiveness.
```dax
Staff AOV =
DIVIDE([Total Revenue], [Total Transactions], 0)
```

---

## 🧩 DAX Calculated Columns

### 📊 `sales_by_store` Calculated Columns

#### `hour`
> Extracts the hour integer (0–23) from the raw transaction time. Used as the X axis on the peak-hour revenue heatmap.
```dax
hour =
HOUR('sales_by_store'[transaction_time])
```

#### `hour_band`
> Groups raw hours into 3-hour buckets for the summarised peak-hour transaction heatmap. Uses `SWITCH(TRUE(), ...)` — the most readable pattern for range-based classification in DAX.
```dax
hour_band =
SWITCH(
    TRUE(),
    'sales_by_store'[hour] >= 0  && 'sales_by_store'[hour] <= 3,  "00–03",
    'sales_by_store'[hour] >= 4  && 'sales_by_store'[hour] <= 7,  "04–07",
    'sales_by_store'[hour] >= 8  && 'sales_by_store'[hour] <= 11, "08–11",
    'Sales_by_store'[hour] >= 12 && 'sales_by_store'[hour] <= 15, "12–15",
    'Sales_by_store'[hour] >= 16 && 'sales_by_store'[hour] <= 19, "16–19",
    'Sales_by_store'[hour] >= 20 && 'sales_by_store'[hour] <= 23, "20–23"
)
```

#### `transaction_key`
> A composite surrogate key that uniquely identifies each transaction line — concatenates transaction_id, date, time, store, staff, and customer. Used by `DISTINCTCOUNT` in `Total Transactions` to ensure accurate counting.
```dax
transaction_key =
'sales_by_store'[transaction_id] & "|" &
FORMAT('sales_by_store'[transaction_date], "YYYYMMDD") & "|" &
FORMAT('sales_by_store'[transaction_time], "HHMMSS") & "|" &
'sales_by_store'[store_id] & "|" &
'sales_by_store'[staff_id] & "|" &
'sales_by_store'[customer_id]
```

#### `Customer_Occurrence_Key`
> A unique key per anonymous customer visit — used to count anonymous walk-ins who all share the same placeholder customer_id. Without this, all anonymous visits would collapse to a single customer.
```dax
Customer_Occurrence_Key =
'sales_by_store'[customer_id] & "-"
& FORMAT('sales_by_store'[transaction_date], "YYYYMMDD") & "-"
& FORMAT('sales_by_store'[transaction_time], "HHMMSS") & "-"
& 'sales_by_store'[store_id] & "-"
& 'sales_by_store'[staff_id]
```

#### `customer1`
> An alternative transaction-level key scoped to customer only (no store/staff) — used in specific customer journey analysis contexts.
```dax
customer1 =
'sales_by_store'[transaction_id] & "|" &
FORMAT('sales_by_store'[transaction_date], "YYYYMMDD") & "|" &
FORMAT('sales_by_store'[transaction_time], "HHMMSS") & "|" &
'sales_by_store'[customer_id]
```

---

## ⚡ Power Query (M Code) — All Tables

> All 6 data tables are loaded from **Google Drive** via direct-download URLs. Power Query enforces strict data types on every column during load.

---

### 📋 `sales_by_store`
> The main fact table. 13 columns, type-enforced on load.
```m
let
    Source = Csv.Document(
        Web.Contents("https://drive.google.com/uc?export=download&id=1Af-KXR6LDm1ZW8gE7zjunjiBy39Nd3ro"),
        [Delimiter=",", Columns=13, Encoding=1252, QuoteStyle=QuoteStyle.None]
    ),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"transaction_id",   Int64.Type},
        {"transaction_date", type date},
        {"transaction_time", type time},
        {"store_id",         Int64.Type},
        {"staff_id",         Int64.Type},
        {"customer_id",      Int64.Type},
        {"instore_yn",       type text},
        {"order",            Int64.Type},
        {"line_item_id",     Int64.Type},
        {"product_id",       Int64.Type},
        {"quantity_sold",    Int64.Type},
        {"unit_price",       type number},
        {"promo_item_yn",    type text}
    })
in
    #"Changed Type"
```

---

### 👤 `customer_lookup`
> 9-column customer dimension with loyalty card and demographic data.
```m
let
    Source = Csv.Document(
        Web.Contents("https://drive.google.com/uc?export=download&id=1fgLC6vCbe7_elqHdiOcNULVoyq4BeelJ"),
        [Delimiter=",", Columns=9, Encoding=1252, QuoteStyle=QuoteStyle.None]
    ),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"customer_id",         Int64.Type},
        {"home_store",          Int64.Type},
        {"customer_first-name", type text},
        {"customer_email",      type text},
        {"customer_since",      type date},
        {"loyalty_card_number", type text},
        {"birthdate",           type date},
        {"gender",              type text},
        {"birth_year",          Int64.Type}
    })
in
    #"Changed Type"
```

---

### 🏪 `store_lookup`
> 11-column store dimension including coordinates for map visuals.
```m
let
    Source = Csv.Document(
        Web.Contents("https://drive.google.com/uc?export=download&id=1mdz6b3Zc1eDzrTxB5tdOlYodMK6x0p6B"),
        [Delimiter=",", Columns=11, Encoding=1252, QuoteStyle=QuoteStyle.None]
    ),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"store_id",            type text},
        {"store_type",          type text},
        {"store_square_feet",   Int64.Type},
        {"store_address",       type text},
        {"store_city",          type text},
        {"store_state_province",type text},
        {"store_postal_code",   Int64.Type},
        {"store_longitude",     type number},
        {"store_latitude",      type number},
        {"manager",             Int64.Type},
        {"Neighorhood",         type text}
    })
in
    #"Changed Type"
```

---

### 🛒 `product_lookup`
> 14-column product dimension with cost, price, and three Y/N flag columns (promo, new, tax-exempt).
```m
let
    Source = Csv.Document(
        Web.Contents("https://drive.google.com/uc?export=download&id=1oUUc6PGad9NYcLUuAlaAv3vaNE7XbJpE"),
        [Delimiter=",", Columns=14, Encoding=1252, QuoteStyle=QuoteStyle.None]
    ),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"product_id",              Int64.Type},
        {"product_group",           type text},
        {"product_category",        type text},
        {"product_type",            type text},
        {"product",                 type text},
        {"product_description",     type text},
        {"unit_of_measure",         type text},
        {"current_cost",            type number},
        {"current_wholesale_price", type number},
        {"current_retail_price",    type number},
        {"tax_exempt_yn",           type text},
        {"promo_yn",                type text},
        {"new_product_yn",          type text},
        {"",                        Int64.Type}
    })
in
    #"Changed Type"
```

---

### 👨‍🍳 `employee_lookup`
> 6-column staff dimension with position and start date.
```m
let
    Source = Csv.Document(
        Web.Contents("https://drive.google.com/uc?export=download&id=1nQM9GFJp9n1K50IHEok-N_b6K3AhDMoN"),
        [Delimiter=",", Columns=6, Encoding=1252, QuoteStyle=QuoteStyle.None]
    ),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"staff_id",   Int64.Type},
        {"first_name", type text},
        {"last_name",  type text},
        {"position",   type text},
        {"start_date", type date},
        {"location",   type text}
    })
in
    #"Changed Type"
```

---

### 🥐 `food_inventory`
> 6-column baked goods inventory table. Uses UTF-8 encoding (65001) unlike other tables which use Windows-1252.
```m
let
    Source = Csv.Document(
        Web.Contents("https://drive.google.com/uc?export=download&id=1xlVxBhB-lQm72JP4XGuetIUZ3emC6uWF"),
        [Delimiter=",", Columns=6, Encoding=65001, QuoteStyle=QuoteStyle.None]
    ),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"store_id",             Int64.Type},
        {"baked_date",           type date},
        {"transaction_date",     type date},
        {"product_id",           Int64.Type},
        {"quantity_start_of_day",Int64.Type},
        {"quantity_sold",        Int64.Type}
    })
in
    #"Changed Type"
```

---

### 📅 `calendar`
> The most complex Power Query script — built entirely from the `sales_by_store` transaction dates rather than a static table. Generates 14 calendar columns dynamically.
```m
let
    // Load the same source file as sales_by_store to extract dates dynamically
    Source = Csv.Document(
        Web.Contents("https://drive.google.com/uc?export=download&id=1Af-KXR6LDm1ZW8gE7zjunjiBy39Nd3ro"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.None]
    ),

    // Promote header row
    #"Promoted Headers"     = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),

    // Keep ONLY the transaction_date column
    #"Keep Date Column"     = Table.SelectColumns(#"Promoted Headers", {"transaction_date"}),
    #"Changed Type"         = Table.TransformColumnTypes(#"Keep Date Column", {{"transaction_date", type date}}),

    // Rename to Date for time intelligence compatibility
    #"Renamed Column"       = Table.RenameColumns(#"Changed Type", {{"transaction_date", "Date"}}),
    #"Removed Duplicates"   = Table.Distinct(#"Renamed Column"),

    // ── Add calendar columns ──────────────────────────────────────────────
    #"Inserted Year"        = Table.AddColumn(#"Removed Duplicates",   "Year",          each Date.Year([Date]),          Int64.Type),
    #"Inserted Quarter"     = Table.AddColumn(#"Inserted Year",        "Quarter",       each Date.QuarterOfYear([Date]), Int64.Type),
    #"Inserted Month"       = Table.AddColumn(#"Inserted Quarter",     "Month",         each Date.Month([Date]),         Int64.Type),
    #"Inserted Month Name"  = Table.AddColumn(#"Inserted Month",       "Month Name",    each Date.MonthName([Date]),     type text),

    // Short month abbreviation (first 3 chars)
    #"Duplicated Column1"          = Table.DuplicateColumn(#"Inserted Month Name", "Month Name", "Month Name - Copy"),
    #"Renamed Columns1"            = Table.RenameColumns(#"Duplicated Column1", {{"Month Name - Copy", "Short Month"}}),
    #"Extracted First Characters"  = Table.TransformColumns(#"Renamed Columns1", {{"Short Month", each Text.Start(_, 3), type text}}),

    #"Inserted Start of Month"     = Table.AddColumn(#"Extracted First Characters", "Start of the Month", each Date.StartOfMonth([Date]),   type date),
    #"Inserted Week of Year"       = Table.AddColumn(#"Inserted Start of Month",    "Week of Year",       each Date.WeekOfYear([Date]),     Int64.Type),
    #"Inserted Week of Month"      = Table.AddColumn(#"Inserted Week of Year",      "Week of Month",      each Date.WeekOfMonth([Date]),    Int64.Type),
    #"Inserted Day Name"           = Table.AddColumn(#"Inserted Week of Month",     "Day Name",           each Date.DayOfWeekName([Date]),  type text),

    // Weekday / Weekend flag — powers the Day Type slicer on all pages
    #"Added Day Type"       = Table.AddColumn(#"Inserted Day Name", "Day Type",
        each if [Day Name] = "Saturday" or [Day Name] = "Sunday"
             then "Weekend" else "Weekday",
        type text),

    // Quarter label: 1 → "Q1", 2 → "Q2" etc.
    #"Quarter Text"         = Table.TransformColumns(#"Added Day Type",
        {{"Quarter", each "Q" & Text.From(_, "en-US"), type text}}),

    // Year-Quarter composite: "2023-Q1"
    #"Year-Quarter"         = Table.AddColumn(#"Quarter Text", "Year-Quarter",
        each Text.Combine({Text.From([Year], "en-US"), [Quarter]}, "-"),
        type text),

    // Short Year: "2023" → "23"
    #"Short Year"           = Table.AddColumn(#"Year-Quarter", "Short Year",
        each Text.End(Text.From([Year], "en-US"), 2),
        type text),

    // Mon-Year label: "January-23"
    #"Mon-Year"             = Table.AddColumn(#"Short Year", "Mon-Year",
        each [Month Name] & "-" & [Short Year],
        type text),

    // Final reorder and type enforcement
    #"Duplicated Column"    = Table.DuplicateColumn(#"Mon-Year", "Date", "Date - Copy"),
    #"Reordered Columns"    = Table.ReorderColumns(#"Duplicated Column", {
        "Date", "Date - Copy", "Year", "Quarter", "Month", "Month Name",
        "Start of the Month", "Week of Year", "Week of Month",
        "Day Name", "Day Type", "Year-Quarter", "Short Year", "Mon-Year"
    }),
    #"Renamed Columns"      = Table.RenameColumns(#"Reordered Columns", {{"Date - Copy", "Date_Value"}}),
    #"Changed Type1"        = Table.TransformColumnTypes(#"Renamed Columns", {
        {"Date_Value",         Int64.Type},
        {"Start of the Month", type date}
    })
in
    #"Changed Type1"
```

---

### 🕐 `time_band`
> A small static helper table mapping hour bands to sort-order index values. Stored as a compressed inline Base64 payload — no external file needed.
```m
let
    Source = Table.FromRows(
        Json.Document(Binary.Decompress(
            Binary.FromText(
                "HcyxFcAgCAXAXX5tAShqZuG5S3Zww0wSPuU1FwFFg8j3Xuk4LWD0oFe50zutWh5pNdrLTk/6Kc+08bP8zg8=",
                BinaryEncoding.Base64
            ),
            Compression.Deflate
        )),
        let _t = ((type nullable text) meta [Serialized.Text = true])
        in type table [Index = _t, #"Hour Band" = _t]
    ),
    #"Changed Type"    = Table.TransformColumnTypes(Source, {{"Index", Int64.Type}, {"Hour Band", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type", {{"Hour Band", "hour_band"}})
in
    #"Renamed Columns"
```

---

### 📆 `month_sort`
> A static 2-column table (Index 1–12, ShortMonth Jan–Dec) that gives months a numeric sort order so charts display Jan→Dec rather than alphabetically.
```m
let
    Source = Table.FromRows(
        Json.Document(Binary.Decompress(
            Binary.FromText(
                "i45WMlTSUfJKzFOK1YlWMgKy3VKTwGxjINs3sQjMNgGyHQsgbFOweCWYbQbSWwrRaw5m54DZFiD1pelgtiWQHZxaAGYbGgA5/sklEA7IZr/8MggHZLVLarJSbCwA",
                BinaryEncoding.Base64
            ),
            Compression.Deflate
        )),
        let _t = ((type nullable text) meta [Serialized.Text = true])
        in type table [Index = _t, ShortMonth = _t]
    ),
    #"Changed Type" = Table.TransformColumnTypes(Source, {{"Index", Int64.Type}, {"ShortMonth", type text}})
in
    #"Changed Type"
```

---

### 🔢 `_all_measures` (Empty Anchor Table)
> An intentionally empty single-column table used only as a container for DAX measures. Isolating all measures here keeps them out of dimension field lists, making the model cleaner.
```m
let
    Source = Table.FromRows(
        Json.Document(Binary.Decompress(
            Binary.FromText("i44FAA==", BinaryEncoding.Base64),
            Compression.Deflate
        )),
        let _t = ((type nullable text) meta [Serialized.Text = true])
        in type table [Column1 = _t]
    ),
    #"Changed Type" = Table.TransformColumnTypes(Source, {{"Column1", type text}})
in
    #"Changed Type"
```

---

## 📊 Quick Reference — Measure Count Summary

| 🏷️ Category | 🔢 Count |
|---|---|
| 💰 Revenue | 6 |
| 📈 Time Intelligence | 4 |
| 🛒 Volume & Basket | 3 |
| 🎯 Promotional | 6 |
| 🏪 Store Performance | 6 |
| 🛍️ Product | 2 |
| 👥 Customer | 8 |
| 👨‍🍳 Staff | 3 |
| **Total Measures** | **38** |
| 🧩 Calculated Columns | 5 |
| ⚡ Power Query Scripts | 10 |

---

## 👤 Author

**Md. Shafat Hossain** — BI Analyst
> 📌 *Full DAX and Power Query code reference for the Daily Brew Coffee Ltd. Power BI dashboard.*
