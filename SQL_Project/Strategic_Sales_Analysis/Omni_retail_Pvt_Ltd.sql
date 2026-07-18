create database sql_sales_project;

DROP TABLE IF EXISTS "sales_data" CASCADE;

CREATE TABLE sales_data (
    sale_id VARCHAR(10) NOT NULL PRIMARY KEY,
    sale_date DATE NOT NULL,
    customer_id VARCHAR(10) NOT NULL,
    product_id VARCHAR(10) NOT NULL,
    product_category VARCHAR(50) NOT NULL,
    store_id VARCHAR(10) NOT NULL,
    region VARCHAR(20) NOT NULL,
    quantity_sold INTEGER NOT NULL,
    unit_price DECIMAL(10, 2) NOT NULL,
    payment_method VARCHAR(30) NOT NULL,
    returned VARCHAR(3) NOT NULL, -- Stores 'Yes' or 'No'
    total_sale_amount DECIMAL(12, 2) NOT NULL
);

SELECT current_user;

select * from sales_data;

-- 1. Which region has the highest total revenue?

CREATE OR REPLACE VIEW highest_revenue_region AS
SELECT 
    region, 
    SUM(total_sale_amount) AS total_revenue
FROM 
    sales_data
GROUP BY 
    region
ORDER BY 
    total_revenue DESC;

SELECT * FROM highest_revenue_region LIMIT 1;

-- 2.1 Which product category generates the highest revenue on average per sale?

CREATE OR REPLACE VIEW category_performance_avg AS
SELECT 
    product_category, 
    AVG(total_sale_amount) AS avg_sale_value,
    COUNT(sale_id) AS transaction_count
FROM 
    sales_data
GROUP BY 
    product_category;

-- 2.2 Which product category generates the highest revenue on average per sale?

SELECT * FROM category_performance_avg LIMIT 1;

-- 3.1. What is the return rate per product category?

CREATE OR REPLACE VIEW category_return_rates AS
SELECT 
    product_category,
    COUNT(*) AS total_transactions,
    SUM(CASE WHEN returned = 'Yes' THEN 1 ELSE 0 END) AS returned_count,
    ROUND(
        100.0 * SUM(CASE WHEN returned = 'Yes' THEN 1 ELSE 0 END) / COUNT(*), 
        2
    ) AS return_rate_percentage
FROM 
    sales_data
GROUP BY 
    product_category;

SELECT * FROM category_return_rates ORDER BY return_rate_percentage DESC;

-- 3.2. What is the allover return rate?

CREATE OR REPLACE VIEW overall_return_rate AS
SELECT 
    COUNT(*) AS total_transactions,
    SUM(CASE WHEN returned = 'Yes' THEN 1 ELSE 0 END) AS total_returns,
    ROUND(
        100.0 * SUM(CASE WHEN returned = 'Yes' THEN 1 ELSE 0 END) / COUNT(*), 
        2
    ) AS return_rate_pct
FROM 
    sales_data;

-- 4. Identify the top 5 products with the highest total sales by quantity.

CREATE OR REPLACE VIEW top_5_products_by_quantity AS
SELECT 
    product_id, 
    product_category,
    SUM(quantity_sold) AS total_quantity_sold,
    SUM(total_sale_amount) AS total_revenue
FROM 
    sales_data
GROUP BY 
    product_id, 
    product_category
ORDER BY 
    total_quantity_sold DESC
LIMIT 5;

SELECT * FROM top_5_products_by_quantity;

-- 5. Which store has the lowest revenue but highest number of sales?
-- all list

CREATE OR REPLACE VIEW store_volume_vs_revenue AS
SELECT 
    store_id,
    SUM(quantity_sold) AS total_items_sold,
    SUM(total_sale_amount) AS total_revenue,
    ROUND(SUM(total_sale_amount) / NULLIF(SUM(quantity_sold), 0), 2) AS avg_rev_per_item
FROM 
    sales_data
GROUP BY 
    store_id
ORDER BY 
    avg_rev_per_item ASC;

-- store with lowest revenue but highest number of sales

CREATE OR REPLACE VIEW store_high_volume_low_revenue AS
SELECT 
    store_id,
    SUM(quantity_sold) AS total_items_sold,
    SUM(total_sale_amount) AS total_revenue,
    ROUND(SUM(total_sale_amount) / NULLIF(SUM(quantity_sold), 0), 2) AS avg_rev_per_item
FROM 
    sales_data
GROUP BY 
    store_id
ORDER BY 
    avg_rev_per_item ASC
	LIMIT 1;

SELECT * FROM store_volume_vs_revenue LIMIT 1;

-- 6. How do different payment methods impact total revenue?

CREATE OR REPLACE VIEW revenue_by_payment_method AS
SELECT 
    payment_method,
    COUNT(*) AS total_transactions,
    SUM(total_sale_amount) AS total_revenue,
    ROUND(AVG(total_sale_amount), 2) AS average_ticket_size,
    ROUND(
        100.0 * SUM(total_sale_amount) / SUM(SUM(total_sale_amount)) OVER (), 
        2
    ) AS percentage_of_total_revenue
FROM 
    sales_data
GROUP BY 
    payment_method
ORDER BY 
    total_revenue DESC;

-- 7. Which customers have made the most purchases in terms of total amount spent?

CREATE OR REPLACE VIEW top_spending_customers AS
SELECT 
    customer_id, 
    COUNT(sale_id) AS transaction_count,
    SUM(total_sale_amount) AS total_spent
FROM 
    sales_data
GROUP BY 
    customer_id
ORDER BY 
    total_spent DESC;

SELECT * FROM top_spending_customers LIMIT 10; -- top ten customers

-- 8. Which quarter sees the highest sales?

CREATE OR REPLACE VIEW sales_by_quarter AS
SELECT 
    EXTRACT(QUARTER FROM sale_date) AS sale_quarter,
    EXTRACT(YEAR FROM sale_date) AS sale_year,
    SUM(total_sale_amount) AS total_revenue,
    COUNT(sale_id) AS transaction_count
FROM 
    sales_data
GROUP BY 
    sale_year, 
    sale_quarter
ORDER BY 
    total_revenue DESC;

SELECT * FROM sales_by_quarter LIMIT 1;

-- 9. What is the average unit price per product category?

CREATE OR REPLACE VIEW avg_unit_price_per_category AS
SELECT 
    product_category,
    ROUND(AVG(unit_price), 2) AS average_item_price,
    MIN(unit_price) AS cheapest_item,
    MAX(unit_price) AS most_expensive_item
FROM 
    sales_data
GROUP BY 
    product_category
ORDER BY 
    average_item_price DESC;

-- 10. Which product categories have the highest return percentage?

CREATE OR REPLACE VIEW highest_return_categories AS
SELECT 
    product_category,
    COUNT(*) AS total_sales_count,
    SUM(CASE WHEN returned = 'Yes' THEN 1 ELSE 0 END) AS total_returns,
    ROUND(
        100.0 * SUM(CASE WHEN returned = 'Yes' THEN 1 ELSE 0 END) / COUNT(*), 
        2
    ) AS return_percentage
FROM 
    sales_data
GROUP BY 
    product_category
ORDER BY 
    return_percentage DESC;

SELECT * FROM highest_return_categories LIMIT 1;

-- High-Return "Money Pit" Analysis

CREATE OR REPLACE VIEW revenue_lost_to_returns AS
SELECT 
    product_category,
    SUM(total_sale_amount) AS potential_revenue,
    SUM(CASE WHEN returned = 'Yes' THEN total_sale_amount ELSE 0 END) AS revenue_lost,
    ROUND(100.0 * SUM(CASE WHEN returned = 'Yes' THEN total_sale_amount ELSE 0 END) / 
          SUM(total_sale_amount), 2) AS pct_revenue_lost
FROM sales_data
GROUP BY product_category
ORDER BY revenue_lost DESC;

-- Store Efficiency (Revenue per Item)

CREATE OR REPLACE VIEW store_efficiency AS
SELECT 
    store_id,
    SUM(total_sale_amount) / SUM(quantity_sold) AS revenue_per_unit_sold,
    AVG(quantity_sold) AS avg_items_per_transaction
FROM sales_data
GROUP BY store_id
ORDER BY revenue_per_unit_sold DESC;

-- Payment Method vs. Product Category

CREATE OR REPLACE VIEW payment_behavior_by_category AS
SELECT 
    product_category,
    payment_method,
    COUNT(*) AS transaction_count,
    AVG(total_sale_amount) AS avg_spend
FROM sales_data
GROUP BY product_category, payment_method
ORDER BY product_category, transaction_count DESC;

-- Weekend vs. Weekday Performance

CREATE OR REPLACE VIEW weekend_vs_weekday_sales AS
SELECT 
    CASE WHEN EXTRACT(DOW FROM sale_date) IN (0, 6) THEN 'Weekend' ELSE 'Weekday' END AS day_type,
    SUM(total_sale_amount) AS total_revenue,
    AVG(total_sale_amount) AS avg_transaction_value
FROM sales_data
GROUP BY 1;

-- cohort analysis

CREATE OR REPLACE VIEW order_cohort_sequence AS
WITH customer_first_purchase AS (
-- Identify the very first transaction for each customer
    SELECT 
        customer_id, 
        MIN(sale_date) AS first_transaction_date
    FROM sales_data
    GROUP BY customer_id
)
SELECT 
    s.sale_date,
    s.customer_id,
    s.total_sale_amount,
-- The month of the current transaction
    DATE_TRUNC('month', s.sale_date)::DATE AS transaction_month,
-- The month of the customer's first-ever transaction
    DATE_TRUNC('month', f.first_transaction_date)::DATE AS first_transaction_month,
-- The "Cohort Month" as a number (0 = their first month, 1 = one month later, etc.)
    (EXTRACT(YEAR FROM s.sale_date) - EXTRACT(YEAR FROM f.first_transaction_date)) * 12 +
    (EXTRACT(MONTH FROM s.sale_date) - EXTRACT(MONTH FROM f.first_transaction_date)) AS cohort_month
FROM sales_data s
JOIN customer_first_purchase f ON s.customer_id = f.customer_id;

-- revenue by month

CREATE OR REPLACE VIEW revenue_by_month AS
SELECT 
    DATE_TRUNC('month', sale_date)::DATE AS transaction_month,
    SUM(total_sale_amount) AS monthly_revenue,
    COUNT(sale_id) AS total_transactions
FROM 
    sales_data
GROUP BY 1
ORDER BY transaction_month DESC;

-- store wise transaction, revenue & avg transaction value

CREATE OR REPLACE VIEW store_performance_summary AS
SELECT 
    store_id,
    COUNT(sale_id) AS number_of_transactions,
    SUM(total_sale_amount) AS total_revenue,
    ROUND(AVG(total_sale_amount), 2) AS avg_transaction_value
FROM 
    sales_data
GROUP BY 
    store_id
ORDER BY 
    total_revenue DESC;

-- Store with highest avg transaction value

CREATE OR REPLACE VIEW highest_avg_trans_value AS
SELECT 
    store_id,
    COUNT(sale_id) AS number_of_transactions,
    SUM(total_sale_amount) AS total_revenue,
    ROUND(AVG(total_sale_amount), 2) AS avg_transaction_value
FROM 
    sales_data
GROUP BY 
    store_id
ORDER BY 
    avg_transaction_value DESC
	LIMIT 1;

-- transaction wise payment method

CREATE OR REPLACE VIEW payment_method_counts AS
SELECT 
    payment_method,
    COUNT(sale_id) AS transaction_count,
-- Bonus: Show what percentage of customers prefer this method
    ROUND(100.0 * COUNT(sale_id) / SUM(COUNT(sale_id)) OVER (), 2) AS usage_percentage
FROM 
    sales_data
GROUP BY 
    payment_method
ORDER BY 
    transaction_count DESC;

-- Top 10 products with the highest total sales by quantity.

CREATE OR REPLACE VIEW top_10_products_by_quantity AS
SELECT 
    product_id, 
    product_category,
    SUM(quantity_sold) AS total_quantity_sold,
    SUM(total_sale_amount) AS total_revenue
FROM 
    sales_data
GROUP BY 
    product_id, 
    product_category
ORDER BY 
    total_quantity_sold DESC
LIMIT 10;

SELECT * FROM top_10_products_by_quantity;

-- best selling category based on selling

CREATE OR REPLACE VIEW best_selling_category AS
SELECT 
    product_category,
    SUM(total_sale_amount) AS total_revenue,
    SUM(quantity_sold) AS units_sold,
    -- Provides context: are they high-priced items or high-volume items?
    ROUND(SUM(total_sale_amount) / SUM(quantity_sold), 2) AS avg_unit_price
FROM 
    sales_data
GROUP BY 
    product_category
ORDER BY 
    total_revenue DESC
	LIMIT 1;

-- worst selling category based on return rate

CREATE OR REPLACE VIEW worst_performing_returns AS
SELECT 
    product_category,
    COUNT(sale_id) AS total_transactions,
    SUM(CASE WHEN returned = 'Yes' THEN 1 ELSE 0 END) AS return_count,
    ROUND(
        100.0 * SUM(CASE WHEN returned = 'Yes' THEN 1 ELSE 0 END) / COUNT(sale_id), 
        2
    ) AS return_rate_pct
FROM 
    sales_data
GROUP BY 
    product_category
ORDER BY 
    return_rate_pct DESC
	LIMIT 1;

-- Cohort Retention Rate

CREATE OR REPLACE VIEW cohort_retention_rates AS
WITH cohort_counts AS (
    -- Get the starting size of each cohort (Month 0)
    SELECT 
        first_transaction_month,
        COUNT(DISTINCT customer_id) AS original_cohort_size
    FROM order_cohort_sequence
    WHERE cohort_month = 0
    GROUP BY first_transaction_month
),
retention_counts AS (
    -- Get the number of those same customers who returned in Month 1
    SELECT 
        first_transaction_month,
        COUNT(DISTINCT customer_id) AS retained_customers
    FROM order_cohort_sequence
    WHERE cohort_month = 1
    GROUP BY first_transaction_month
)
SELECT 
    c.first_transaction_month,
    c.original_cohort_size,
    COALESCE(r.retained_customers, 0) AS month_1_retained,
    ROUND(100.0 * COALESCE(r.retained_customers, 0) / c.original_cohort_size, 2) AS retention_rate_pct
FROM cohort_counts c
LEFT JOIN retention_counts r ON c.first_transaction_month = r.first_transaction_month;

-- Quarter shown Highest sales

CREATE OR REPLACE VIEW quarterly_performance_labels AS
SELECT 
    CASE 
        WHEN EXTRACT(QUARTER FROM sale_date) = 1 THEN 'First'
        WHEN EXTRACT(QUARTER FROM sale_date) = 2 THEN 'Second'
        WHEN EXTRACT(QUARTER FROM sale_date) = 3 THEN 'Third'
        WHEN EXTRACT(QUARTER FROM sale_date) = 4 THEN 'Fourth'
    END AS quarter_label,
    EXTRACT(YEAR FROM sale_date) AS sale_year,
    SUM(total_sale_amount) AS total_revenue,
    COUNT(sale_id) AS transaction_count
FROM 
    sales_data
GROUP BY 
    sale_year, 
    EXTRACT(QUARTER FROM sale_date)
ORDER BY 
    total_revenue DESC
	Limit 1;