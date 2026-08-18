-- =====================================================================
-- Retail Sales Performance Dashboard - SQL Analysis Queries
-- Table: sales (loaded from superstore_cleaned.csv)
-- =====================================================================


-- 1. MONTHLY REVENUE TREND
-- Shows revenue trend over time to spot seasonal patterns
-- ---------------------------------------------------------------------
SELECT
    "Order Year-Month" AS month,
    ROUND(SUM(Sales), 2) AS total_revenue,
    COUNT(DISTINCT "Order ID") AS num_orders,
    ROUND(SUM(Sales) * 1.0 / COUNT(DISTINCT "Order ID"), 2) AS avg_order_value
FROM sales
GROUP BY "Order Year-Month"
ORDER BY "Order Year-Month";


-- 2. YEAR-OVER-YEAR (YoY) REVENUE GROWTH
-- Compares each year's total revenue against the previous year
-- ---------------------------------------------------------------------
WITH yearly AS (
    SELECT "Order Year" AS year, ROUND(SUM(Sales), 2) AS revenue
    FROM sales
    GROUP BY "Order Year"
)
SELECT
    year,
    revenue,
    LAG(revenue) OVER (ORDER BY year) AS prev_year_revenue,
    ROUND(
        (revenue - LAG(revenue) OVER (ORDER BY year)) * 100.0
        / LAG(revenue) OVER (ORDER BY year), 2
    ) AS yoy_growth_pct
FROM yearly
ORDER BY year;


-- 3. SEASONAL PATTERN - AVG REVENUE BY MONTH (ACROSS ALL YEARS)
-- Reveals which months consistently perform best/worst
-- ---------------------------------------------------------------------
SELECT
    "Order Month" AS month_num,
    "Order Month Name" AS month_name,
    ROUND(SUM(Sales), 2) AS total_revenue,
    ROUND(AVG(Sales), 2) AS avg_transaction_value,
    COUNT(*) AS num_transactions
FROM sales
GROUP BY "Order Month", "Order Month Name"
ORDER BY "Order Month";


-- 4. SALES BY CATEGORY (TOP REVENUE-DRIVING CATEGORIES)
-- Identifies the top revenue drivers - used for stakeholder presentation
-- ---------------------------------------------------------------------
SELECT
    Category,
    ROUND(SUM(Sales), 2) AS total_revenue,
    COUNT(*) AS num_transactions,
    ROUND(SUM(Sales) * 100.0 / (SELECT SUM(Sales) FROM sales), 2) AS pct_of_total_revenue
FROM sales
GROUP BY Category
ORDER BY total_revenue DESC;


-- 5. SALES BY SUB-CATEGORY (DRILL-DOWN)
-- ---------------------------------------------------------------------
SELECT
    Category,
    "Sub-Category",
    ROUND(SUM(Sales), 2) AS total_revenue,
    COUNT(*) AS num_transactions
FROM sales
GROUP BY Category, "Sub-Category"
ORDER BY total_revenue DESC
LIMIT 10;


-- 6. REGION-WISE PERFORMANCE (IDENTIFY UNDERPERFORMING REGIONS)
-- ---------------------------------------------------------------------
SELECT
    Region,
    ROUND(SUM(Sales), 2) AS total_revenue,
    COUNT(DISTINCT "Order ID") AS num_orders,
    ROUND(SUM(Sales) * 1.0 / COUNT(DISTINCT "Order ID"), 2) AS avg_order_value,
    RANK() OVER (ORDER BY SUM(Sales) DESC) AS revenue_rank
FROM sales
GROUP BY Region
ORDER BY total_revenue DESC;


-- 7. STATE-LEVEL PERFORMANCE (BOTTOM 10 - UNDERPERFORMING STATES)
-- ---------------------------------------------------------------------
SELECT
    State,
    Region,
    ROUND(SUM(Sales), 2) AS total_revenue,
    COUNT(*) AS num_transactions
FROM sales
GROUP BY State, Region
ORDER BY total_revenue ASC
LIMIT 10;


-- 8. CUSTOMER SEGMENT PERFORMANCE
-- ---------------------------------------------------------------------
SELECT
    Segment,
    ROUND(SUM(Sales), 2) AS total_revenue,
    COUNT(DISTINCT "Customer ID") AS num_customers,
    ROUND(SUM(Sales) * 1.0 / COUNT(DISTINCT "Customer ID"), 2) AS avg_revenue_per_customer
FROM sales
GROUP BY Segment
ORDER BY total_revenue DESC;


-- 9. QUARTERLY REVENUE BY REGION (FOR DRILL-DOWN VISUAL)
-- ---------------------------------------------------------------------
SELECT
    Region,
    "Order Year" AS year,
    "Order Quarter" AS quarter,
    ROUND(SUM(Sales), 2) AS total_revenue
FROM sales
GROUP BY Region, "Order Year", "Order Quarter"
ORDER BY Region, year, quarter;


-- 10. TOP 3 REVENUE-DRIVING PRODUCT CATEGORIES (FOR STAKEHOLDER SUMMARY)
-- ---------------------------------------------------------------------
SELECT
    Category,
    ROUND(SUM(Sales), 2) AS total_revenue,
    ROUND(SUM(Sales) * 100.0 / (SELECT SUM(Sales) FROM sales), 2) AS pct_of_total
FROM sales
GROUP BY Category
ORDER BY total_revenue DESC
LIMIT 3;
