-- Commercial analytics validation examples.
-- Adapt table/column names to the target SQL platform.

-- 1. Total procurement spend
SELECT
    SUM(GrossSpend) AS total_procurement_spend
FROM procurement_spend;

-- 2. Spend by category
SELECT
    Category,
    SUM(GrossSpend) AS category_spend,
    COUNT(*) AS invoice_count
FROM procurement_spend
GROUP BY Category
ORDER BY category_spend DESC;

-- 3. Regional spend
SELECT
    Region,
    SUM(GrossSpend) AS regional_spend
FROM procurement_spend
GROUP BY Region
ORDER BY regional_spend DESC;

-- 4. Duplicate invoice detection
SELECT
    InvoiceID,
    COUNT(*) AS invoice_count
FROM procurement_spend
GROUP BY InvoiceID
HAVING COUNT(*) > 1;

-- 5. Supplier concentration
SELECT
    Supplier,
    SUM(GrossSpend) AS supplier_spend,
    SUM(GrossSpend) / NULLIF((SELECT SUM(GrossSpend) FROM procurement_spend), 0) AS spend_share
FROM procurement_spend
GROUP BY Supplier
ORDER BY supplier_spend DESC;
