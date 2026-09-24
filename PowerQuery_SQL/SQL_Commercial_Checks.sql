-- Commercial data checks. Logic mirrors the Excel controls; not run against a live database.
-- Tables: procurement_spend, supplier_master, finance, fleet, agency, workforce. DATE_TRUNC is PostgreSQL/Snowflake syntax; in SQL Server use DATEFROMPARTS(YEAR(d),MONTH(d),1).

-- 1. Spend by standardised category (case-insensitive), excluding true duplicates
WITH ranked AS (
    SELECT p.*,
           ROW_NUMBER() OVER (PARTITION BY p.InvoiceID, UPPER(TRIM(p.Supplier)), p.GrossSpend
                              ORDER BY p.InvoiceDate) AS dup_rank
    FROM procurement_spend p
)
SELECT UPPER(TRIM(Category))   AS category,
       SUM(GrossSpend)          AS spend,
       COUNT(*)                 AS invoices
FROM ranked
WHERE dup_rank = 1 AND GrossSpend IS NOT NULL
GROUP BY UPPER(TRIM(Category))
ORDER BY spend DESC;

-- 2. Data-quality exception list
SELECT InvoiceID, Supplier, SiteID, GrossSpend,
       CASE WHEN Supplier IS NULL OR TRIM(Supplier) = '' THEN 'Missing supplier'
            WHEN SiteID   IS NULL OR TRIM(SiteID)   = '' THEN 'Missing site'
            WHEN GrossSpend IS NULL                      THEN 'Missing spend'
            WHEN COUNT(*) OVER (PARTITION BY InvoiceID) > 1 THEN 'Duplicate ID - review'
            ELSE 'OK' END AS dq_flag
FROM procurement_spend;

-- 3. Invoice category disagrees with supplier master
SELECT p.InvoiceID, p.Supplier, p.Category AS invoice_category, s.Category AS master_category
FROM procurement_spend p
JOIN supplier_master s ON UPPER(TRIM(p.Supplier)) = UPPER(s.Supplier)
WHERE UPPER(TRIM(p.Category)) <> UPPER(s.Category);

-- 4. Supplier concentration and contract compliance
SELECT Supplier,
       SUM(GrossSpend) AS spend,
       SUM(GrossSpend) * 1.0 / SUM(SUM(GrossSpend)) OVER () AS share,
       SUM(CASE WHEN Contracted = 'N' THEN GrossSpend ELSE 0 END) AS off_contract_spend
FROM procurement_spend
GROUP BY Supplier
ORDER BY spend DESC;

-- 5. Reconciliation: procurement extract vs finance ledger by month
SELECT f.month,
       f.ledger_procurement,
       p.extract_procurement,
       p.extract_procurement - f.ledger_procurement AS difference
FROM (SELECT Month AS month, SUM(Procurement) AS ledger_procurement FROM finance GROUP BY Month) f
LEFT JOIN (SELECT DATE_TRUNC('month', InvoiceDate) AS month, SUM(GrossSpend) AS extract_procurement
           FROM procurement_spend GROUP BY DATE_TRUNC('month', InvoiceDate)) p
       ON p.month = f.month
ORDER BY f.month;

-- 6. Budget vs actual Opex by site with RAG
SELECT SiteID, SiteName,
       SUM(BudgetOpex) AS budget, SUM(ActualOpex) AS actual,
       (SUM(BudgetOpex) - SUM(ActualOpex)) / NULLIF(SUM(BudgetOpex), 0) AS variance_pct,
       CASE WHEN (SUM(BudgetOpex) - SUM(ActualOpex)) / NULLIF(SUM(BudgetOpex), 0) < -0.01 THEN 'Red'
            WHEN SUM(ActualOpex) > SUM(BudgetOpex) THEN 'Amber' ELSE 'Green' END AS rag
FROM finance
GROUP BY SiteID, SiteName
ORDER BY variance_pct;

-- 7. Agency share of staff cost by region
SELECT w.Region,
       a.agency_spend,
       a.agency_spend / (a.agency_spend + w.payroll) AS agency_share
FROM (SELECT Region, SUM(TotalPayroll) AS payroll FROM workforce GROUP BY Region) w
JOIN (SELECT Region, SUM(AgencySpend) AS agency_spend FROM agency GROUP BY Region) a
  ON a.Region = w.Region;
