-- Query Name: LblPrinting2
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT dbo_WIPMASTER.Job, Left$([JobDescription],9) AS Lot, dbo_WIPMASTER.StockCode, dbo_WIPMASTER.Warehouse
FROM dbo_WIPMASTER
WHERE (((dbo_WIPMASTER.Job)=[Forms]![Frm1StockFG]![txtJob]));

