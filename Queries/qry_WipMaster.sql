-- Query Name: qry_WipMaster
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT dbo_WipMaster.Job, dbo_WipMaster.Customer, dbo_WipMaster.CustomerName, dbo_WipMaster.GrossQty, dbo_WipMaster.StockCode, Left([JobDescription],InStr(1,[JobDescription],"")) AS Lot, dbo_WipMaster.Warehouse
FROM dbo_WipMaster;

