-- Query Name: QryLblRestock
-- Type: Select
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

SELECT dbo_MDNMASTER.DispatchNote, dbo_MDNDETAIL.MStockCode, dbo_MDNDETAILLOT.Lot, dbo_MDNMASTER.CustomerName, dbo_SORDETAIL.MCustRequestDat
FROM ((dbo_MDNMASTER INNER JOIN dbo_MDNDETAIL ON dbo_MDNMASTER.DispatchNote = dbo_MDNDETAIL.DispatchNote) INNER JOIN dbo_MDNDETAILLOT ON dbo_MDNMASTER.DispatchNote = dbo_MDNDETAILLOT.DispatchNote) INNER JOIN dbo_SORDETAIL ON (dbo_MDNDETAILLOT.SalesOrderLine = dbo_SORDETAIL.SalesOrderLine) AND (dbo_MDNDETAILLOT.SalesOrder = dbo_SORDETAIL.SalesOrder)
WHERE (((dbo_MDNMASTER.DispatchNote)="000000000018599") AND ((dbo_MDNDETAIL.DispatchStatus)="9") AND ((dbo_MDNDETAIL.LineType)="1") AND ((dbo_MDNDETAILLOT.DispatchNoteLine)=1));

