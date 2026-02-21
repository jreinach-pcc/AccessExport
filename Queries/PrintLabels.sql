-- Query Name: PrintLabels
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT dbo_MDNMASTER.DispatchNote, dbo_MDNMASTER.Customer, dbo_MDNMASTER.CustomerName, dbo_MDNMASTER.CustomerPoNumber, dbo_MDNMASTER.DispatchCustName, dbo_MDNMASTER.DispatchAddress1, dbo_MDNMASTER.DispatchAddress2, dbo_MDNMASTER.DispatchAddress3, dbo_MDNMASTER.DispatchAddress4, dbo_MDNMASTER.DispatchAddress5, dbo_MDNMASTER.DispatchPostalCode, dbo_MDNDETAIL.LineType, dbo_MDNDETAIL.MStockCode, dbo_MDNDETAILLOT.Lot
FROM (dbo_MDNMASTER INNER JOIN dbo_MDNDETAIL ON dbo_MDNMASTER.DispatchNote = dbo_MDNDETAIL.DispatchNote) INNER JOIN dbo_MDNDETAILLOT ON dbo_MDNMASTER.DispatchNote = dbo_MDNDETAILLOT.DispatchNote
WHERE (((dbo_MDNDETAIL.LineType)="1"));

