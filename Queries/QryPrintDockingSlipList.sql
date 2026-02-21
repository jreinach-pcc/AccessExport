-- Query Name: QryPrintDockingSlipList
-- Type: Select
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

SELECT TblDockingRecords.Job, TblDockingRecords.Customer, TblDockingRecords.CustomerName, TblDockingRecords.GrossQty, TblDockingRecords.StockCode, TblDockingRecords.Lot, TblDockingRecords.PansReceived, TblDockingRecords.ReceivedBy, TblDockingRecords.Warehouse, TblDockingRecords.NoBox1, TblDockingRecords.NoBox2, TblDockingRecords.NoBox3, TblDockingRecords.NoBox4, TblDockingRecords.NoBox5, TblDockingRecords.Bin1, TblDockingRecords.Bin2, TblDockingRecords.Bin3, TblDockingRecords.Bin4, TblDockingRecords.Bin5, TblDockingRecords.Qty1, TblDockingRecords.Qty2, TblDockingRecords.Qty3, TblDockingRecords.Qty4, TblDockingRecords.Qty5, TblDockingRecords.NoBoxTotal, TblDockingRecords.QtyTotal, TblDockingRecords.StockedBy, "*" & Trim([StockCode]) & "*" AS StockCodeBar, "*" & Trim([Lot]) & "*" AS LotBar, "*" & Trim([Bin1]) & "*" AS Bin1Bar, "*" & Trim([Bin2]) & "*" AS Bin2Bar, "*" & Trim([Bin3]) & "*" AS Bin3Bar, "*" & Trim([Bin4]) & "*" AS Bin4Bar, "*" & Trim([Bin5]) & "*" AS Bin5Bar, TblDockingRecords.DateTimeReceived
FROM TblDockingRecords
WHERE (((DateValue([DateTimeReceived]))>=[Forms]![Frm1DockingSlip]![txtFromDate] And (DateValue([DateTimeReceived]))<=[Forms]![Frm1DockingSlip]![txtToDate]));

