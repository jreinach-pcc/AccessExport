-- Query Name: QryPrintDockingSlip
-- Type: Select
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

SELECT TblDockingRecords.Job, TblDockingRecords.Customer, TblDockingRecords.CustomerName, TblDockingRecords.GrossQty, TblDockingRecords.StockCode, TblDockingRecords.Lot, TblDockingRecords.PansReceived, TblDockingRecords.DateTimeReceived, TblDockingRecords.ReceivedBy, TblDockingRecords.Warehouse, TblDockingRecords.NoBox1, TblDockingRecords.NoBox2, TblDockingRecords.NoBox3, TblDockingRecords.NoBox4, TblDockingRecords.NoBox5, TblDockingRecords.Bin1, TblDockingRecords.Bin2, TblDockingRecords.Bin3, TblDockingRecords.Bin4, TblDockingRecords.Bin5, TblDockingRecords.Qty1, TblDockingRecords.Qty2, TblDockingRecords.Qty3, TblDockingRecords.Qty4, TblDockingRecords.Qty5, TblDockingRecords.NoBoxTotal, TblDockingRecords.QtyTotal, TblDockingRecords.StockedBy, "*" & Trim([StockCode]) & "*" AS StockCodeBar, "*" & Trim([Lot]) & "*" AS LotBar, "*" & Trim([Bin1]) & "*" AS Bin1Bar, "*" & Trim([Bin2]) & "*" AS Bin2Bar, "*" & Trim([Bin3]) & "*" AS Bin3Bar, "*" & Trim([Bin4]) & "*" AS Bin4Bar, "*" & Trim([Bin5]) & "*" AS Bin5Bar, "*" & Trim(Str([Qty1])) & "*" AS Qty1Bar, "*" & Trim(Str([Qty2])) & "*" AS Qty2Bar, "*" & Trim(Str([Qty3])) & "*" AS Qty3Bar, "*" & Trim(Str([Qty4])) & "*" AS Qty4Bar, "*" & Trim(Str([Qty5])) & "*" AS Qty5Bar
FROM TblDockingRecords
WHERE (((TblDockingRecords.Job)=[Forms]![Frm1DockingSlip]![txtJob]));

