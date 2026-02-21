-- Query Name: FIX the Bins
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblDockingRecords.DateTimeReceived, TblDockingRecords.NoBox1, TblDockingRecords.Bin1, TblDockingRecords.NoBox2, TblDockingRecords.Bin2, TblDockingRecords.NoBox3, TblDockingRecords.Bin3, TblDockingRecords.NoBox4, TblDockingRecords.Bin4, TblDockingRecords.NoBox5, TblDockingRecords.Bin5
FROM TblDockingRecords
WHERE (((TblDockingRecords.Bin3)>"0"));

