-- Query Name: UpdateCancelJob
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblJobs SET TblJobs.Cancel = Forms!Frm1StockFG!chkCancel, TblJobs.DateCancelled = Forms!Frm1StockFG!txtDateCancelled
WHERE (((TblJobs.Job)=[Forms]![Frm1StockFG]![txtJob]));

