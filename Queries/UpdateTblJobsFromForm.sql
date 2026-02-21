-- Query Name: UpdateTblJobsFromForm
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblJobs SET TblJobs.Lot = Forms!Frm1StockFG!txtLot, TblJobs.StockCode = Forms!Frm1StockFG!txtStockCode, TblJobs.Revision = Forms!Frm1StockFG!txRevision, TblJobs.Warehouse = Forms!Frm1StockFG!txtWarehouse, TblJobs.StockedBy = Forms!Frm1StockFG!cboStockedBy, TblJobs.DateStocked = Forms!Frm1StockFG!txtDateStocked, TblJobs.Cancel = Forms!Frm1StockFG!chkCancel, TblJobs.DateCancelled = Forms!Frm1StockFG!txtDateCancelled, TblJobs.NoBoxTotal = Forms!Frm1StockFG!txtNoBoxTotal, TblJobs.QtyTotal = Forms!Frm1StockFG!txQtyTotal, TblJobs.NoBox1 = Forms!Frm1StockFG!txtNoBox1, TblJobs.NoBox2 = Forms!Frm1StockFG!txtNoBox2, TblJobs.NoBox3 = Forms!Frm1StockFG!txtNoBox3, TblJobs.NoBox4 = Forms!Frm1StockFG!txtNoBox4, TblJobs.NoBox5 = Forms!Frm1StockFG!txtNoBox5, TblJobs.NoBox6 = Forms!Frm1StockFG!txtNoBox6, TblJobs.Bin1 = Forms!Frm1StockFG!txtBin1, TblJobs.Bin2 = Forms!Frm1StockFG!txtBin2, TblJobs.Bin3 = Forms!Frm1StockFG!txtBin3, TblJobs.Bin4 = Forms!Frm1StockFG!txtBin4, TblJobs.Bin5 = Forms!Frm1StockFG!txtBin5, TblJobs.Bin6 = Forms!Frm1StockFG!txtBin6, TblJobs.Qty1 = Forms!Frm1StockFG!txtQty1, TblJobs.Qty2 = Forms!Frm1StockFG!txtQty2, TblJobs.Qty3 = Forms!Frm1StockFG!txtQty3, TblJobs.Qty4 = Forms!Frm1StockFG!txtQty4, TblJobs.Qty5 = Forms!Frm1StockFG!txtQty5, TblJobs.Qty6 = Forms!Frm1StockFG!txtQty6
WHERE (((TblJobs.Job)=[Forms]![Frm1StockFG]![txtJob]));

