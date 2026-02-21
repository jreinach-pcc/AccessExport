-- Query Name: qry_DispatchLoginWarehouseRpt
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblDispatchLoginWarehouse.DispatchNote, TblDispatchLoginWarehouse.DateTimeLogIn, TblDispatchLoginWarehouse.EmpID
FROM TblDispatchLoginWarehouse
WHERE (((DateValue([DateTimeLogIn]))>=[Forms]![Frm2DispatchLoginWarehouse].[txtFromDate] And (DateValue([DateTimeLogIn]))<=[Forms]![Frm2DispatchLoginWarehouse].[txtToDate]));

