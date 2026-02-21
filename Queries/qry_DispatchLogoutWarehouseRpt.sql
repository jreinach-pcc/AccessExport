-- Query Name: qry_DispatchLogoutWarehouseRpt
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblDispatchLogoutWarehouse.DispatchNote, TblDispatchLogoutWarehouse.DateTimeLogout, TblDispatchLogoutWarehouse.EmpID
FROM TblDispatchLogoutWarehouse
WHERE (((DateValue([DateTimeLogOut]))>=[Forms]![Frm3DispatchLogoutWarehouse].[txtFromDate] And (DateValue([DateTimeLogOut]))<=[Forms]![Frm3DispatchLogoutWarehouse].[txtToDate]));

