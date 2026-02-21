-- Query Name: BillOfLadingBoeing
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblShippingRecords.DateShip, TblShippingRecords.Customer, TblShippingRecords.DispatchNote, TblShippingRecords.BillOfLading, TblShippingRecords.Hold, TblShippingRecords.DateHeld
FROM TblShippingRecords
WHERE (((TblShippingRecords.Customer)=[Enter Customer #]));

