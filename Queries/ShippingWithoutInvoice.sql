-- Query Name: ShippingWithoutInvoice
-- Type: Select
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

SELECT TblShippingRecords.DispatchNote, TblShippingRecords.Customer, TblShippingRecords.SoldTo, TblShippingRecords.DateCurrent, TblShippingRecords.DateShip, TblShippingRecords.ShippedBy, TblShippingRecords.Carrier, dbo_MDNMASTER.Invoice, TblShippingRecords.Hold, TblShippingRecords.LblDatePrinted, TblShippingRecords.LblDateReprinted, dbo_MDNMASTER.DispatchNoteStatus, TblShippingRecords.Cancel
FROM TblShippingRecords INNER JOIN dbo_MDNMASTER ON TblShippingRecords.DispatchNote = dbo_MDNMASTER.DispatchNote
WHERE (((TblShippingRecords.Hold)<>True) AND ((dbo_MDNMASTER.DispatchNoteStatus)="5" Or (dbo_MDNMASTER.DispatchNoteStatus)="7")) OR (((dbo_MDNMASTER.DispatchNoteStatus)="5" Or (dbo_MDNMASTER.DispatchNoteStatus)="7") AND ((TblShippingRecords.Cancel)<>True));

