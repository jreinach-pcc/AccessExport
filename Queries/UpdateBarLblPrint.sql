-- Query Name: UpdateBarLblPrint
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblShippingRecords SET TblShippingRecords.ASN = False, TblShippingRecords.LblDatePrinted = Date(), TblShippingRecords.LblPrinted = True
WHERE (((TblShippingRecords.DispatchNote)=[Forms]![Frm7PackingSlip]![txtDispatchNote]));

