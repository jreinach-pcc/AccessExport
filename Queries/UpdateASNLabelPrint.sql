-- Query Name: UpdateASNLabelPrint
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblShippingRecords SET TblShippingRecords.LblDatePrinted = Date(), TblShippingRecords.LblPrinted = True, TblShippingRecords.ASN = True
WHERE (((TblShippingRecords.DispatchNote)=[Forms]![Frm7PackingSlip]![txtDispatchNote]));

