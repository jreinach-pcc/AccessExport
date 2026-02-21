-- Query Name: UpdateASNLabelReprint
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblShippingRecords SET TblShippingRecords.LblDateReprinted = Date(), TblShippingRecords.LblReprinted = True, TblShippingRecords.ASN = True
WHERE (((TblShippingRecords.DispatchNote)=[Forms]![Frm7PackingSlip]![txtDispatchNote]));

