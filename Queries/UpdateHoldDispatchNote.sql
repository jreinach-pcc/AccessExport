-- Query Name: UpdateHoldDispatchNote
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblShippingRecords SET TblShippingRecords.Hold = Forms!Frm7PackingSlip!ChkHold, TblShippingRecords.DateHeld = Forms!Frm7PackingSlip!txtDateHeld
WHERE (((TblShippingRecords.DispatchNote)=[Forms]![Frm7PackingSlip]![txtDispatchNote]));

