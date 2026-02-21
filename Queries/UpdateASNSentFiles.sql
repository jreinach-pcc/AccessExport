-- Query Name: UpdateASNSentFiles
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblShippingRecords SET TblShippingRecords.ASNSent = True, TblShippingRecords.ASNSentDateTime = Format(Date(),"yyyymmdd") & Format(Time(),"hhnnss")
WHERE (((TblShippingRecords.DateShip)>=[Forms]![Frm10ASNFiles]![txtDateFrom] And (TblShippingRecords.DateShip)<=[Forms]![Frm10ASNFiles]![txtDateTo]));

