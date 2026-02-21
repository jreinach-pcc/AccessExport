-- Query Name: QryFedexPORptKLX
-- Type: Select
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

SELECT TblShippingRecords.DispatchNote, TblShippingRecords.Customer, TblShippingLogs.DateLog, [PO] & " - " & [NoBoxTotal] AS POInfo, TblOption.Option, TblOption.Name, TblShippingRecords.ShipToAddr2, TblShippingRecords.ShipVia, IIf(TblOption.Customer In ('0021006','0021009'),"",[ShipVia]) AS SV, TblShippingRecords.TotalWt, TblShippingRecords.NoBoxTotal, Trim([Forms]![Frm12Fedex]![cboShipToState]) AS State
FROM (TblShippingRecords INNER JOIN TblShippingLogs ON TblShippingRecords.DispatchNote = TblShippingLogs.DispatchNote) INNER JOIN TblOption ON TblShippingRecords.Customer = TblOption.Customer
WHERE (((TblShippingLogs.DateLog) Between Forms!Frm12Fedex!txtFromDate And Forms!Frm12Fedex!txtToDate) And ((TblOption.Option)=Forms!Frm12Fedex!FormSelection) And ((TblShippingRecords.ShipToAddr2) Like "*" & Trim(Forms!Frm12Fedex!cboShipToState) & "*") And ((TblShippingRecords.ShipVia) Not Like "*" & "OVERNIGHT" & "*") And ((IIf(TblShippingRecords.Customer In ('0021006','0021009'),"",[ShipVia])) Not In ("FEDX#1","FEDX E*","FEDX #1","FEDXE*")));

