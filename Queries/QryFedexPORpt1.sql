-- Query Name: QryFedexPORpt1
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblShippingRecords.DispatchNote, TblShippingRecords.Customer, TblShippingLogs.DateLog, [PO] & " - " & [NoBoxTotal] AS POInfo, TblOption.Option, TblOption.Name, TblShippingRecords.ShipToCity, IIf(TblOption.Customer In ('0021006','0021009'),"",[ShipVia]) AS SV, TblShippingRecords.TotalWt, TblShippingRecords.NoBoxTotal
FROM (TblShippingRecords INNER JOIN TblShippingLogs ON TblShippingRecords.DispatchNote = TblShippingLogs.DispatchNote) INNER JOIN TblOption ON TblShippingRecords.Customer = TblOption.Customer
WHERE (((TblShippingLogs.DateLog) Between Forms!Frm12Fedex!txtFromDate And Forms!Frm12Fedex!txtToDate) And ((TblOption.Option)=Forms!Frm12Fedex!FormSelection) And ((IIf(TblShippingRecords.Customer In ('0021006','0021009'),"",[ShipVia])) Not In ("FEDX#1","FEDX E*","FEDX #1","FEDXE*")));

