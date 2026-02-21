-- Query Name: Copy Of QryFedexPORpt1
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblShippingRecords.DispatchNote, TblShippingRecords.Customer, TblShippingLogs.DateLog, [PO] & " - " & [NoBoxTotal] AS POInfo, TblOption.Option, TblOption.Name, IIf(TblOption.Customer In ('0021006','0021009'),"",[ShipVia]) AS SV, TblShippingRecords.TotalWt
FROM (TblShippingRecords INNER JOIN TblShippingLogs ON TblShippingRecords.DispatchNote = TblShippingLogs.DispatchNote) INNER JOIN TblOption ON TblShippingRecords.Customer = TblOption.Customer
WHERE (((TblShippingLogs.DateLog) Between Forms!Frm12Fedex!txtFromDate And Forms!Frm12Fedex!txtToDate) And ((TblOption.Option)=Forms!Frm12Fedex!FormSelection) And ((IIf(TblOption.Customer In ('0021006','0021009'),"",[ShipVia]))<>"FEDX#1" And (IIf(TblOption.Customer In ('0021006','0021009'),"",[ShipVia]))<>"FEDX E*" And (IIf(TblOption.Customer In ('0021006','0021009'),"",[ShipVia]))<>"FEDX #1" And (IIf(TblOption.Customer In ('0021006','0021009'),"",[ShipVia]))<>"FEDXE*"))
ORDER BY TblShippingRecords.DispatchNote;

