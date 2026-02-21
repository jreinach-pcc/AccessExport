-- Query Name: QryAdeptPORpt
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblShippingRecords.DispatchNote, TblShippingRecords.Customer, IIf([TblShippingRecords].[Customer]<>"0040000",1,IIf(InStr([ShipToAddr1],"KOKU")<>0,1,0)) AS Expr1, TblShippingLogs.DateLog, [PO] & " - " & [NoBoxTotal] AS POInfo, TblOption.Option, TblOption.Customer, TblOption.Name, IIf(TblOption.Customer In ('0021006','0021009'),"",[ShipVia]) AS SV, [Wt1]+[Wt2]+[Wt3]+[Wt4] AS TotalWt
FROM (TblShippingRecords INNER JOIN TblShippingLogs ON TblShippingRecords.DispatchNote = TblShippingLogs.DispatchNote) INNER JOIN TblOption ON TblShippingRecords.Customer = TblOption.Customer
WHERE (((TblShippingRecords.Customer)="00010855") And ((IIf(TblShippingRecords.Customer<>"0040000",1,IIf(InStr([ShipToAddr1],"KOKU")<>0,1,0)))=1) And ((TblShippingLogs.DateLog) Between Forms!Frm12Fedex!txtFromDate And Forms!Frm12Fedex!txtToDate) And ((TblOption.Customer)="0010855"))
ORDER BY TblShippingRecords.DispatchNote;

