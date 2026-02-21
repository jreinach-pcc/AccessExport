-- Query Name: QryFedexPORpt2
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblShippingRecords.DispatchNote, TblShippingRecords.Customer, TblShipAddr.ShipAddr1, TblShippingRecords.ShipToAddr1, IIf([ShipAddr1]="All",1,IIf(InStr([ShipToAddr1],[Forms]![Frm12Fedex]![cboShipToAddr1])<>0,1,0)) AS CheckShipAddr1, TblShippingLogs.DateLog, [PO] & " - " & [NoBoxTotal] AS POInfo, TblOption.Option, TblOption.Name, IIf(TblOption.Customer In ('0021006','0021009'),"",[ShipVia]) AS SV, TblShippingRecords.TotalWt, TblShippingRecords.NoBoxTotal
FROM ((TblShippingRecords INNER JOIN TblShippingLogs ON TblShippingRecords.DispatchNote = TblShippingLogs.DispatchNote) INNER JOIN TblOption ON TblShippingRecords.Customer = TblOption.Customer) INNER JOIN TblShipAddr ON TblOption.Customer = TblShipAddr.Customer
WHERE (((TblShippingRecords.Customer)=Forms!Frm12Fedex!cboCustomer) And ((TblShipAddr.ShipAddr1)=Forms!Frm12Fedex!cboShipToAddr1) And ((TblShippingLogs.DateLog) Between Forms!Frm12Fedex!txtFromDate And Forms!Frm12Fedex!txtToDate))
ORDER BY TblShippingRecords.DispatchNote;

