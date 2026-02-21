-- Query Name: QryFedexPORpt4
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblShippingRecords.DispatchNote, TblShippingRecords.Customer, TblShippingRecords.ShipToAddr1, TblShippingRecords.ShipToAddr2, TblShippingRecords.ShipToAddr3, TblShippingLogs.DateLog, [PO] & " - " & [NoBoxTotal] AS POInfo, TblShippingRecords.TotalWt, dbo_ArCustomer.Name, TblShippingRecords.NoBoxTotal
FROM (TblShippingRecords INNER JOIN dbo_ArCustomer ON TblShippingRecords.Customer = dbo_ArCustomer.Customer) INNER JOIN TblShippingLogs ON TblShippingRecords.DispatchNote = TblShippingLogs.DispatchNote
WHERE (((TblShippingRecords.Customer)=Forms!Frm12Fedex!cboCustomer4) And ((TblShippingLogs.DateLog) Between Forms!Frm12Fedex!txtFromDate And Forms!Frm12Fedex!txtToDate))
ORDER BY TblShippingRecords.DispatchNote;

