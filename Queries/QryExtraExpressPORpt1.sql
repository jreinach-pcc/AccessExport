-- Query Name: QryExtraExpressPORpt1
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblShippingRecords.Customer, "BOEING - NEW BREED LOGISTICS" AS Name, TblShippingLogs.DateLog, TblShippingRecords.DispatchNote, TblShippingRecords.BillOfLading, [PO] & " - " & [NoBoxTotal] AS POInfo, TblShippingRecords.TotalWt, TblShippingRecords.QtyTotal, TblShippingRecords.NoBoxTotal
FROM TblShippingRecords INNER JOIN TblShippingLogs ON TblShippingRecords.DispatchNote = TblShippingLogs.DispatchNote
WHERE (((TblShippingRecords.Customer) In ('0021008','0021010')) And ((TblShippingLogs.DateLog) Between Forms!Frm12Fedex!txtFromDate And Forms!Frm12Fedex!txtToDate))
ORDER BY TblShippingRecords.DispatchNote;

