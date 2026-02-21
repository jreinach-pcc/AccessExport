-- Query Name: QryInputCount
-- Type: Select
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

SELECT TblShippingLogs.DateLog, Count(*) AS InputCount
FROM TblShippingLogs INNER JOIN dbo_MdnMaster ON TblShippingLogs.DispatchNote = dbo_MdnMaster.DispatchNote
WHERE (((TblShippingLogs.dateLog)=Date()) AND ((dbo_MdnMaster.DispatchNoteStatus)="7"))
GROUP BY TblShippingLogs.DateLog;

