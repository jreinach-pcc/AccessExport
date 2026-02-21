-- Query Name: QryTblPackingSlipUpdateScheduleB
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblPackingSlip INNER JOIN QryScheduleB ON (TblPackingSlip.SalesOrderLine = QryScheduleB.SOLine) AND (TblPackingSlip.SalesOrder = QryScheduleB.SO) SET TblPackingSlip.ScheduleB = [QryScheduleB].[ScheduleB];

