-- Query Name: QryTblPackingSlipUpdateECCN
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE DISTINCTROW TblPackingSlip INNER JOIN QryECCN ON (TblPackingSlip.SalesOrderLine = QryECCN.SOLine) AND (TblPackingSlip.SalesOrder = QryECCN.SO) SET TblPackingSlip.ECCN = [QryECCN].[ECCN];

