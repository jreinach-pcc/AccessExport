-- Query Name: QryTblPackingSlipUpdatePOLine
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblPackingSlip INNER JOIN qryPOLine ON TblPackingSlip.DispatchNote = qryPOLine.DispatchNote SET TblPackingSlip.POLine = [qryPOLine].[POLine];

