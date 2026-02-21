-- Query Name: QryTblPackingSlipUpdateNdtTSO_Cert
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblPackingSlip INNER JOIN qryNdtTSO_Cert ON TblPackingSlip.DispatchNote = qryNdtTSO_Cert.DispatchNote SET TblPackingSlip.NDTSpec = [qryNdtTSO_Cert].[NDTSpec], TblPackingSlip.NDTestSampleSize = [qryNdtTSO_Cert].[NDTestSampleSize], TblPackingSlip.TSO = [qryNdtTSO_Cert].[txtTSO];

