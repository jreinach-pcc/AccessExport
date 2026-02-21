-- Query Name: QryTblPackingSlipUpdateNdt_LabC
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblPackingSlip INNER JOIN qryNdtTSO_Lab ON TblPackingSlip.Lot = qryNdtTSO_Lab.LotNumber SET TblPackingSlip.NDTSpec = [qryNdtTSO_Lab].[NDTSpec], TblPackingSlip.NDTestSampleSize = [qryNdtTSO_Lab].[NDTestSampleSize]
WHERE (((Trim([TblPackingSlip].[NDTSpec])) Is Null Or (Trim([TblPackingSlip].[NDTSpec]))<"0" Or (Trim([TblPackingSlip].[NDTSpec]))=""));

