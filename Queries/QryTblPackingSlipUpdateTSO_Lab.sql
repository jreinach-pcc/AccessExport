-- Query Name: QryTblPackingSlipUpdateTSO_Lab
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE TblPackingSlip INNER JOIN qryNdtTSO_Lab ON TblPackingSlip.Lot = qryNdtTSO_Lab.LotNumber SET TblPackingSlip.TSO = [qryNdtTSO_Lab].[txtTSO]
WHERE (((Trim([TblPackingSlip].[TSO])) Is Null Or (Trim([TblPackingSlip].[TSO]))<"0" Or (Trim([TblPackingSlip].[TSO]))=""));

