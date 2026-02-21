-- Query Name: qryNdtTSO_Cert
-- Type: Select
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

SELECT dbo_TblCertificationMaster.DispatchNote, IIf(InStr([NDTestSpec],"E1417")>0,"E1417",IIf(InStr([NDTestSpec],"E1444")>0,"E1444","")) AS NDTSpec, dbo_TblCertificationMaster.NDTestSampleSize, IIf(Trim([TSO])="Y","TSO-C148",IIf(Trim([TSO])="N" Or IsNull([TSO]),"",[TSO])) AS txtTSO
FROM dbo_TblCertificationMaster;

