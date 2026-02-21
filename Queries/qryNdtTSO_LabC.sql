-- Query Name: qryNdtTSO_LabC
-- Type: Select
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

SELECT dbo_LabCertRecord.[LOT NUMBER] AS LotNumber, IIf(InStr([NDT],"E1417")>0,"E1417",IIf(InStr([NDT],"E1444")>0,"E1444","")) AS NDTSpecLC, dbo_LabCertRecord.NDTSamp AS NDTestSampleSizeLC, IIf(Trim([TSO])="Y" Or [TSO]=True,"TSO-C148",IIf(Trim([TSO])="N" Or IsNull([TSO]) Or [TSO]=False,"",[TSO])) AS txtTSOLC
FROM dbo_LabCertRecord;

