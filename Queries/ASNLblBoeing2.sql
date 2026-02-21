-- Query Name: ASNLblBoeing2
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblPrintASNLabels.BoxSequence, TblPrintASNLabels.Customer, TblPrintASNLabels.ShipTo, TblPrintASNLabels.ShipToAddr1, TblPrintASNLabels.ShipToAddr2, TblPrintASNLabels.ShipToAddr3, TblPrintASNLabels.ShipToCity, TblPrintASNLabels.ShipToState, TblPrintASNLabels.ShipToPostalCode, TblPrintASNLabels.PO, TblPrintASNLabels.DispatchNote, TblPrintASNLabels.StockCode, TblPrintASNLabels.QtyTotal, TblPrintASNLabels.NoBoxTotal, TblPrintASNLabels.TotalWt, TblPrintASNLabels.POLine, TblPrintASNLabels.DateShip, TblPrintASNLabels.DateArrive, TblPrintASNLabels.SourceAccepted, TblPrintASNLabels.Lot, TblPrintASNLabels.TSO, TblPrintASNLabels.NDTestSpec, TblPrintASNLabels.NDTestSampleSize, IIf([NDTestSpec]="E1417","P",IIf([NDTestSpec]="E1444","M","")) AS P, IIf([P]="P",IIf([NDTestSampleSize]<>"100%","O",""),IIf([P]="M",IIf([NDTestSampleSize]="100%","O",""),"")) AS O, Left([LicensePlate],21) AS LicensePlateTxt, '*3J~' & Left(Trim([LicensePlate]),21) & '*' AS LicensePlateBar, '3J ' & Left(Trim([LicensePlate]),21) AS LicensePlateBar2D
FROM TblPrintASNLabels INNER JOIN [TblAIC-Boeing] ON TblPrintASNLabels.Customer = [TblAIC-Boeing].Customer
WHERE ((([TblAIC-Boeing].ASNLblName)="ASNLbl_Boeing2"));

