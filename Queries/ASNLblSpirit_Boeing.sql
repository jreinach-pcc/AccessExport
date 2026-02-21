-- Query Name: ASNLblSpirit_Boeing
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT TblPrintASNLabels.BoxSequence, TblPrintASNLabels.Customer, TblPrintASNLabels.ShipTo, TblPrintASNLabels.ShipToAddr1, TblPrintASNLabels.ShipToAddr2, TblPrintASNLabels.ShipToAddr3, TblPrintASNLabels.ShipToCity, TblPrintASNLabels.ShipToState, TblPrintASNLabels.ShipToPostalCode, TblPrintASNLabels.LicensePlateTxt AS Expr1, TblPrintASNLabels.LicensePlateBar AS Expr2, TblPrintASNLabels.PO, TblPrintASNLabels.DispatchNote, TblPrintASNLabels.StockCode, TblPrintASNLabels.QtyTotal, TblPrintASNLabels.NoBoxTotal, TblPrintASNLabels.TotalWt, TblPrintASNLabels.POLine, TblPrintASNLabels.DateShip, TblPrintASNLabels.DateArrive, TblPrintASNLabels.SourceAccepted, "00" & Right(Trim([PO]),7) AS POTxt, Left$(Trim([PO]),3) AS CompanyCodeTxt, TblPrintASNLabels.Lot
FROM TblPrintASNLabels;

