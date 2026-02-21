-- Query Name: InvestigateQry
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT Comp_1_MdnMaster.DispatchNote, Comp_1_MdnMaster.CustomerPoNumber, Comp_1_MdnMaster.Customer, Comp_1_MdnMaster.DispatchName, Comp_1_MdnMaster.DispatchAddress1, Comp_1_MdnMaster.DispatchAddress2 AS Expr1, Comp_1_MdnMaster.DispatchAddress3 AS Expr2, Comp_1_MdnMaster.DispatchPostalCode, Comp_1_ArCustomer.Name, Comp_1_ArCustomer.SoldToAddr1, Comp_1_ArCustomer.SoldToAddr2, Comp_1_ArCustomer.SoldToAddr3, Comp_1_ArCustomer.SoldPostalCode, Mid$(Comp_1_MdnMaster.ShippingInstrs,3,20) AS ShipVia, Comp_1_MdnMaster.SalesOrder, Comp_1_TblArTerms.Description, Comp_1_MdnDetail.LineType, Comp_1_MdnDetail.MOrderQty, Comp_1_MdnDetail.MStockQtyToShp, Comp_1_MdnDetail.MStockCode, Comp_1_MdnDetail.MStockDes, Comp_1_MdnDetail.MOrderUom, Comp_1_MdnDetailLot.Lot, Comp_1_InvMaster.UserField1
FROM Comp_1_InvMaster INNER JOIN (Comp_1_MdnDetailLot RIGHT JOIN (((Comp_1_MdnMaster INNER JOIN Comp_1_ArCustomer ON Comp_1_MdnMaster.Customer = Comp_1_ArCustomer.Customer) INNER JOIN Comp_1_TblArTerms ON Comp_1_ArCustomer.TermsCode = Comp_1_TblArTerms.TermsCode) INNER JOIN Comp_1_MdnDetail ON Comp_1_MdnMaster.DispatchNote = Comp_1_MdnDetail.DispatchNote) ON Comp_1_MdnDetailLot.DispatchNote = Comp_1_MdnMaster.DispatchNote) ON Comp_1_InvMaster.StockCode = Comp_1_MdnDetail.MStockCode
WHERE (((Comp_1_MdnMaster.DispatchNote)=Forms!InputForm!txtDispatchNote) And ((Comp_1_MdnDetail.LineType)="1"));

