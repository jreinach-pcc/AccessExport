-- Query Name: Copy Of zQryASNBoeing
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT [TblAIC-Boeing].RecordIdentifier AS ASN00, [TblAIC-Boeing].AICISA AS ASN01, Left$(Trim([PO]),3) AS ASN02, Right$(Trim([PO]),9) AS ASN03, Format([POLine],"000") AS ASN04, Trim([StockCode]) AS ASN05, TblShippingRecords.QtyTotal AS ASN06, "EA" AS ASN07, TblShippingRecords.NoBoxTotal AS ASN08, TblShippingRecords.TotalWt AS ASN11, "LBS" AS ASN12, Format([DateShip],"yyyymmdd") AS ASN13, Format([DateArrive],"yyyymmdd") AS ASN14, Trim([SCAC]) AS ASN15, Right(Trim([DispatchNote]),6) AS ASN16, TblShippingRecords.ShipTo AS ASN19, TblShippingRecords.ShipToAddr1 AS ASN20, TblShippingRecords.ShipToAddr2 AS ASN21, TblShippingRecords.ShipToAddr3 AS ASN22, TblShippingRecords.ShipToCity AS ASN23, TblShippingRecords.ShipToState AS ASN24, "USA" AS ASN26, TblShippingRecords.SourceAccepted AS ASN35, [TblAIC-Boeing].WarehouseCode AS ASN36, Trim([DeliveryMethod]) AS ASN38, TblShippingRecords.ASNSent, TblShippingRecords.ASNSentDateTime
FROM TblShippingRecords INNER JOIN [TblAIC-Boeing] ON TblShippingRecords.Customer = [TblAIC-Boeing].Customer
WHERE (((TblShippingRecords.DateShip) Between [ASN856_DateFrom] And [ASN856_DateTo]));

