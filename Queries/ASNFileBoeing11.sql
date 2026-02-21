-- Query Name: ASNFileBoeing11
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT 'SA1' AS SA101, Left$(Trim([PO]),3) & Format(Date(),'yyyymmdd') & Format(Time(),'hhnnss') AS SA102, 'UN623351603' & Right(Trim([DispatchNote]),10) & Format([NoBoxTotal],'000') AS SA105, Format(Date(),'yyyymmdd') & Format(Time(),'hhnnss') AS SA106, 'SA1_END' AS SA107, 'SA2' AS SA201, Left$(Trim([PO]),3) & Format(Date(),'yyyymmdd') & Format(Time(),'hhnnss') AS SA202, 'UN623351603' & Right(Trim([DispatchNote]),10) & Format([NoBoxTotal],'000') AS SA204, Trim([BillOfLading]) AS SA205, Right(Trim([DispatchNote]),10) AS SA206, '2' AS SA207, Format([DateShip],'yyyymmdd') & '120000' AS SA208, Format([DateArrive],'yyyymmdd') & '110000' AS SA209, 'SA2_END' AS SA212, 'SA7' AS SA701, Left$(Trim([PO]),3) & Format(Date(),'yyyymmdd') & Format(Time(),'hhnnss') AS SA702, 'UN623351603' & Right(Trim([DispatchNote]),10) & Format([NoBoxTotal],'000') AS SA704, '1' AS SA705, Trim([StockCode]) AS SA706, Trim(QtyTotal) AS SA707, 'EA' AS SA708, Right$(Trim([PO]),9) AS SA709, Format(Trim([POLine]),'0000') AS SA710, 'SA7_END' AS SA716, 'SA8' AS SA801, Left$(Trim([PO]),3) & Format(Date(),'yyyymmdd') & Format(Time(),'hhnnss') AS SA802, 'UN623351603' & Right(Trim([DispatchNote]),10) & Format([NoBoxTotal],'000') AS SA804, '1' AS SA805, Trim(NoBoxTotal) AS SA806, 'SA8_END' AS SA807, TblShippingRecords.DateShip, TblShippingRecords.ASN, TblShippingRecords.ASNSent, TblShippingRecords.ASNSentDateTime, TblShippingRecords.Hold, [TblAIC-Boeing].ASNLblName, [TblAIC-Boeing].ASN2
FROM TblShippingRecords INNER JOIN [TblAIC-Boeing] ON TblShippingRecords.Customer = [TblAIC-Boeing].Customer
WHERE (((TblShippingRecords.DateShip)=#10/26/2017#) AND ((TblShippingRecords.ASN)=True) AND ((TblShippingRecords.ASNSent)=False) AND ((TblShippingRecords.Hold)=False) AND (([TblAIC-Boeing].ASN2)="N"))
ORDER BY TblShippingRecords.DateShip;

