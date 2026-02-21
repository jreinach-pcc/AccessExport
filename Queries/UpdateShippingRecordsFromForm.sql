-- Query Name: UpdateShippingRecordsFromForm
-- Type: Union
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

UPDATE ShippingRecords SET ShippingRecords.ProposedShipDate = Forms!Frm7PackingSlip!txtDateShip, ShippingRecords.ShipBy = Forms!Frm7PackingSlip!cboShippedBy, ShippingRecords.Carrier = Forms!Frm7PackingSlip!cboCarrier, ShippingRecords.ShipDate = Date(), ShippingRecords.Void = False
WHERE (((ShippingRecords.DispatchNote)=[Forms]![Frm7PackingSlip]![txtDispatchNote]));

