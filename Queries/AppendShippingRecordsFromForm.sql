-- Query Name: AppendShippingRecordsFromForm
-- Type: Unknown (64)
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

INSERT INTO ShippingRecords ( DispatchNote, ShipDate, ProposedShipDate, Carrier, ShipBy )
SELECT Forms!Frm7PackingSlip!txtDispatchNote AS DispatchNote, Date() AS DateT, Forms!Frm7PackingSlip!txtDateShip AS DateShip, Forms!Frm7PackingSlip!cboCarrier AS Carrier, Forms!Frm7PackingSlip!cboShippedBy AS ShippedBy;

