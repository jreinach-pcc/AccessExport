-- Query Name: CheckPO
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT dbo_MDNMASTER.Customer, dbo_MDNMASTER.DispatchNote, dbo_MDNMASTER.*
FROM dbo_MDNMASTER
WHERE (((dbo_MDNMASTER.Customer)="0021700"));

