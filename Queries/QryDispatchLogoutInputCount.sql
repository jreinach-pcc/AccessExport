-- Query Name: QryDispatchLogoutInputCount
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT DateValue([DateTimeLogout]) AS datelog, Count(*) AS InputCount
FROM TblDispatchLogoutWarehouse
GROUP BY DateValue([DateTimeLogout])
HAVING (((DateValue([DateTimeLogout]))=Date()));

