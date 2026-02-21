-- Query Name: QryDispatchLoginInputCount
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT DateValue([DateTimeLogIn]) AS datelog, Count(*) AS InputCount
FROM TblDispatchLoginWarehouse
GROUP BY DateValue([DateTimeLogIn])
HAVING (((DateValue([DateTimeLogIn]))=Date()));

