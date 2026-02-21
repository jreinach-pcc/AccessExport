-- Query Name: MdnDetailQry
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT dbo_MDNDETAIL.*
FROM dbo_MDNDETAIL
WHERE (((dbo_MDNDETAIL.LineType)<>"5") AND ((dbo_MDNDETAIL.DispatchNoteLine)>=2));

