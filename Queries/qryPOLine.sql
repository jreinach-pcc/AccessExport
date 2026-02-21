-- Query Name: qryPOLine
-- Type: Select
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

SELECT dbo_MdnDetail.DispatchNote, Val(Mid([NComment],InStr([NComment],'#')+1,3)) AS POLine, dbo_MdnDetail.NComment
FROM dbo_MdnDetail
WHERE (((dbo_MdnDetail.NComment) Like '*ITEM*'));

