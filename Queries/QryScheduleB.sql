-- Query Name: QryScheduleB
-- Type: Select
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

SELECT Left([KeyField],6) AS SO, Val(Right(Trim([KeyField]),4)) AS SOLine, Trim([AlphaValue]) AS ScheduleB
FROM dbo_AdmFormData
WHERE (((dbo_AdmFormData.AlphaValue)>"0") AND ((dbo_AdmFormData.FieldName)="SCH001"));

