-- Query Name: TblPasswordAppend
-- Type: Unknown (64)
-- Date Exported: 2/21/2026 9:53:56 AM
-- ----------------------------------------------------------------------

INSERT INTO TblPassword ( UserName, KeyCode, SuperUser )
SELECT dbo_TblPassword.UserName, dbo_TblPassword.KeyCode, dbo_TblPassword.SuperUser
FROM dbo_TblPassword;

