-- CREATE USER [ori19] FROM EXTERNAL PROVIDER;
-- ALTER ROLE db_datareader ADD MEMBER [ori19];
-- ALTER ROLE db_datawriter ADD MEMBER [ori19];
-- GRANT EXECUTE TO [ori19];
delete from synclog
delete from deltatokens
delete from users

select name from sys.tables

select * from synclog
order by startedat desc;

select * from deltatokens

select userprincipalname from Users

-- 1. Fix the primary storage for delta tokens
--ALTER TABLE dbo.DeltaTokens
-- ALTER COLUMN TokenValue NVARCHAR(MAX);

-- 2. Fix the logging table to prevent errors when recording the last token
-- ALTER TABLE dbo.SyncLog
-- ALTER COLUMN TokenAdvancedTo NVARCHAR(MAX);

-- 3. Optional: Ensure ErrorMessage can handle long stack traces
-- ALTER TABLE dbo.SyncLog
-- ALTER COLUMN ErrorMessage NVARCHAR(MAX);
