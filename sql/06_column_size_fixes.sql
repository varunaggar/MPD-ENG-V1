-- =============================================================
-- 06_column_size_fixes.sql
-- Widens columns identified during development as too narrow
-- for production data (Graph delta tokens, long error messages).
--
-- Safe to run multiple times — ALTER COLUMN is idempotent.
-- =============================================================

PRINT 'Running 06_column_size_fixes.sql in database: ' + DB_NAME();

-- Graph delta token URLs can exceed 1000 characters
ALTER TABLE dbo.DeltaTokens
    ALTER COLUMN TokenValue NVARCHAR(MAX);
PRINT '  Widened: DeltaTokens.TokenValue -> NVARCHAR(MAX)';

-- SyncLog stores the delta token value at end of run
ALTER TABLE dbo.SyncLog
    ALTER COLUMN TokenAdvancedTo NVARCHAR(MAX);
PRINT '  Widened: SyncLog.TokenAdvancedTo -> NVARCHAR(MAX)';

-- Error messages can include full stack traces
ALTER TABLE dbo.SyncLog
    ALTER COLUMN ErrorMessage NVARCHAR(MAX);
PRINT '  Widened: SyncLog.ErrorMessage -> NVARCHAR(MAX)';
GO
