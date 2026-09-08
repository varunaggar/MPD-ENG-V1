-- =============================================================
-- 01_shared_tables.sql
-- Shared operational tables used by every process.
-- Run this script first, before any other SQL scripts.
-- Connect to: db-m365permissions (NOT master)
-- =============================================================

PRINT 'Creating shared tables in database: ' + DB_NAME();
PRINT '';

-- ──────────────────────────────────────────────────────────────
-- DeltaTokens
-- Stores cursor positions for incremental Microsoft Graph queries.
-- One row per named token (e.g. users_delta, groups_delta).
-- ──────────────────────────────────────────────────────────────

IF OBJECT_ID('dbo.DeltaTokens', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DeltaTokens (
        -- The logical name identifying which source this token belongs to
        -- e.g. 'users_delta', 'groups_delta', 'mailboxes_delta'
        TokenName           NVARCHAR(100)   NOT NULL
            CONSTRAINT PK_DeltaTokens PRIMARY KEY,

        -- The opaque cursor value returned by Microsoft Graph
        TokenValue          NVARCHAR(MAX)  NOT NULL,

        -- When this token row was first created
        CreatedAt           DATETIME2       NOT NULL
            CONSTRAINT DF_DeltaTokens_CreatedAt DEFAULT SYSUTCDATETIME(),

        -- When this token was last successfully advanced
        UpdatedAt           DATETIME2       NOT NULL
            CONSTRAINT DF_DeltaTokens_UpdatedAt DEFAULT SYSUTCDATETIME(),

        -- 0 means the token is no longer valid (e.g. HTTP 410 Gone received)
        -- Must run the baseline process to reactivate
        IsActive            BIT             NOT NULL
            CONSTRAINT DF_DeltaTokens_IsActive DEFAULT 1,

        -- Populated when IsActive is set to 0
        DeactivatedAt       DATETIME2       NULL,
        DeactivationReason  NVARCHAR(500)   NULL
    );

    PRINT '  Created: DeltaTokens';
END
ELSE
    PRINT '  Already exists: DeltaTokens';
GO

-- ──────────────────────────────────────────────────────────────
-- SyncLog
-- One row per process execution. Used for:
--   - Monitoring freshness (when was the last successful run?)
--   - Alerting on failures
--   - Debugging (how many records were processed, any errors?)
--   - Trend analysis (how long does each process normally take?)
-- ──────────────────────────────────────────────────────────────

IF OBJECT_ID('dbo.SyncLog', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.SyncLog (
        -- Unique identifier for this execution
        RunId               UNIQUEIDENTIFIER NOT NULL
            CONSTRAINT PK_SyncLog PRIMARY KEY
            CONSTRAINT DF_SyncLog_RunId DEFAULT NEWID(),

        -- Name of the script that created this row
        -- e.g. 'Invoke-UserBaselineLoad', 'Invoke-UserDeltaSync'
        FunctionName        NVARCHAR(100)   NOT NULL,

        -- UTC timestamps
        StartedAt           DATETIME2       NOT NULL,
        CompletedAt         DATETIME2       NULL,

        -- Running | Success | PartialFailure | Failed
        Status              NVARCHAR(20)    NULL,

        -- Row-level counters for the Users sync processes
        -- Extend these columns (or convert to JSON) when adding
        -- mailbox/permission sync processes
        UsersInserted       INT             NOT NULL
            CONSTRAINT DF_SyncLog_UsersInserted DEFAULT 0,
        UsersUpdated        INT             NOT NULL
            CONSTRAINT DF_SyncLog_UsersUpdated DEFAULT 0,
        UsersSoftDeleted    INT             NOT NULL
            CONSTRAINT DF_SyncLog_UsersSoftDeleted DEFAULT 0,
        UsersProcessed      INT             NOT NULL
            CONSTRAINT DF_SyncLog_UsersProcessed DEFAULT 0,

        -- Error summary
        ErrorCount          INT             NOT NULL
            CONSTRAINT DF_SyncLog_ErrorCount DEFAULT 0,
        ErrorMessage        NVARCHAR(MAX)  NULL,

        -- The delta token value that was saved at the end of this run
        -- NULL for baseline loads that capture the initial token separately
        TokenAdvancedTo     NVARCHAR(MAX)  NULL
    );

    -- ── Indexes ──────────────────────────────────────────────

    -- Most common monitoring query: last N runs for a given process
    CREATE INDEX IX_SyncLog_FunctionName_StartedAt
        ON dbo.SyncLog (FunctionName, StartedAt DESC);

    -- Alert query: recent failed runs regardless of process
    CREATE INDEX IX_SyncLog_Status_StartedAt
        ON dbo.SyncLog (Status, StartedAt DESC);

    PRINT '  Created: SyncLog';
END
ELSE
    PRINT '  Already exists: SyncLog';
GO

-- ── Verification ──────────────────────────────────────────────

PRINT '';
PRINT 'Tables in ' + DB_NAME() + ':';
SELECT name AS TableName, create_date AS CreatedAt
FROM sys.tables
WHERE schema_id = SCHEMA_ID('dbo')
ORDER BY name;
GO
