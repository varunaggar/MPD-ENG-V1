-- =============================================================
-- 02_users_table.sql
-- Destination table for the user information sync processes.
-- Run AFTER 01_shared_tables.sql.
-- Connect to: db-m365permissions (NOT master)
-- =============================================================

PRINT 'Creating Users table in database: ' + DB_NAME();
PRINT '';

-- ──────────────────────────────────────────────────────────────
-- Users
-- A record of every user in the Entra ID tenant.
-- Populated by Invoke-UserBaselineLoad (full snapshot)
-- and kept current by Invoke-UserDeltaSync (incremental).
--
-- Design decisions:
--   PK = UserId (Entra Object ID, GUID) — immutable.
--        UPNs change with name changes and domain renames.
--        Object IDs never change for the lifetime of the user.
--
--   Soft-delete only — IsDeleted flag rather than DELETE.
--        Permission history that references this user remains
--        valid and queryable even after the user leaves the org.
--
--   Filtered indexes on IsDeleted = 0 — the hot-path queries
--        (find active user by UPN or mail) only scan live rows.
-- ──────────────────────────────────────────────────────────────

IF OBJECT_ID('dbo.Users', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Users (

        -- ── Primary key ─────────────────────────────────────
        -- Entra Object ID. Immutable. Never changes even if
        -- UPN, mail, display name, or tenant change.
        UserId                  UNIQUEIDENTIFIER    NOT NULL
            CONSTRAINT PK_Users PRIMARY KEY,

        -- ── Core identity (can change over time) ────────────
        UserPrincipalName       NVARCHAR(256)       NOT NULL,
        DisplayName             NVARCHAR(256)       NULL,

        -- Primary SMTP address. Used as the trustee identifier
        -- in all five permission tables. Null for unlicensed users.
        Mail                    NVARCHAR(256)       NULL,

        -- ── Account state ───────────────────────────────────
        AccountEnabled          BIT                 NOT NULL
            CONSTRAINT DF_Users_AccountEnabled DEFAULT 1,

        -- Member = internal user, Guest = external collaborator (B2B)
        UserType                NVARCHAR(20)        NULL,

        -- NULL means cloud-only (no on-prem sync).
        -- 1 means the account is synced from on-prem Active Directory.
        OnPremisesSyncEnabled   BIT                 NULL,

        -- ── Organisational attributes ────────────────────────
        -- Used for Row-Level Security in Power BI and reporting filters.
        Department              NVARCHAR(128)       NULL,
        JobTitle                NVARCHAR(128)       NULL,

        -- ── Lifecycle timestamps ─────────────────────────────
        -- When Entra itself created the user account
        EntraCreatedDateTime    DATETIME2           NULL,

        -- When this row first appeared in our database
        FirstSeenAt             DATETIME2           NOT NULL
            CONSTRAINT DF_Users_FirstSeenAt DEFAULT SYSUTCDATETIME(),

        -- Last time any sync process touched this row
        LastSyncedAt            DATETIME2           NOT NULL
            CONSTRAINT DF_Users_LastSyncedAt DEFAULT SYSUTCDATETIME(),

        -- Last time we detected a change to any attribute
        -- (NULL = never changed since first seen)
        LastModifiedAt          DATETIME2           NULL,

        -- ── Soft-delete ──────────────────────────────────────
        IsDeleted               BIT                 NOT NULL
            CONSTRAINT DF_Users_IsDeleted DEFAULT 0,

        -- Populated when the user is removed from Entra
        DeletedAt               DATETIME2           NULL,

        -- ── Audit trail ──────────────────────────────────────
        -- Which process last wrote to this row
        SyncSource              NVARCHAR(20)        NOT NULL,
            -- Baseline = written by Invoke-UserBaselineLoad
            -- Delta    = written by Invoke-UserDeltaSync

        -- FK to SyncLog.RunId — which specific execution last touched this row
        LastSyncRunId           UNIQUEIDENTIFIER    NULL
    );

    PRINT '  Created: Users table';
END
ELSE
BEGIN
    PRINT '  Already exists: Users table';
END
GO

-- ── Indexes ───────────────────────────────────────────────────

-- Filtered index: look up active users by UPN (login queries, permission joins)
-- WHERE IsDeleted = 0 excludes deleted users and keeps the index small
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Users_UPN' AND object_id = OBJECT_ID('dbo.Users'))
BEGIN
    CREATE INDEX IX_Users_UPN
        ON dbo.Users (UserPrincipalName)
        WHERE IsDeleted = 0;
    PRINT '  Created: IX_Users_UPN';
END

-- Filtered index: look up active users by Mail (permission trustee joins)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Users_Mail' AND object_id = OBJECT_ID('dbo.Users'))
BEGIN
    CREATE INDEX IX_Users_Mail
        ON dbo.Users (Mail)
        WHERE IsDeleted = 0;
    PRINT '  Created: IX_Users_Mail';
END

-- Index to support soft-delete queries
-- e.g. "how many users were deleted this week?"
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Users_IsDeleted' AND object_id = OBJECT_ID('dbo.Users'))
BEGIN
    CREATE INDEX IX_Users_IsDeleted
        ON dbo.Users (IsDeleted, DeletedAt);
    PRINT '  Created: IX_Users_IsDeleted';
END

-- Index to support monitoring queries
-- e.g. "which users were not synced in the last hour?"
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Users_LastSyncedAt' AND object_id = OBJECT_ID('dbo.Users'))
BEGIN
    CREATE INDEX IX_Users_LastSyncedAt
        ON dbo.Users (LastSyncedAt);
    PRINT '  Created: IX_Users_LastSyncedAt';
END
GO

-- ── Useful reporting views ────────────────────────────────────

-- View: active users (the set used for most permission reporting)
IF OBJECT_ID('dbo.vw_ActiveUsers', 'V') IS NOT NULL
    DROP VIEW dbo.vw_ActiveUsers;
GO
CREATE VIEW dbo.vw_ActiveUsers AS
    SELECT
        UserId,
        UserPrincipalName,
        DisplayName,
        Mail,
        AccountEnabled,
        UserType,
        OnPremisesSyncEnabled,
        Department,
        JobTitle,
        EntraCreatedDateTime,
        FirstSeenAt,
        LastSyncedAt,
        LastModifiedAt,
        SyncSource
    FROM dbo.Users
    WHERE IsDeleted = 0;
GO
PRINT '  Created: vw_ActiveUsers';
GO

-- View: recently deleted users (useful for access revocation audits)
IF OBJECT_ID('dbo.vw_RecentlyDeletedUsers', 'V') IS NOT NULL
    DROP VIEW dbo.vw_RecentlyDeletedUsers;
GO
CREATE VIEW dbo.vw_RecentlyDeletedUsers AS
    SELECT
        UserId,
        UserPrincipalName,
        DisplayName,
        Mail,
        Department,
        DeletedAt,
        LastSyncedAt
    FROM dbo.Users
    WHERE IsDeleted = 1
      AND DeletedAt >= DATEADD(DAY, -90, SYSUTCDATETIME());
GO
PRINT '  Created: vw_RecentlyDeletedUsers';
GO

-- ── Verification ──────────────────────────────────────────────

PRINT '';
PRINT 'Tables and views in ' + DB_NAME() + ':';

SELECT
    o.name          AS ObjectName,
    o.type_desc     AS ObjectType,
    o.create_date   AS CreatedAt
FROM sys.objects o
WHERE o.schema_id = SCHEMA_ID('dbo')
  AND o.type IN ('U', 'V')
ORDER BY o.type_desc, o.name;
GO

PRINT '';
PRINT 'Column list for Users:';

SELECT
    c.column_id     AS Pos,
    c.name          AS ColumnName,
    t.name          AS DataType,
    c.max_length    AS MaxLen,
    c.is_nullable   AS Nullable
FROM sys.columns c
JOIN sys.types   t ON t.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('dbo.Users')
ORDER BY c.column_id;
GO
