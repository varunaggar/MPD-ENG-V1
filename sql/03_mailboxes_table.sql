-- =============================================================
-- 03_mailboxes_table.sql
-- Mailboxes table and SyncLog extension for mailbox metrics.
-- Run AFTER 01_shared_tables.sql and 02_users_table.sql.
-- Connect to: db-m365permissions (NOT master)
-- =============================================================

PRINT 'Running 03_mailboxes_table.sql in database: ' + DB_NAME();
PRINT '';

-- ──────────────────────────────────────────────────────────────
-- Extend SyncLog with mailbox-specific counters
-- These columns are added alongside the existing user columns.
-- NULL-safe: existing rows default to 0.
-- ──────────────────────────────────────────────────────────────

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.SyncLog')
      AND name = 'MailboxesInserted'
)
BEGIN
    ALTER TABLE dbo.SyncLog ADD MailboxesInserted    INT NOT NULL DEFAULT 0;
    ALTER TABLE dbo.SyncLog ADD MailboxesUpdated     INT NOT NULL DEFAULT 0;
    ALTER TABLE dbo.SyncLog ADD MailboxesSoftDeleted INT NOT NULL DEFAULT 0;
    ALTER TABLE dbo.SyncLog ADD MailboxesProcessed   INT NOT NULL DEFAULT 0;
    PRINT '  Extended: SyncLog (added mailbox counter columns)';
END
ELSE
    PRINT '  Already extended: SyncLog';
GO

-- ──────────────────────────────────────────────────────────────
-- Mailboxes
-- A record of every Exchange Online mailbox in the tenant.
-- Covers: UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox.
--
-- Design decisions:
--   PK = ExchangeGuid — the immutable Exchange identifier.
--        ExchangeGuid is assigned when the mailbox is created and
--        never changes, even if UPN or SMTP address changes.
--
--   FK = UserId (nullable) → Users.UserId
--        Populated for UserMailbox (the associated Entra user).
--        NULL for SharedMailbox, RoomMailbox, EquipmentMailbox
--        (these are not user accounts, they are resource objects).
--
--   GrantSendOnBehalfTo is stored as a denormalised semicolon-
--        separated string on this row. This is a temporary hold:
--        when Phase 3 (permission tables) is built, this data will
--        be promoted to the SendOnBehalfPermissions table and this
--        column can be dropped or kept as a cache.
--
--   Soft-delete only — IsDeleted flag. Exchange soft-deletes
--        mailboxes for 30 days before permanent removal. The weekly
--        baseline reconciliation catches permanent deletions.
--
--   WhenChangedUTC is the Exchange-side change marker used to drive
--        the delta sync via WhenChangedUTC filter queries.
-- ──────────────────────────────────────────────────────────────

IF OBJECT_ID('dbo.Mailboxes', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Mailboxes (

        -- ── Primary key ──────────────────────────────────────
        ExchangeGuid                UNIQUEIDENTIFIER    NOT NULL
            CONSTRAINT PK_Mailboxes PRIMARY KEY,

        -- ── Entra identity link ──────────────────────────────
        -- Maps to Users.UserId for UserMailbox type.
        -- NULL for Shared, Room, Equipment mailboxes.
        UserId                      UNIQUEIDENTIFIER    NULL
            CONSTRAINT FK_Mailboxes_Users
                FOREIGN KEY REFERENCES dbo.Users(UserId),

        -- ── Identity ─────────────────────────────────────────
        PrimarySmtpAddress          NVARCHAR(256)       NOT NULL,
        UserPrincipalName           NVARCHAR(256)       NULL,  -- NULL for non-user mailboxes
        DisplayName                 NVARCHAR(256)       NULL,
        Alias                       NVARCHAR(64)        NULL,

        -- ── Mailbox type ─────────────────────────────────────
        -- Exchange raw value: UserMailbox | SharedMailbox | RoomMailbox | EquipmentMailbox
        RecipientTypeDetails        NVARCHAR(50)        NOT NULL,

        -- Simplified label for reporting: User | Shared | Room | Equipment
        MailboxType                 NVARCHAR(20)        NOT NULL,

        -- ── Properties (all types) ───────────────────────────
        HiddenFromAddressLists      BIT                 NULL,
        LitigationHoldEnabled       BIT                 NULL,
        ArchiveStatus               NVARCHAR(30)        NULL,
        ForwardingAddress           NVARCHAR(256)       NULL,
        ForwardingSmtpAddress       NVARCHAR(256)       NULL,

        -- Denormalised Send-on-Behalf trustees (semicolon-separated UPNs).
        -- Populated from GrantSendOnBehalfTo multi-value EXO property.
        -- Promoted to SendOnBehalfPermissions table in Phase 3.
        GrantSendOnBehalfTo         NVARCHAR(MAX)       NULL,

        -- Whether the mailbox is synced from on-prem AD
        IsDirSynced                 BIT                 NULL,

        -- ── Lifecycle timestamps ─────────────────────────────
        WhenMailboxCreated          DATETIME2           NULL,

        -- Exchange's own last-modified timestamp.
        -- Used as the filter value in delta sync queries:
        --   Get-EXOMailbox -Filter "WhenChangedUTC -ge '$lastCheck'"
        WhenChangedUTC              DATETIME2           NULL,

        FirstSeenAt                 DATETIME2           NOT NULL
            CONSTRAINT DF_Mailboxes_FirstSeenAt DEFAULT SYSUTCDATETIME(),

        LastSyncedAt                DATETIME2           NOT NULL
            CONSTRAINT DF_Mailboxes_LastSyncedAt DEFAULT SYSUTCDATETIME(),

        LastModifiedAt              DATETIME2           NULL,

        -- ── Soft-delete ──────────────────────────────────────
        IsDeleted                   BIT                 NOT NULL
            CONSTRAINT DF_Mailboxes_IsDeleted DEFAULT 0,

        DeletedAt                   DATETIME2           NULL,

        -- ── Audit trail ──────────────────────────────────────
        -- Baseline = written by Invoke-MailboxBaselineLoad
        -- Delta    = written by Invoke-MailboxDeltaSync
        SyncSource                  NVARCHAR(20)        NOT NULL,

        LastSyncRunId               UNIQUEIDENTIFIER    NULL
    );

    PRINT '  Created: Mailboxes';
END
ELSE
    PRINT '  Already exists: Mailboxes';
GO

-- ── Indexes ───────────────────────────────────────────────────

-- Filtered: look up active mailboxes by primary SMTP (most common permission lookup)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Mailboxes_PrimarySmtp' AND object_id = OBJECT_ID('dbo.Mailboxes'))
BEGIN
    CREATE INDEX IX_Mailboxes_PrimarySmtp
        ON dbo.Mailboxes (PrimarySmtpAddress)
        WHERE IsDeleted = 0;
    PRINT '  Created: IX_Mailboxes_PrimarySmtp';
END

-- Filtered: look up active mailboxes by UPN (joins to Users)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Mailboxes_UPN' AND object_id = OBJECT_ID('dbo.Mailboxes'))
BEGIN
    CREATE INDEX IX_Mailboxes_UPN
        ON dbo.Mailboxes (UserPrincipalName)
        WHERE IsDeleted = 0;
    PRINT '  Created: IX_Mailboxes_UPN';
END

-- Filtered: join from permission tables to the owning user
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Mailboxes_UserId' AND object_id = OBJECT_ID('dbo.Mailboxes'))
BEGIN
    CREATE INDEX IX_Mailboxes_UserId
        ON dbo.Mailboxes (UserId)
        WHERE IsDeleted = 0;
    PRINT '  Created: IX_Mailboxes_UserId';
END

-- Type + delete status (supports "show me all active shared mailboxes" queries)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Mailboxes_Type' AND object_id = OBJECT_ID('dbo.Mailboxes'))
BEGIN
    CREATE INDEX IX_Mailboxes_Type
        ON dbo.Mailboxes (MailboxType, IsDeleted);
    PRINT '  Created: IX_Mailboxes_Type';
END

-- WhenChangedUTC: used in delta sync query to filter by last-changed time
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Mailboxes_WhenChangedUTC' AND object_id = OBJECT_ID('dbo.Mailboxes'))
BEGIN
    CREATE INDEX IX_Mailboxes_WhenChangedUTC
        ON dbo.Mailboxes (WhenChangedUTC);
    PRINT '  Created: IX_Mailboxes_WhenChangedUTC';
END

-- IsDeleted + DeletedAt (soft-delete audit queries)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Mailboxes_IsDeleted' AND object_id = OBJECT_ID('dbo.Mailboxes'))
BEGIN
    CREATE INDEX IX_Mailboxes_IsDeleted
        ON dbo.Mailboxes (IsDeleted, DeletedAt);
    PRINT '  Created: IX_Mailboxes_IsDeleted';
END
GO

-- ── Reporting views ───────────────────────────────────────────

IF OBJECT_ID('dbo.vw_ActiveMailboxes', 'V') IS NOT NULL
    DROP VIEW dbo.vw_ActiveMailboxes;
GO
CREATE VIEW dbo.vw_ActiveMailboxes AS
    SELECT
        m.ExchangeGuid,
        m.UserId,
        m.PrimarySmtpAddress,
        m.UserPrincipalName,
        m.DisplayName,
        m.MailboxType,
        m.RecipientTypeDetails,
        m.HiddenFromAddressLists,
        m.LitigationHoldEnabled,
        m.ArchiveStatus,
        m.ForwardingSmtpAddress,
        m.IsDirSynced,
        m.WhenMailboxCreated,
        m.WhenChangedUTC,
        m.FirstSeenAt,
        m.LastSyncedAt,
        -- Join to Users for display convenience
        u.UserPrincipalName     AS OwnerUPN,
        u.DisplayName           AS OwnerDisplayName,
        u.Department            AS OwnerDepartment,
        u.AccountEnabled        AS OwnerAccountEnabled
    FROM dbo.Mailboxes   m
    LEFT JOIN dbo.Users  u ON u.UserId = m.UserId AND u.IsDeleted = 0
    WHERE m.IsDeleted = 0;
GO
PRINT '  Created: vw_ActiveMailboxes';
GO

-- View: shared mailboxes with their Send-on-Behalf trustees
IF OBJECT_ID('dbo.vw_SharedMailboxes', 'V') IS NOT NULL
    DROP VIEW dbo.vw_SharedMailboxes;
GO
CREATE VIEW dbo.vw_SharedMailboxes AS
    SELECT
        ExchangeGuid,
        PrimarySmtpAddress,
        DisplayName,
        GrantSendOnBehalfTo,
        HiddenFromAddressLists,
        LitigationHoldEnabled,
        ArchiveStatus,
        IsDeleted,
        DeletedAt,
        LastSyncedAt
    FROM dbo.Mailboxes
    WHERE MailboxType = 'Shared';
GO
PRINT '  Created: vw_SharedMailboxes';
GO

-- ── Verification ──────────────────────────────────────────────

PRINT '';
PRINT 'Current schema in ' + DB_NAME() + ':';

SELECT
    o.name          AS ObjectName,
    o.type_desc     AS ObjectType,
    o.create_date   AS CreatedAt
FROM sys.objects o
WHERE o.schema_id = SCHEMA_ID('dbo')
  AND o.type IN ('U','V')
ORDER BY o.type_desc, o.name;
GO

PRINT '';
PRINT 'SyncLog columns:';

SELECT c.column_id, c.name, t.name AS DataType, c.is_nullable
FROM sys.columns c
JOIN sys.types   t ON t.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('dbo.SyncLog')
ORDER BY c.column_id;
GO
