-- =============================================================
-- 04_permissions_tables.sql
-- Phase 3: Permission tables for Exchange Online permissions sync.
-- Run AFTER 01_shared_tables.sql, 02_users_table.sql, 03_mailboxes_table.sql.
-- Connect to: db-m365permissions (NOT master)
--
-- Tables created:
--   MfcGroups               — registry of mfc-SHXXXXX distribution groups
--   MfcGroupMembers         — current membership of each MFC group (raw)
--   FullAccessPermissions   — Full Access on user mailboxes (raw trustee)
--   SendAsPermissions       — Send-As on all mailboxes (raw trustee)
--   SendOnBehalfPermissions — Send-on-Behalf on user mailboxes (raw trustee)
--   FolderPermissions       — Folder-level permissions on all mailboxes (raw trustee)
--
-- Design principles:
--   - Raw trustee identity stored as returned by the API/cmdlet.
--     Resolution to Users.UserId happens in a separate phase (Phase 4).
--   - Soft-delete on all tables (IsDeleted + DeletedAt).
--   - SyncSource: Baseline | Delta
--   - All permission tables FK → Mailboxes.ExchangeGuid (not UserId)
--     because permissions exist on the mailbox, not the Entra user.
--   - MfcGroupMembers has no IsDeleted — membership is replaced in full
--     on each sync (delete-all-then-reinsert per group). Historical
--     membership changes are tracked via SyncLog, not row-level audit.
-- =============================================================

PRINT 'Running 04_permissions_tables.sql in database: ' + DB_NAME();
PRINT '';

-- ══════════════════════════════════════════════════════════════
-- MfcGroups
-- Registry of mfc-SHXXXXX distribution groups.
-- One row per shared mailbox that has an associated MFC group.
-- Populated at baseline by matching group displayName prefix 'mfc-SH'
-- against known shared mailboxes in the Mailboxes table.
--
-- GroupObjectId  — Entra/Graph Object ID of the distribution group.
--                  Used as the identity for Graph delta queries on
--                  group membership changes.
-- ExchangeGuid   — FK to Mailboxes. Resolved at baseline load time
--                  by matching the SH-number suffix of the group name
--                  to the mailbox Alias (SH48329 → mfc-SH48329).
-- ══════════════════════════════════════════════════════════════

IF OBJECT_ID('dbo.MfcGroups', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.MfcGroups (

        -- Entra Object ID of the distribution group (Graph identity)
        GroupObjectId           UNIQUEIDENTIFIER    NOT NULL
            CONSTRAINT PK_MfcGroups PRIMARY KEY,

        -- Display name as returned by Graph (e.g. mfc-SH48329)
        DisplayName             NVARCHAR(256)       NOT NULL,

        -- Mail address of the group
        Mail                    NVARCHAR(256)       NULL,

        -- Link to the shared mailbox this group controls.
        -- Resolved from the SH-number suffix at baseline load.
        -- NULL if no matching mailbox found (should not happen in a clean tenant).
        ExchangeGuid            UNIQUEIDENTIFIER    NULL
            CONSTRAINT FK_MfcGroups_Mailboxes
                FOREIGN KEY REFERENCES dbo.Mailboxes(ExchangeGuid),

        -- Extracted SH identifier (e.g. SH48329) for readability and joins
        SharedMailboxAlias      NVARCHAR(20)        NULL,

        -- ── Soft-delete ──────────────────────────────────────
        IsDeleted               BIT                 NOT NULL
            CONSTRAINT DF_MfcGroups_IsDeleted DEFAULT 0,
        DeletedAt               DATETIME2           NULL,

        -- ── Audit ────────────────────────────────────────────
        FirstSeenAt             DATETIME2           NOT NULL
            CONSTRAINT DF_MfcGroups_FirstSeenAt DEFAULT SYSUTCDATETIME(),
        LastSyncedAt            DATETIME2           NOT NULL
            CONSTRAINT DF_MfcGroups_LastSyncedAt DEFAULT SYSUTCDATETIME(),
        SyncSource              NVARCHAR(20)        NOT NULL,
        LastSyncRunId           UNIQUEIDENTIFIER    NULL
    );

    PRINT '  Created: MfcGroups';
END
ELSE
    PRINT '  Already exists: MfcGroups';
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_MfcGroups_DisplayName' AND object_id = OBJECT_ID('dbo.MfcGroups'))
BEGIN
    CREATE INDEX IX_MfcGroups_DisplayName
        ON dbo.MfcGroups (DisplayName)
        WHERE IsDeleted = 0;
    PRINT '  Created: IX_MfcGroups_DisplayName';
END

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_MfcGroups_ExchangeGuid' AND object_id = OBJECT_ID('dbo.MfcGroups'))
BEGIN
    CREATE INDEX IX_MfcGroups_ExchangeGuid
        ON dbo.MfcGroups (ExchangeGuid)
        WHERE IsDeleted = 0;
    PRINT '  Created: IX_MfcGroups_ExchangeGuid';
END
GO

-- ══════════════════════════════════════════════════════════════
-- MfcGroupMembers
-- Current membership snapshot for each MFC group.
-- Source: EXO Admin API DistributionGroupMember endpoint.
--
-- Sync pattern: full replace per group (DELETE existing rows for
-- the group, then INSERT current members). This means no IsDeleted
-- column — absence of a row = not a member.
--
-- TrusteeRawIdentity — raw value returned by the API
--   (PrimarySmtpAddress from the DistributionGroupMember response).
-- ResolvedUserId     — populated in Phase 4 trustee resolution.
--
-- PermissionType column encodes why the member has access:
--   FullAccess       — member gets Full Access to the shared mailbox
--   SendOnBehalf     — member gets Send-on-Behalf on the shared mailbox
-- In this environment both permissions are granted by membership
-- in the same MFC group, so both rows are written per member.
-- ══════════════════════════════════════════════════════════════

IF OBJECT_ID('dbo.MfcGroupMembers', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.MfcGroupMembers (

        MfcGroupMemberId        BIGINT              NOT NULL
            CONSTRAINT PK_MfcGroupMembers PRIMARY KEY IDENTITY(1,1),

        -- FK to the controlling group
        GroupObjectId           UNIQUEIDENTIFIER    NOT NULL
            CONSTRAINT FK_MfcGroupMembers_MfcGroups
                FOREIGN KEY REFERENCES dbo.MfcGroups(GroupObjectId),

        -- The permission this membership represents
        -- FullAccess | SendOnBehalf
        PermissionType          NVARCHAR(20)        NOT NULL,

        -- Raw identity as returned by the Admin API
        TrusteeRawIdentity      NVARCHAR(256)       NOT NULL,

        -- Trustee display name (from API response, for readability)
        TrusteeDisplayName      NVARCHAR(256)       NULL,

        -- RecipientTypeDetails of the trustee (UserMailbox, MailUniversalSecurityGroup, etc.)
        TrusteeRecipientType    NVARCHAR(100)       NULL,

        -- Phase 4: resolved FK to Users.UserId (NULL until resolved)
        ResolvedUserId          UNIQUEIDENTIFIER    NULL,

        -- ── Audit ────────────────────────────────────────────
        FirstSeenAt             DATETIME2           NOT NULL
            CONSTRAINT DF_MfcGroupMembers_FirstSeenAt DEFAULT SYSUTCDATETIME(),
        LastSyncedAt            DATETIME2           NOT NULL
            CONSTRAINT DF_MfcGroupMembers_LastSyncedAt DEFAULT SYSUTCDATETIME(),
        LastSyncRunId           UNIQUEIDENTIFIER    NULL
    );

    PRINT '  Created: MfcGroupMembers';
END
ELSE
    PRINT '  Already exists: MfcGroupMembers';
GO

-- Unique constraint: one row per group + permission type + trustee
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UQ_MfcGroupMembers_GroupPermTrustee' AND object_id = OBJECT_ID('dbo.MfcGroupMembers'))
BEGIN
    CREATE UNIQUE INDEX UQ_MfcGroupMembers_GroupPermTrustee
        ON dbo.MfcGroupMembers (GroupObjectId, PermissionType, TrusteeRawIdentity);
    PRINT '  Created: UQ_MfcGroupMembers_GroupPermTrustee';
END

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_MfcGroupMembers_GroupObjectId' AND object_id = OBJECT_ID('dbo.MfcGroupMembers'))
BEGIN
    CREATE INDEX IX_MfcGroupMembers_GroupObjectId
        ON dbo.MfcGroupMembers (GroupObjectId, PermissionType);
    PRINT '  Created: IX_MfcGroupMembers_GroupObjectId';
END

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_MfcGroupMembers_ResolvedUserId' AND object_id = OBJECT_ID('dbo.MfcGroupMembers'))
BEGIN
    CREATE INDEX IX_MfcGroupMembers_ResolvedUserId
        ON dbo.MfcGroupMembers (ResolvedUserId)
        WHERE ResolvedUserId IS NOT NULL;
    PRINT '  Created: IX_MfcGroupMembers_ResolvedUserId';
END
GO

-- ══════════════════════════════════════════════════════════════
-- FullAccessPermissions
-- Full Access (FullAccess) permissions on USER mailboxes only.
-- Shared mailbox Full Access is tracked via MfcGroupMembers.
-- Source: EXO Get-EXOMailboxPermission
--
-- IsInherited  — permissions inherited from parent (e.g. SELF) are
--                filtered out at collection time; this column records
--                what EXO returned for transparency.
-- AutoMapping  — whether Outlook auto-maps the mailbox. Important
--                for auditing and helpdesk troubleshooting.
-- ══════════════════════════════════════════════════════════════

IF OBJECT_ID('dbo.FullAccessPermissions', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.FullAccessPermissions (

        FullAccessPermissionId  BIGINT              NOT NULL
            CONSTRAINT PK_FullAccessPermissions PRIMARY KEY IDENTITY(1,1),

        -- The mailbox on which the permission is granted
        ExchangeGuid            UNIQUEIDENTIFIER    NOT NULL
            CONSTRAINT FK_FullAccessPermissions_Mailboxes
                FOREIGN KEY REFERENCES dbo.Mailboxes(ExchangeGuid),

        -- Raw trustee as returned by Get-EXOMailboxPermission
        -- Typically in DOMAIN\user or UPN format
        TrusteeRawIdentity      NVARCHAR(256)       NOT NULL,

        -- Phase 4: resolved FK to Users.UserId
        ResolvedUserId          UNIQUEIDENTIFIER    NULL,

        IsInherited             BIT                 NULL,
        AutoMapping             BIT                 NULL,

        -- ── Soft-delete ──────────────────────────────────────
        IsDeleted               BIT                 NOT NULL
            CONSTRAINT DF_FullAccessPermissions_IsDeleted DEFAULT 0,
        DeletedAt               DATETIME2           NULL,

        -- ── Audit ────────────────────────────────────────────
        FirstSeenAt             DATETIME2           NOT NULL
            CONSTRAINT DF_FullAccessPermissions_FirstSeenAt DEFAULT SYSUTCDATETIME(),
        LastSyncedAt            DATETIME2           NOT NULL
            CONSTRAINT DF_FullAccessPermissions_LastSyncedAt DEFAULT SYSUTCDATETIME(),
        SyncSource              NVARCHAR(20)        NOT NULL,
        LastSyncRunId           UNIQUEIDENTIFIER    NULL
    );

    PRINT '  Created: FullAccessPermissions';
END
ELSE
    PRINT '  Already exists: FullAccessPermissions';
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UQ_FullAccess_MailboxTrustee' AND object_id = OBJECT_ID('dbo.FullAccessPermissions'))
BEGIN
    CREATE UNIQUE INDEX UQ_FullAccess_MailboxTrustee
        ON dbo.FullAccessPermissions (ExchangeGuid, TrusteeRawIdentity)
        WHERE IsDeleted = 0;
    PRINT '  Created: UQ_FullAccess_MailboxTrustee';
END

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_FullAccess_ExchangeGuid' AND object_id = OBJECT_ID('dbo.FullAccessPermissions'))
BEGIN
    CREATE INDEX IX_FullAccess_ExchangeGuid
        ON dbo.FullAccessPermissions (ExchangeGuid)
        WHERE IsDeleted = 0;
    PRINT '  Created: IX_FullAccess_ExchangeGuid';
END

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_FullAccess_ResolvedUserId' AND object_id = OBJECT_ID('dbo.FullAccessPermissions'))
BEGIN
    CREATE INDEX IX_FullAccess_ResolvedUserId
        ON dbo.FullAccessPermissions (ResolvedUserId)
        WHERE ResolvedUserId IS NOT NULL AND IsDeleted = 0;
    PRINT '  Created: IX_FullAccess_ResolvedUserId';
END
GO

-- ══════════════════════════════════════════════════════════════
-- SendAsPermissions
-- Send-As permissions on ALL mailboxes (user and shared).
-- Source: EXO Get-EXORecipientPermission
-- Not managed via MFC groups — collected directly for all mailbox types.
--
-- AccessControlType — typically 'Allow'; 'Deny' entries are rare
--                     but must be captured for accurate audit reporting.
-- ══════════════════════════════════════════════════════════════

IF OBJECT_ID('dbo.SendAsPermissions', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.SendAsPermissions (

        SendAsPermissionId      BIGINT              NOT NULL
            CONSTRAINT PK_SendAsPermissions PRIMARY KEY IDENTITY(1,1),

        ExchangeGuid            UNIQUEIDENTIFIER    NOT NULL
            CONSTRAINT FK_SendAsPermissions_Mailboxes
                FOREIGN KEY REFERENCES dbo.Mailboxes(ExchangeGuid),

        -- Raw trustee as returned by Get-EXORecipientPermission
        TrusteeRawIdentity      NVARCHAR(256)       NOT NULL,

        -- Phase 4: resolved FK to Users.UserId
        ResolvedUserId          UNIQUEIDENTIFIER    NULL,

        -- Allow | Deny
        AccessControlType       NVARCHAR(20)        NULL,

        -- ── Soft-delete ──────────────────────────────────────
        IsDeleted               BIT                 NOT NULL
            CONSTRAINT DF_SendAsPermissions_IsDeleted DEFAULT 0,
        DeletedAt               DATETIME2           NULL,

        -- ── Audit ────────────────────────────────────────────
        FirstSeenAt             DATETIME2           NOT NULL
            CONSTRAINT DF_SendAsPermissions_FirstSeenAt DEFAULT SYSUTCDATETIME(),
        LastSyncedAt            DATETIME2           NOT NULL
            CONSTRAINT DF_SendAsPermissions_LastSyncedAt DEFAULT SYSUTCDATETIME(),
        SyncSource              NVARCHAR(20)        NOT NULL,
        LastSyncRunId           UNIQUEIDENTIFIER    NULL
    );

    PRINT '  Created: SendAsPermissions';
END
ELSE
    PRINT '  Already exists: SendAsPermissions';
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UQ_SendAs_MailboxTrustee' AND object_id = OBJECT_ID('dbo.SendAsPermissions'))
BEGIN
    CREATE UNIQUE INDEX UQ_SendAs_MailboxTrustee
        ON dbo.SendAsPermissions (ExchangeGuid, TrusteeRawIdentity)
        WHERE IsDeleted = 0;
    PRINT '  Created: UQ_SendAs_MailboxTrustee';
END

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SendAs_ExchangeGuid' AND object_id = OBJECT_ID('dbo.SendAsPermissions'))
BEGIN
    CREATE INDEX IX_SendAs_ExchangeGuid
        ON dbo.SendAsPermissions (ExchangeGuid)
        WHERE IsDeleted = 0;
    PRINT '  Created: IX_SendAs_ExchangeGuid';
END

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SendAs_ResolvedUserId' AND object_id = OBJECT_ID('dbo.SendAsPermissions'))
BEGIN
    CREATE INDEX IX_SendAs_ResolvedUserId
        ON dbo.SendAsPermissions (ResolvedUserId)
        WHERE ResolvedUserId IS NOT NULL AND IsDeleted = 0;
    PRINT '  Created: IX_SendAs_ResolvedUserId';
END
GO

-- ══════════════════════════════════════════════════════════════
-- SendOnBehalfPermissions
-- Send-on-Behalf permissions on USER mailboxes only.
-- Shared mailbox Send-on-Behalf is tracked via MfcGroupMembers.
-- Source: EXO Get-EXOMailbox (GrantSendOnBehalfTo attribute).
--
-- Note: The Mailboxes table currently stores GrantSendOnBehalfTo
-- as a denormalised semicolon string. This table is the normalised
-- form — one row per trustee — promoted from that field.
-- The Mailboxes.GrantSendOnBehalfTo column can be deprecated
-- once this table is fully populated.
-- ══════════════════════════════════════════════════════════════

IF OBJECT_ID('dbo.SendOnBehalfPermissions', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.SendOnBehalfPermissions (

        SendOnBehalfPermissionId BIGINT             NOT NULL
            CONSTRAINT PK_SendOnBehalfPermissions PRIMARY KEY IDENTITY(1,1),

        ExchangeGuid            UNIQUEIDENTIFIER    NOT NULL
            CONSTRAINT FK_SendOnBehalfPermissions_Mailboxes
                FOREIGN KEY REFERENCES dbo.Mailboxes(ExchangeGuid),

        -- Raw trustee as returned by EXO (display name or UPN — EXO
        -- returns the DN form; normalise to SMTP at collection time where possible)
        TrusteeRawIdentity      NVARCHAR(256)       NOT NULL,

        -- Phase 4: resolved FK to Users.UserId
        ResolvedUserId          UNIQUEIDENTIFIER    NULL,

        -- ── Soft-delete ──────────────────────────────────────
        IsDeleted               BIT                 NOT NULL
            CONSTRAINT DF_SendOnBehalfPermissions_IsDeleted DEFAULT 0,
        DeletedAt               DATETIME2           NULL,

        -- ── Audit ────────────────────────────────────────────
        FirstSeenAt             DATETIME2           NOT NULL
            CONSTRAINT DF_SendOnBehalfPermissions_FirstSeenAt DEFAULT SYSUTCDATETIME(),
        LastSyncedAt            DATETIME2           NOT NULL
            CONSTRAINT DF_SendOnBehalfPermissions_LastSyncedAt DEFAULT SYSUTCDATETIME(),
        SyncSource              NVARCHAR(20)        NOT NULL,
        LastSyncRunId           UNIQUEIDENTIFIER    NULL
    );

    PRINT '  Created: SendOnBehalfPermissions';
END
ELSE
    PRINT '  Already exists: SendOnBehalfPermissions';
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UQ_SendOnBehalf_MailboxTrustee' AND object_id = OBJECT_ID('dbo.SendOnBehalfPermissions'))
BEGIN
    CREATE UNIQUE INDEX UQ_SendOnBehalf_MailboxTrustee
        ON dbo.SendOnBehalfPermissions (ExchangeGuid, TrusteeRawIdentity)
        WHERE IsDeleted = 0;
    PRINT '  Created: UQ_SendOnBehalf_MailboxTrustee';
END

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SendOnBehalf_ExchangeGuid' AND object_id = OBJECT_ID('dbo.SendOnBehalfPermissions'))
BEGIN
    CREATE INDEX IX_SendOnBehalf_ExchangeGuid
        ON dbo.SendOnBehalfPermissions (ExchangeGuid)
        WHERE IsDeleted = 0;
    PRINT '  Created: IX_SendOnBehalf_ExchangeGuid';
END

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_SendOnBehalf_ResolvedUserId' AND object_id = OBJECT_ID('dbo.SendOnBehalfPermissions'))
BEGIN
    CREATE INDEX IX_SendOnBehalf_ResolvedUserId
        ON dbo.SendOnBehalfPermissions (ResolvedUserId)
        WHERE ResolvedUserId IS NOT NULL AND IsDeleted = 0;
    PRINT '  Created: IX_SendOnBehalf_ResolvedUserId';
END
GO

-- ══════════════════════════════════════════════════════════════
-- FolderPermissions
-- Folder-level permissions on ALL mailboxes, all visible folders.
-- Source: EXO Admin API MailboxFolderPermission endpoint
--         (Get-MailboxFolderPermission via REST).
--
-- FolderPath     — full folder path as returned by the API,
--                  e.g. '\Calendar', '\Inbox', '\Inbox\Reports'
-- FolderName     — display name of the folder (from API response)
-- AccessRights   — role or granular rights, stored as returned
--                  (may be comma-separated list, e.g. 'Editor' or
--                  'CreateItems,DeleteOwnedItems,FolderVisible')
-- SharingPermissionFlags — Calendar-only flags:
--                  ViewPrivateItems, ReceiveCopiesOfMeetingMessages
--                  NULL for non-calendar folders.
-- IsDefault      — true for well-known trustees: Default, Anonymous
--                  Filtered out at collection time unless non-None rights.
--
-- Sync pattern: full replace per mailbox per sync run.
--   DELETE WHERE ExchangeGuid = @guid AND IsDeleted = 0
--   then INSERT current snapshot.
--   Soft-delete is set on the DELETE step rather than physical delete
--   to preserve audit history.
-- ══════════════════════════════════════════════════════════════

IF OBJECT_ID('dbo.FolderPermissions', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.FolderPermissions (

        FolderPermissionId      BIGINT              NOT NULL
            CONSTRAINT PK_FolderPermissions PRIMARY KEY IDENTITY(1,1),

        ExchangeGuid            UNIQUEIDENTIFIER    NOT NULL
            CONSTRAINT FK_FolderPermissions_Mailboxes
                FOREIGN KEY REFERENCES dbo.Mailboxes(ExchangeGuid),

        -- Folder identity: '\Calendar', '\Inbox', '\Inbox\Reports'
        FolderPath              NVARCHAR(512)       NOT NULL,

        -- Display name: 'Calendar', 'Inbox'
        FolderName              NVARCHAR(256)       NULL,

        -- Raw trustee as returned by the Admin API
        -- May be UPN, display name, 'Default', or 'Anonymous'
        TrusteeRawIdentity      NVARCHAR(256)       NOT NULL,

        -- Phase 4: resolved FK to Users.UserId
        -- NULL for Default/Anonymous trustees (not resolvable)
        ResolvedUserId          UNIQUEIDENTIFIER    NULL,

        -- Role or granular rights string as returned by the API
        AccessRights            NVARCHAR(512)       NOT NULL,

        -- Calendar-only: ViewPrivateItems, ReceiveCopiesOfMeetingMessages
        -- NULL for non-calendar folders
        SharingPermissionFlags  NVARCHAR(256)       NULL,

        -- True for Default and Anonymous well-known trustees
        IsDefaultTrustee        BIT                 NOT NULL
            CONSTRAINT DF_FolderPermissions_IsDefaultTrustee DEFAULT 0,

        -- ── Soft-delete ──────────────────────────────────────
        IsDeleted               BIT                 NOT NULL
            CONSTRAINT DF_FolderPermissions_IsDeleted DEFAULT 0,
        DeletedAt               DATETIME2           NULL,

        -- ── Audit ────────────────────────────────────────────
        FirstSeenAt             DATETIME2           NOT NULL
            CONSTRAINT DF_FolderPermissions_FirstSeenAt DEFAULT SYSUTCDATETIME(),
        LastSyncedAt            DATETIME2           NOT NULL
            CONSTRAINT DF_FolderPermissions_LastSyncedAt DEFAULT SYSUTCDATETIME(),
        SyncSource              NVARCHAR(20)        NOT NULL,
        LastSyncRunId           UNIQUEIDENTIFIER    NULL
    );

    PRINT '  Created: FolderPermissions';
END
ELSE
    PRINT '  Already exists: FolderPermissions';
GO

-- Unique: one active permission row per mailbox + folder + trustee
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UQ_FolderPerm_MailboxFolderTrustee' AND object_id = OBJECT_ID('dbo.FolderPermissions'))
BEGIN
    CREATE UNIQUE INDEX UQ_FolderPerm_MailboxFolderTrustee
        ON dbo.FolderPermissions (ExchangeGuid, FolderPath, TrusteeRawIdentity)
        WHERE IsDeleted = 0;
    PRINT '  Created: UQ_FolderPerm_MailboxFolderTrustee';
END

-- Primary reporting join: all folder permissions on a given mailbox
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_FolderPerm_ExchangeGuid' AND object_id = OBJECT_ID('dbo.FolderPermissions'))
BEGIN
    CREATE INDEX IX_FolderPerm_ExchangeGuid
        ON dbo.FolderPermissions (ExchangeGuid, FolderPath)
        WHERE IsDeleted = 0;
    PRINT '  Created: IX_FolderPerm_ExchangeGuid';
END

-- "Where does this user have folder access?" — common audit query
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_FolderPerm_ResolvedUserId' AND object_id = OBJECT_ID('dbo.FolderPermissions'))
BEGIN
    CREATE INDEX IX_FolderPerm_ResolvedUserId
        ON dbo.FolderPermissions (ResolvedUserId)
        WHERE ResolvedUserId IS NOT NULL AND IsDeleted = 0;
    PRINT '  Created: IX_FolderPerm_ResolvedUserId';
END

-- Calendar permissions specifically — frequently queried in financial services audits
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_FolderPerm_Calendar' AND object_id = OBJECT_ID('dbo.FolderPermissions'))
BEGIN
    CREATE INDEX IX_FolderPerm_Calendar
        ON dbo.FolderPermissions (ExchangeGuid, TrusteeRawIdentity)
        WHERE FolderName = 'Calendar' AND IsDeleted = 0;
    PRINT '  Created: IX_FolderPerm_Calendar';
END
GO

-- ══════════════════════════════════════════════════════════════
-- Extend SyncLog with Phase 3 permission counters
-- ══════════════════════════════════════════════════════════════

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.SyncLog')
      AND name = 'MfcGroupsProcessed'
)
BEGIN
    ALTER TABLE dbo.SyncLog ADD MfcGroupsProcessed       INT NOT NULL DEFAULT 0;
    ALTER TABLE dbo.SyncLog ADD MfcMembersProcessed      INT NOT NULL DEFAULT 0;
    ALTER TABLE dbo.SyncLog ADD FullAccessProcessed      INT NOT NULL DEFAULT 0;
    ALTER TABLE dbo.SyncLog ADD SendAsProcessed          INT NOT NULL DEFAULT 0;
    ALTER TABLE dbo.SyncLog ADD SendOnBehalfProcessed    INT NOT NULL DEFAULT 0;
    ALTER TABLE dbo.SyncLog ADD FolderPermissionsProcessed INT NOT NULL DEFAULT 0;
    ALTER TABLE dbo.SyncLog ADD MailboxesSkipped         INT NOT NULL DEFAULT 0;
    PRINT '  Extended: SyncLog (added Phase 3 permission counter columns)';
END
ELSE
    PRINT '  Already extended: SyncLog (Phase 3 columns present)';
GO

-- ══════════════════════════════════════════════════════════════
-- Reporting views
-- ══════════════════════════════════════════════════════════════

-- All permissions for a given mailbox (union across all types)
-- Useful as the base for Power BI and audit reports.
IF OBJECT_ID('dbo.vw_AllMailboxPermissions', 'V') IS NOT NULL
    DROP VIEW dbo.vw_AllMailboxPermissions;
GO
CREATE VIEW dbo.vw_AllMailboxPermissions AS

    -- Full Access — user mailboxes
    SELECT
        m.ExchangeGuid,
        m.PrimarySmtpAddress,
        m.DisplayName          AS MailboxDisplayName,
        m.MailboxType,
        'FullAccess'           AS PermissionType,
        NULL                   AS FolderPath,
        p.TrusteeRawIdentity,
        p.ResolvedUserId,
        NULL                   AS AccessRights,
        NULL                   AS SharingPermissionFlags,
        p.IsDeleted,
        p.LastSyncedAt
    FROM dbo.FullAccessPermissions p
    JOIN dbo.Mailboxes m ON m.ExchangeGuid = p.ExchangeGuid
    WHERE p.IsDeleted = 0 AND m.IsDeleted = 0

    UNION ALL

    -- Full Access — shared mailboxes via MFC group
    SELECT
        mb.ExchangeGuid,
        mb.PrimarySmtpAddress,
        mb.DisplayName,
        mb.MailboxType,
        'FullAccess'           AS PermissionType,
        NULL                   AS FolderPath,
        mm.TrusteeRawIdentity,
        mm.ResolvedUserId,
        NULL                   AS AccessRights,
        NULL                   AS SharingPermissionFlags,
        0                      AS IsDeleted,
        mm.LastSyncedAt
    FROM dbo.MfcGroupMembers mm
    JOIN dbo.MfcGroups       mg ON mg.GroupObjectId = mm.GroupObjectId
    JOIN dbo.Mailboxes       mb ON mb.ExchangeGuid  = mg.ExchangeGuid
    WHERE mm.PermissionType = 'FullAccess'
      AND mg.IsDeleted = 0
      AND mb.IsDeleted = 0

    UNION ALL

    -- Send-As — all mailboxes
    SELECT
        m.ExchangeGuid,
        m.PrimarySmtpAddress,
        m.DisplayName,
        m.MailboxType,
        'SendAs'               AS PermissionType,
        NULL                   AS FolderPath,
        p.TrusteeRawIdentity,
        p.ResolvedUserId,
        NULL                   AS AccessRights,
        NULL                   AS SharingPermissionFlags,
        p.IsDeleted,
        p.LastSyncedAt
    FROM dbo.SendAsPermissions p
    JOIN dbo.Mailboxes m ON m.ExchangeGuid = p.ExchangeGuid
    WHERE p.IsDeleted = 0 AND m.IsDeleted = 0

    UNION ALL

    -- Send-on-Behalf — user mailboxes
    SELECT
        m.ExchangeGuid,
        m.PrimarySmtpAddress,
        m.DisplayName,
        m.MailboxType,
        'SendOnBehalf'         AS PermissionType,
        NULL                   AS FolderPath,
        p.TrusteeRawIdentity,
        p.ResolvedUserId,
        NULL                   AS AccessRights,
        NULL                   AS SharingPermissionFlags,
        p.IsDeleted,
        p.LastSyncedAt
    FROM dbo.SendOnBehalfPermissions p
    JOIN dbo.Mailboxes m ON m.ExchangeGuid = p.ExchangeGuid
    WHERE p.IsDeleted = 0 AND m.IsDeleted = 0

    UNION ALL

    -- Send-on-Behalf — shared mailboxes via MFC group
    SELECT
        mb.ExchangeGuid,
        mb.PrimarySmtpAddress,
        mb.DisplayName,
        mb.MailboxType,
        'SendOnBehalf'         AS PermissionType,
        NULL                   AS FolderPath,
        mm.TrusteeRawIdentity,
        mm.ResolvedUserId,
        NULL                   AS AccessRights,
        NULL                   AS SharingPermissionFlags,
        0                      AS IsDeleted,
        mm.LastSyncedAt
    FROM dbo.MfcGroupMembers mm
    JOIN dbo.MfcGroups       mg ON mg.GroupObjectId = mm.GroupObjectId
    JOIN dbo.Mailboxes       mb ON mb.ExchangeGuid  = mg.ExchangeGuid
    WHERE mm.PermissionType = 'SendOnBehalf'
      AND mg.IsDeleted = 0
      AND mb.IsDeleted = 0

    UNION ALL

    -- Folder Permissions — all mailboxes
    SELECT
        m.ExchangeGuid,
        m.PrimarySmtpAddress,
        m.DisplayName,
        m.MailboxType,
        'FolderPermission'     AS PermissionType,
        p.FolderPath,
        p.TrusteeRawIdentity,
        p.ResolvedUserId,
        p.AccessRights,
        p.SharingPermissionFlags,
        p.IsDeleted,
        p.LastSyncedAt
    FROM dbo.FolderPermissions p
    JOIN dbo.Mailboxes m ON m.ExchangeGuid = p.ExchangeGuid
    WHERE p.IsDeleted = 0 AND m.IsDeleted = 0;
GO
PRINT '  Created: vw_AllMailboxPermissions';
GO

-- MFC group membership with mailbox context — operational view
IF OBJECT_ID('dbo.vw_MfcGroupMembership', 'V') IS NOT NULL
    DROP VIEW dbo.vw_MfcGroupMembership;
GO
CREATE VIEW dbo.vw_MfcGroupMembership AS
    SELECT
        mg.GroupObjectId,
        mg.DisplayName          AS GroupDisplayName,
        mg.SharedMailboxAlias,
        mb.PrimarySmtpAddress   AS SharedMailboxSmtp,
        mb.DisplayName          AS SharedMailboxDisplayName,
        mm.PermissionType,
        mm.TrusteeRawIdentity,
        mm.TrusteeDisplayName,
        mm.ResolvedUserId,
        u.UserPrincipalName     AS ResolvedUPN,
        mm.LastSyncedAt
    FROM dbo.MfcGroupMembers mm
    JOIN dbo.MfcGroups   mg ON mg.GroupObjectId = mm.GroupObjectId
    LEFT JOIN dbo.Mailboxes mb ON mb.ExchangeGuid = mg.ExchangeGuid
    LEFT JOIN dbo.Users      u  ON u.UserId       = mm.ResolvedUserId
    WHERE mg.IsDeleted = 0;
GO
PRINT '  Created: vw_MfcGroupMembership';
GO

-- ══════════════════════════════════════════════════════════════
-- Verification
-- ══════════════════════════════════════════════════════════════

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
PRINT 'SyncLog columns (Phase 3 additions):';

SELECT c.column_id, c.name, t.name AS DataType, c.is_nullable
FROM sys.columns c
JOIN sys.types   t ON t.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('dbo.SyncLog')
  AND c.name IN (
    'MfcGroupsProcessed','MfcMembersProcessed',
    'FullAccessProcessed','SendAsProcessed',
    'SendOnBehalfProcessed','FolderPermissionsProcessed',
    'MailboxesSkipped'
  )
ORDER BY c.column_id;
GO
