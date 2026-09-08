-- =============================================================
-- 05_add_permissions_baselined.sql
-- Adds PermissionsBaselinedAt to dbo.Mailboxes.
--
-- This column tracks whether Invoke-PermissionBaselineLoad (or the
-- mini-baseline in Invoke-PermissionDeltaSync Phase 0) has collected
-- permissions for each mailbox. NULL means not yet collected.
--
-- Invoke-PermissionDeltaSync Phase 0 queries:
--   SELECT ... FROM Mailboxes WHERE IsDeleted=0 AND PermissionsBaselinedAt IS NULL
-- and runs a per-mailbox permission collection for any result.
--
-- DEPLOYMENT NOTE:
-- If Invoke-PermissionBaselineLoad has ALREADY been run before this
-- migration, execute the UPDATE below to avoid re-baselining all
-- mailboxes on the next delta run. If the baseline has NOT yet run,
-- leave all rows as NULL — the delta script will baseline them.
-- =============================================================

PRINT 'Running 05_add_permissions_baselined.sql in database: ' + DB_NAME();

-- ── Add column (idempotent) ───────────────────────────────────

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.Mailboxes')
      AND name = 'PermissionsBaselinedAt'
)
BEGIN
    ALTER TABLE dbo.Mailboxes
    ADD PermissionsBaselinedAt DATETIME2 NULL;

    PRINT '  Added: Mailboxes.PermissionsBaselinedAt';
END
ELSE
    PRINT '  Already exists: Mailboxes.PermissionsBaselinedAt';
GO

-- ── Mark existing mailboxes as baselined if permissions data exists ─
-- Safe to run even if no permission data exists — the UPDATE
-- will affect 0 rows and the delta script will baseline them all.
-- Comment this block out if you want to force a full re-collection.

IF EXISTS (
    SELECT 1 FROM dbo.MfcGroupMembers
    UNION ALL
    SELECT 1 FROM dbo.FullAccessPermissions WHERE IsDeleted = 0
    UNION ALL
    SELECT 1 FROM dbo.FolderPermissions      WHERE IsDeleted = 0
)
BEGIN
    UPDATE dbo.Mailboxes
    SET PermissionsBaselinedAt = SYSUTCDATETIME()
    WHERE IsDeleted = 0
      AND PermissionsBaselinedAt IS NULL;

    PRINT '  Marked ' + CAST(@@ROWCOUNT AS NVARCHAR(10)) +
          ' existing mailbox(es) as baselined (permission data already present)';
END
ELSE
    PRINT '  No permission data found — all mailboxes left as NULL (will be baselined by delta sync)';
GO

-- ── Verification ─────────────────────────────────────────────

SELECT
    COUNT(*)                                                    AS TotalActive,
    SUM(CASE WHEN PermissionsBaselinedAt IS NOT NULL THEN 1 ELSE 0 END) AS Baselined,
    SUM(CASE WHEN PermissionsBaselinedAt IS NULL     THEN 1 ELSE 0 END) AS PendingBaseline
FROM dbo.Mailboxes
WHERE IsDeleted = 0;
GO
