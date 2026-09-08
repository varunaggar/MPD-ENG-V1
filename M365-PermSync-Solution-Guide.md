# M365 Permissions Sync — Solution Guide

**Version:** 1.0 (PoC)
**Environment:** Windows Server + Scheduled Tasks (production target: Azure Functions)
**Language:** PowerShell 7+
**Database:** Azure SQL Database (Serverless General Purpose)
**Tenant scale:** ~40,000 mailboxes

---

## 1. Purpose

This solution maintains a near-real-time database of Exchange Online mailbox permissions for compliance reporting at a large financial institution. It captures who can access which mailbox, by what means, and when the permission was granted — across Full Access, Send-As, Send-on-Behalf, and folder-level permissions.

The database feeds downstream reporting via Power BI DirectQuery and, eventually, a REST API / Static Web App.

---

## 2. Architecture Overview

### 2.1 Phased Build

| Phase | Scope | Status |
|-------|-------|--------|
| 1 | Users — Entra ID user directory | Complete |
| 2 | Mailboxes — Exchange Online mailbox inventory | Complete |
| 3 | Permissions — all permission types | Complete |
| 4 | Trustee Resolution — raw identity → Users FK | Not built |
| 5 | Reporting — Power BI, API, Static Web App | Out of scope |

### 2.2 Data Sources

| Source | Protocol | Auth | Used For |
|--------|----------|------|----------|
| Microsoft Graph REST API | HTTPS GET/POST | Certificate-based service principal | User sync, group discovery, folder enumeration, CloudAppEvents hunting |
| Exchange Online PowerShell V3 | PowerShell remoting | Certificate-based service principal | Mailbox sync, Full Access, Send-As |
| EXO Admin API v2.0 | HTTPS POST | Certificate-based service principal (Exchange.ManageAsAppV2) | MFC group members, folder permissions |
| Microsoft Defender CloudAppEvents | Graph Advanced Hunting | Certificate-based service principal (ThreatHunting.Read.All) | Permission delta detection |

### 2.3 Scheduling

| Script | Frequency | Trigger |
|--------|-----------|---------|
| Invoke-UserDeltaSync | Every 15 minutes | Scheduled Task |
| Invoke-MailboxDeltaSync | Every 15 minutes | Scheduled Task |
| Invoke-MfcGroupDeltaSync | Every 15 minutes | Scheduled Task |
| Invoke-PermissionDeltaSync | Every 15 minutes | Scheduled Task |
| Invoke-MailboxBaselineLoad | Weekly (e.g. Sunday 03:00) | Scheduled Task |
| Invoke-PermissionBaselineLoad | Monthly / on-demand | Scheduled Task |
| Invoke-UserBaselineLoad | Monthly / on-demand | Scheduled Task |

### 2.4 App Registration Permissions

| Scope | Permission | Type | Used By |
|-------|------------|------|---------|
| Microsoft Graph | User.Read.All | Application | User baseline/delta |
| Microsoft Graph | Group.Read.All | Application | MFC group discovery + delta |
| Microsoft Graph | Mail.ReadBasic.All | Application | Folder enumeration (Graph mailFolders) |
| Microsoft Graph | ThreatHunting.Read.All | Application | CloudAppEvents hunting query |
| Exchange Online | Exchange.ManageAsApp | Application | EXO PowerShell V3 connection |
| EXO Admin API (outlook.office365.com) | Exchange.ManageAsAppV2 | Application | Admin API (group members, folder permissions) |

The service principal also requires an Exchange Online management role assignment:

```powershell
New-ManagementRoleAssignment -App "<app-display-name>" -Role "Mail Recipients"
```

---

## 3. Environment-Specific Business Logic

### 3.1 Shared Mailbox Naming Convention

All shared mailboxes follow the naming pattern `SHXXXXX` (e.g. `SH48329`). The mailbox alias is `SH48329`.

### 3.2 MFC Group Model

For every shared mailbox, there is a corresponding Microsoft 365 distribution group named `mfc-SHXXXXX` (e.g. `mfc-SH48329`). Membership in this group grants two permissions on the shared mailbox:

- **Full Access** — the member can open and read the mailbox
- **Send-on-Behalf** — the member can send as "on behalf of" the mailbox

These permissions are NOT set directly on the mailbox (no `Add-MailboxPermission` calls). Instead, the group membership IS the permission grant. This means:

- Full Access and Send-on-Behalf for shared mailboxes are collected from **MFC group membership**, not from `Get-EXOMailboxPermission`.
- Full Access and Send-on-Behalf for user mailboxes are collected **directly** (they don't use MFC groups).
- **Send-As** is always direct for all mailbox types (never group-managed).
- **Folder permissions** are always direct for all mailbox types.

All shared mailboxes have an MFC group (no legacy exceptions).

### 3.3 Permission Collection Matrix

| Permission Type | User Mailbox Source | Shared Mailbox Source |
|----------------|--------------------|-----------------------|
| Full Access | Get-EXOMailboxPermission (EXO) | MFC group membership (Admin API) |
| Send-As | Get-EXORecipientPermission (EXO) | Get-EXORecipientPermission (EXO) |
| Send-on-Behalf | Mailboxes.GrantSendOnBehalfTo (DB) | MFC group membership (Admin API) |
| Folder Permissions | Graph mailFolders + Admin API | Graph mailFolders + Admin API |

### 3.4 Permission Delta Detection

| Permission Type | Delta Mechanism | Strategy |
|----------------|-----------------|----------|
| Full Access (direct) | CloudAppEvents: Add/Remove/Set-MailboxPermission | Re-fetch from EXO per affected mailbox |
| Send-As | CloudAppEvents: Add/Remove-RecipientPermission | Re-fetch from EXO per affected mailbox |
| Send-on-Behalf (direct) | CloudAppEvents: Set-Mailbox (GrantSendOnBehalfTo param) | Surgical replace from event payload |
| Folder Permissions | CloudAppEvents: Add/Set/Remove-MailboxFolderPermission | Surgical INSERT/soft-delete per event |
| MFC Group Membership | Graph /groups/delta | Re-fetch members per changed group (Admin API) |
| New Mailbox Permissions | Mailboxes.PermissionsBaselinedAt IS NULL | Mini-baseline per mailbox on next delta cycle |

---

## 4. Database Schema

### 4.1 Tables

**DeltaTokens** (`01_shared_tables.sql`)

Stores Graph delta tokens and EXO timestamps. One row per named token.

| Column | Type | Description |
|--------|------|-------------|
| TokenName | NVARCHAR(100) PK | Unique name: `users_delta`, `mailboxes_delta_timestamp`, `mfc_groups_delta_token`, `permissions_baseline_timestamp`, `permissions_delta_timestamp` |
| TokenValue | NVARCHAR(MAX) | Token URL or ISO timestamp |
| CreatedAt | DATETIME2 | First created |
| UpdatedAt | DATETIME2 | Last updated |
| IsActive | BIT | 0 = deactivated (e.g. expired) |
| DeactivatedAt | DATETIME2 NULL | When deactivated |
| DeactivationReason | NVARCHAR(500) NULL | Why deactivated |

**SyncLog** (`01_shared_tables.sql` + `04_permissions_tables.sql`)

One row per script execution for operational monitoring.

| Column | Type | Description |
|--------|------|-------------|
| RunId | UNIQUEIDENTIFIER PK | GUID generated at script start |
| FunctionName | NVARCHAR(100) | Script name |
| StartedAt / CompletedAt | DATETIME2 | Execution window |
| Status | NVARCHAR(20) | Running / Success / PartialFailure / Failed |
| UsersInserted / UsersUpdated / UsersSoftDeleted / UsersProcessed | INT | User script counters |
| MailboxesInserted / MailboxesUpdated / MailboxesSoftDeleted / MailboxesProcessed | INT | Mailbox script counters |
| MfcGroupsProcessed / MfcMembersProcessed | INT | MFC script counters |
| FullAccessProcessed / SendAsProcessed / SendOnBehalfProcessed / FolderPermissionsProcessed | INT | Permission script counters |
| MailboxesSkipped | INT | Mailboxes that errored or had no data |
| ErrorCount | INT | Total errors during the run |
| ErrorMessage | NVARCHAR(MAX) | Summary error text |
| TokenAdvancedTo | NVARCHAR(MAX) | Delta token value saved at end of run |

**Users** (`02_users_table.sql`)

Entra ID user directory. Primary key is the Entra Object ID.

| Column | Type | Description |
|--------|------|-------------|
| UserId | UNIQUEIDENTIFIER PK | Entra Object ID |
| UserPrincipalName | NVARCHAR(320) | UPN |
| DisplayName | NVARCHAR(256) | Display name |
| Mail | NVARCHAR(320) | Primary email |
| AccountEnabled | BIT NULL | Account status |
| UserType | NVARCHAR(50) | Member / Guest |
| OnPremisesSyncEnabled | BIT NULL | Hybrid identity flag |
| Department / JobTitle | NVARCHAR(100) | Org metadata |
| EntraCreatedDateTime | DATETIME2 NULL | Account creation date |
| IsDeleted / DeletedAt | BIT / DATETIME2 | Soft-delete pattern |
| SyncSource | NVARCHAR(20) | Baseline or Delta |
| LastSyncRunId | UNIQUEIDENTIFIER | FK to SyncLog |
| LastSyncedAt / LastModifiedAt | DATETIME2 | Audit timestamps |

**Mailboxes** (`03_mailboxes_table.sql` + `05_add_permissions_baselined.sql`)

Exchange Online mailbox inventory. Primary key is the ExchangeGuid.

| Column | Type | Description |
|--------|------|-------------|
| ExchangeGuid | UNIQUEIDENTIFIER PK | Mailbox GUID (stable across moves/renames) |
| UserId | UNIQUEIDENTIFIER NULL FK → Users | Entra Object ID (nullable for FK race condition handling) |
| PrimarySmtpAddress / UserPrincipalName / DisplayName / Alias | NVARCHAR | Mailbox identity attributes |
| RecipientTypeDetails | NVARCHAR(50) | EXO type (UserMailbox, SharedMailbox, etc.) |
| MailboxType | NVARCHAR(20) | Normalised label (User, Shared, Room, Equipment) |
| HiddenFromAddressLists / LitigationHoldEnabled / IsDirSynced | BIT NULL | Policy flags |
| ArchiveStatus | NVARCHAR(50) | Archive state |
| ForwardingAddress / ForwardingSmtpAddress | NVARCHAR(320) | Forwarding config |
| GrantSendOnBehalfTo | NVARCHAR(MAX) | Semicolon-joined list of SoB trustees (raw) |
| WhenMailboxCreated / WhenChangedUTC | DATETIME2 | EXO timestamps |
| PermissionsBaselinedAt | DATETIME2 NULL | When permissions were first collected for this mailbox. NULL = not yet collected. |
| IsDeleted / DeletedAt | BIT / DATETIME2 | Soft-delete pattern |
| SyncSource / LastSyncRunId / LastSyncedAt / LastModifiedAt | Various | Audit columns |

**MfcGroups** (`04_permissions_tables.sql`)

One row per mfc-SHXXXXX distribution group.

| Column | Type | Description |
|--------|------|-------------|
| GroupObjectId | UNIQUEIDENTIFIER PK | Entra Object ID of the group |
| DisplayName | NVARCHAR(256) | e.g. `mfc-SH48329` |
| Mail | NVARCHAR(320) | Group email |
| ExchangeGuid | UNIQUEIDENTIFIER NULL FK → Mailboxes | Link to the shared mailbox |
| SharedMailboxAlias | NVARCHAR(100) | Derived: `SH48329` |
| IsDeleted / DeletedAt | BIT / DATETIME2 | Soft-delete |
| LastSyncedAt / SyncSource / LastSyncRunId | Various | Audit |

**MfcGroupMembers** (`04_permissions_tables.sql`)

Current membership of each MFC group. Full-replace pattern (DELETE + INSERT per group), NO soft-delete.

| Column | Type | Description |
|--------|------|-------------|
| MfcGroupMemberId | INT IDENTITY PK | Surrogate key |
| GroupObjectId | UNIQUEIDENTIFIER FK → MfcGroups | Which group |
| PermissionType | NVARCHAR(20) | `FullAccess` or `SendOnBehalf` (one row per permission type per member) |
| TrusteeRawIdentity | NVARCHAR(320) | PrimarySmtpAddress of the member |
| TrusteeDisplayName | NVARCHAR(256) | Display name |
| TrusteeRecipientType | NVARCHAR(100) | RecipientTypeDetails |
| ResolvedUserId | UNIQUEIDENTIFIER NULL FK → Users | Populated by Phase 4 |
| FirstSeenAt / LastSyncedAt / LastSyncRunId | Various | Audit |

Unique index: `(GroupObjectId, PermissionType, TrusteeRawIdentity)` — not filtered (no soft-delete).

**FullAccessPermissions** (`04_permissions_tables.sql`)

Direct Full Access grants on user mailboxes. Shared mailbox Full Access is tracked via MfcGroupMembers.

| Column | Type | Description |
|--------|------|-------------|
| FullAccessPermissionId | INT IDENTITY PK | Surrogate key |
| ExchangeGuid | UNIQUEIDENTIFIER FK → Mailboxes | Target mailbox |
| TrusteeRawIdentity | NVARCHAR(320) | Trustee as returned by EXO (UPN, SMTP, DOMAIN\user, SID) |
| ResolvedUserId | UNIQUEIDENTIFIER NULL FK → Users | Populated by Phase 4 |
| IsInherited | BIT | Whether inherited from parent |
| AutoMapping | BIT NULL | Outlook auto-mapping flag |
| IsDeleted / DeletedAt | BIT / DATETIME2 | Soft-delete |
| SyncSource / LastSyncRunId / LastSyncedAt | Various | Audit |

Unique filtered index: `(ExchangeGuid, TrusteeRawIdentity) WHERE IsDeleted = 0`

**SendAsPermissions** (`04_permissions_tables.sql`)

Send-As grants on all mailbox types. Always direct (never group-managed).

| Column | Type | Description |
|--------|------|-------------|
| SendAsPermissionId | INT IDENTITY PK | Surrogate key |
| ExchangeGuid | UNIQUEIDENTIFIER FK → Mailboxes | Target mailbox |
| TrusteeRawIdentity | NVARCHAR(320) | Trustee identity |
| ResolvedUserId | UNIQUEIDENTIFIER NULL FK → Users | Populated by Phase 4 |
| AccessControlType | NVARCHAR(10) | Allow or Deny |
| IsDeleted / DeletedAt | BIT / DATETIME2 | Soft-delete |
| SyncSource / LastSyncRunId / LastSyncedAt | Various | Audit |

Unique filtered index: `(ExchangeGuid, TrusteeRawIdentity) WHERE IsDeleted = 0`

**SendOnBehalfPermissions** (`04_permissions_tables.sql`)

Direct Send-on-Behalf grants on user mailboxes. Shared mailbox SoB is tracked via MfcGroupMembers.

| Column | Type | Description |
|--------|------|-------------|
| SendOnBehalfPermissionId | INT IDENTITY PK | Surrogate key |
| ExchangeGuid | UNIQUEIDENTIFIER FK → Mailboxes | Target mailbox |
| TrusteeRawIdentity | NVARCHAR(320) | Trustee identity (from GrantSendOnBehalfTo) |
| ResolvedUserId | UNIQUEIDENTIFIER NULL FK → Users | Populated by Phase 4 |
| IsDeleted / DeletedAt | BIT / DATETIME2 | Soft-delete |
| SyncSource / LastSyncRunId / LastSyncedAt | Various | Audit |

Unique filtered index: `(ExchangeGuid, TrusteeRawIdentity) WHERE IsDeleted = 0`

**FolderPermissions** (`04_permissions_tables.sql`)

Folder-level permissions on all mailbox types.

| Column | Type | Description |
|--------|------|-------------|
| FolderPermissionId | INT IDENTITY PK | Surrogate key |
| ExchangeGuid | UNIQUEIDENTIFIER FK → Mailboxes | Target mailbox |
| FolderPath | NVARCHAR(500) | e.g. `\Calendar`, `\Inbox\Reports` |
| FolderName | NVARCHAR(256) | Leaf folder name |
| TrusteeRawIdentity | NVARCHAR(320) | Trustee (user, Default, Anonymous) |
| ResolvedUserId | UNIQUEIDENTIFIER NULL FK → Users | Populated by Phase 4 |
| AccessRights | NVARCHAR(500) | e.g. `Reviewer`, `FullAccess`, `ReadItems,CreateItems` |
| SharingPermissionFlags | NVARCHAR(200) NULL | Sharing-specific flags |
| IsDefaultTrustee | BIT | True for Default/Anonymous entries |
| IsDeleted / DeletedAt | BIT / DATETIME2 | Soft-delete |
| SyncSource / LastSyncRunId / LastSyncedAt | Various | Audit |

Unique filtered index: `(ExchangeGuid, FolderPath, TrusteeRawIdentity) WHERE IsDeleted = 0`
Calendar-specific index: `(ExchangeGuid, TrusteeRawIdentity) WHERE FolderName = 'Calendar' AND IsDeleted = 0`

### 4.2 Views

**vw_AllMailboxPermissions** — union of all six permission sources (FullAccess direct + MFC, SendAs, SendOnBehalf direct + MFC, FolderPermissions) joined to Mailboxes for context. Filters `p.IsDeleted = 0` on each permission table and `mg.IsDeleted = 0` on MFC groups. This is the primary reporting view for Power BI.

**vw_MfcGroupMembership** — operational view joining MfcGroupMembers → MfcGroups → Mailboxes → Users for a complete picture of MFC group membership with resolved user details.

### 4.3 SQL Migrations (Run Order)

1. `01_shared_tables.sql` — DeltaTokens, SyncLog
2. `02_users_table.sql` — Users
3. `03_mailboxes_table.sql` — Mailboxes
4. `04_permissions_tables.sql` — All permission tables, views, SyncLog extensions
5. `05_add_permissions_baselined.sql` — Adds PermissionsBaselinedAt to Mailboxes
6. `06_column_size_fixes.sql` — ALTER COLUMN fixes for TokenValue, TokenAdvancedTo, ErrorMessage to NVARCHAR(MAX) (**not yet created**)

---

## 5. Shared Modules

All modules are in the `shared/` folder. Scripts import them via `Import-Module -Force`.

### 5.1 ConfigHelpers.psm1

Reads `config.xml` and provides certificate-based Azure authentication.

| Function | Description |
|----------|-------------|
| Import-SyncConfig | Parses config.xml, returns a structured config object |
| Connect-SyncServicePrincipal | Connects to Azure using Az.Accounts with certificate auth. Sets Az context for token acquisition. |

### 5.2 LoggingHelpers.psm1

File-based logging. One log file per script run: `ProcessName_yyyy-MM-dd_HH-mm-ss.log`.

| Function | Description |
|----------|-------------|
| Initialize-Logging | Creates log file, writes header |
| Write-LogInfo / Write-LogWarning / Write-LogError | Writes timestamped entries at appropriate severity. Write-LogError accepts -ErrorRecord for stack trace. |
| Write-LogSection | Writes a section divider with title |
| Write-LogSummary | Writes a key-value summary block (used at end of each script) |
| Close-Logging | Writes footer with final status and duration |
| Remove-OldLogFiles | Deletes logs older than RetentionDays (from config) |

### 5.3 DependencyHelpers.psm1

Validates and imports required PowerShell modules from a local `\Modules` folder.

| Function | Description |
|----------|-------------|
| Initialize-ModuleDependencies | Reads `<Dependencies>` from config. For each required module, locates it in the local Modules folder, validates version ≥ minimum, imports it. Fatal exit if any module is missing. |

Required modules (from config.xml):
- `Az.Accounts` ≥ 3.0.0
- `ExchangeOnlineManagement` ≥ 3.4.0

These must be pre-installed in a `\Modules` folder one level above the scripts root.

### 5.4 GraphHelpers.psm1

Microsoft Graph REST API infrastructure. Token caching, retry, pagination, delta queries, and hunting queries.

| Function | Description |
|----------|-------------|
| Initialize-GraphContext | Stores Graph config settings. Call once per script. |
| Get-GraphToken | Acquires/caches a Graph access token via Az.Accounts. Auto-refreshes when within 2 minutes of expiry. Accepts -ForceRefresh. |
| Invoke-GraphRequest | Single GET request with retry on 429/401/5xx. Handles Retry-After header. ConsistencyLevel:eventual for advanced queries. |
| Invoke-GraphPagedRequest | Follows @odata.nextLink to retrieve all pages. Returns the combined results array. |
| Invoke-GraphDeltaQuery | Pages through a delta URL, returns `{ Objects, DeltaToken, TokenExpired }`. Detects HTTP 410 (token expired). |
| Get-GraphMailFolderPaths | Recursively enumerates all visible folder paths for a mailbox via GET /users/{id}/mailFolders. Returns string array (`\Calendar`, `\Inbox\Reports`). |
| Invoke-GraphHuntingQuery | POST to /security/runHuntingQuery with KQL. Retry on 429/401/5xx. Returns results array. Warns if 10K row cap hit. |

### 5.5 SqlHelpers.psm1

Azure SQL Database operations. Certificate-based token auth via Az.Accounts.

| Function | Description |
|----------|-------------|
| Initialize-SqlContext | Stores SQL connection settings from config. Call once per script. |
| Open-SqlConnection | Returns an open, authenticated SqlConnection. **Caller must call .Dispose()** when done. Used to hold a connection open across bulk operations. |
| Invoke-SqlNonQuery | Executes INSERT/UPDATE/DELETE/MERGE. Returns rows affected. Accepts optional -Connection and -Transaction for shared-connection patterns. If no -Connection provided, opens and closes its own. Error messages include first 150 chars of the failing query for troubleshooting. |
| Invoke-SqlScalar | Executes a query, returns the first column of the first row. Accepts optional -Connection. |
| Invoke-SqlQuery | Executes a SELECT, returns array of PSCustomObjects. Accepts optional -Connection. |
| Get-DeltaToken | Returns the active TokenValue for a named token, or $null. |
| Save-DeltaToken | Upserts a named delta token (MERGE). |
| Disable-DeltaToken | Marks a token inactive with a reason (e.g. "expired"). |
| ConvertFrom-JwtToken | Decodes a JWT payload for logging identity. |

**SQL connection patterns used in scripts:**

Single operation (connection auto-managed):
```powershell
Invoke-SqlNonQuery -Query $sql -Parameters @{ '@Id' = $id }
```

Bulk loop (one connection held open):
```powershell
$conn = Open-SqlConnection
try {
    foreach ($item in $items) {
        Invoke-SqlNonQuery -Connection $conn -Query $sql -Parameters @{...}
    }
}
finally { $conn.Dispose() }
```

Transactional batch (atomic delete + insert):
```powershell
$conn = Open-SqlConnection
try {
    $tx = $conn.BeginTransaction()
    Invoke-SqlNonQuery -Connection $conn -Transaction $tx -Query $deleteSql ...
    Invoke-SqlNonQuery -Connection $conn -Transaction $tx -Query $insertSql ...
    $tx.Commit()
}
catch { try { $tx.Rollback() } catch {}; throw }
finally { $conn.Dispose() }
```

### 5.6 ExoHelpers.psm1

Exchange Online PowerShell connection management only.

| Function | Description |
|----------|-------------|
| Initialize-ExoContext | Validates EXO config, checks ExchangeOnlineManagement module |
| Connect-ExoSession | Certificate-based Connect-ExchangeOnline. Safe to call multiple times (no-op if connected). |
| Disconnect-ExoSession | Cleanly disconnects. Call at end of every script that uses EXO. |

### 5.7 AdminApiHelpers.psm1

EXO Admin API v2.0 infrastructure. POST-based requests with retry and pagination.

| Function | Description |
|----------|-------------|
| Initialize-AdminApiContext | Stores Admin API config. Call once per script. |
| Get-AdminApiToken | Acquires/caches token for resource `https://outlook.office365.com/`. |
| Get-AnchorMailboxHeader | Returns the X-AnchorMailbox header value. Mode=Mailbox → `UPN:<upn>`. Mode=AppOnly → `APP:SystemMailbox{...}@<org>`. |
| Invoke-AdminApiRequest | Single POST to an Admin API endpoint with CmdletInput envelope. Retry on 429/401/5xx. Returns the response.value array. |
| Invoke-AdminApiPagedRequest | POST + follow @odata.nextLink (re-POST with original body to continuation URL). Returns all pages combined. |

**Admin API call pattern from scripts:**
```powershell
$anchor  = Get-AnchorMailboxHeader -Mode AppOnly
$members = Invoke-AdminApiPagedRequest `
    -Endpoint      'DistributionGroupMember' `
    -CmdletName    'Get-DistributionGroupMember' `
    -Parameters    @{ Identity = $groupName; ResultSize = 'Unlimited' } `
    -AnchorMailbox $anchor `
    -Select        'PrimarySmtpAddress,DisplayName,RecipientTypeDetails'
```

---

## 6. Scripts

### 6.1 Invoke-UserBaselineLoad.ps1

**Purpose:** Full snapshot of all Entra ID users into the Users table. Run on initial deployment and when the delta token expires.

**Duration:** 10–30 minutes for 40K users.

**Phases:**

| Phase | Action |
|-------|--------|
| 1 — Graph user fetch | Pages through `/users?$select=...&$top=999` via Invoke-GraphPagedRequest |
| 2 — SQL upsert | One Open-SqlConnection held for the loop. Per-user MERGE into Users with per-row error handling. Progress logged every 5,000 users. |
| 3 — Capture delta token | Calls `/users/delta?$select=id&$top=999` to consume all pages and capture the deltaLink. Saves as `users_delta` in DeltaTokens. |

**Crash safety:** Delta token saved only after all MERGEs complete. A crash during Phase 2 means re-run processes the same users (idempotent MERGEs).

### 6.2 Invoke-UserDeltaSync.ps1

**Purpose:** 15-minute incremental sync using Graph `/users/delta` token.

**Phases:**

| Phase | Action |
|-------|--------|
| 1 — Fetch delta token | Reads `users_delta` from DeltaTokens. Exits if not found. |
| 2 — Graph delta fetch | Calls `/users/delta?$select=...&$deltatoken=...`. Handles HTTP 410 (token expired → deactivates token, exits). |
| 3 — Apply changes | Processes `@removed` objects first (soft-delete). Then MERGE for new/changed users. Uses COALESCE on UPDATE — Graph returns only changed properties, not the full object. |
| 4 — Advance delta token | Saves the new deltaLink. Only saved after all Phase 3 processing. |

### 6.3 Invoke-MailboxBaselineLoad.ps1

**Purpose:** Full snapshot of all Exchange Online mailboxes. Run weekly for reconciliation.

**Duration:** 15–45 minutes for 40K mailboxes.

**Phases:**

| Phase | Action |
|-------|--------|
| 1 — EXO mailbox fetch | `Get-EXOMailbox -ResultSize Unlimited -Properties $mailboxProps` filtered to configured MailboxTypes. |
| 2 — SQL upsert | One Open-SqlConnection for the loop. Per-mailbox MERGE with inline normalisation (RecipientTypeDetails → MailboxType, GrantSendOnBehalfTo flattened to semicolons). |
| 3 — Deletion reconciliation | Compares DB active GUIDs against EXO returned GUIDs. Soft-deletes any DB row no longer in EXO. Catches permanent deletions the delta misses. |
| 4 — Save delta timestamp | Saves current UTC as `mailboxes_delta_timestamp`. |

### 6.4 Invoke-MailboxDeltaSync.ps1

**Purpose:** 15-minute incremental sync using WhenChangedUTC filter.

**Phases:**

| Phase | Action |
|-------|--------|
| 1 — Determine sync window | Reads `mailboxes_delta_timestamp`, subtracts DeltaOverlapMinutes (30). |
| 2 — EXO delta fetch | Queries both active and soft-deleted mailboxes with `WhenChangedUTC -ge $filterValue`. |
| 3 — Apply changes | Soft-deletes for EXO soft-deleted mailboxes. MERGE with COALESCE for active mailboxes. |
| 4 — Advance timestamp | Saves current UTC. |

### 6.5 Invoke-PermissionBaselineLoad.ps1

**Purpose:** Full collection of all permission types. Run monthly or on-demand.

**Duration:** 18–34 hours for 40K mailboxes.

**Phases:**

| Phase | Action | Connection Pattern |
|-------|--------|-------------------|
| 1 — Load mailboxes from DB | SELECT active mailboxes. Builds aliasToGuid lookup for shared mailbox Alias → ExchangeGuid resolution. | Single query |
| 2 — MFC Groups | Graph discovery (`startsWith(displayName,'mfc-SH')`) + Admin API `Get-DistributionGroupMember` per group. DELETE + INSERT in transaction per group. Saves initial `mfc_groups_delta_token`. | Per-group connection + transaction |
| 3 — Full Access + Send-As | Soft-delete all existing rows at phase start. Then per-mailbox: `Get-EXOMailboxPermission` (user only) + `Get-EXORecipientPermission` (all). One connection held for entire loop. | Single connection for phase |
| 4 — Send-on-Behalf | Soft-delete all, then promote from `Mailboxes.GrantSendOnBehalfTo`. One connection for loop. | Single connection for phase |
| 5 — Folder Permissions | Per-mailbox: Graph folder enumeration → Admin API `Get-MailboxFolderPermission` per folder → soft-delete + INSERT in transaction. One connection, per-mailbox transaction. | Single connection, per-mailbox transaction |
| 6 — Save timestamp | Saves `permissions_baseline_timestamp`. Sets `PermissionsBaselinedAt = SYSUTCDATETIME()` on all active mailboxes. | Single query |

### 6.6 Invoke-MfcGroupDeltaSync.ps1

**Purpose:** 15-minute incremental sync for MFC group membership using Graph `/groups/delta`.

**Phases:**

| Phase | Action |
|-------|--------|
| 1 — Fetch delta token | Reads `mfc_groups_delta_token`. |
| 2 — Load reference data | Known MFC GroupObjectIds from DB + shared mailbox Alias→ExchangeGuid lookup. |
| 3 — Graph groups delta | Calls `/groups/delta?$select=id,displayName,mail&$deltatoken=...`. Handles 410. |
| 4 — Apply changes | For each delta object: (a) Skip non-MFC groups. (b) @removed → soft-delete group + delete members (transaction). (c) New/changed → MERGE group + re-fetch full membership from Admin API + replace members (transaction). |
| 5 — Advance delta token | Saved only after all processing. |

**Classification logic:** A group is MFC if its GroupObjectId is in the known set (existing MFC group) OR its displayName starts with `mfc-SH` (new MFC group).

### 6.7 Invoke-PermissionDeltaSync.ps1

**Purpose:** 15-minute incremental sync for all non-MFC permission types, using Microsoft Defender CloudAppEvents for change detection and new-mailbox mini-baselining.

**Phases:**

| Phase | Action |
|-------|--------|
| 0 — New mailbox mini-baseline | Queries `Mailboxes WHERE PermissionsBaselinedAt IS NULL`. For each: collects Full Access + Send-As (EXO), Send-on-Behalf (DB), Folder Permissions (Graph + Admin API). Marks PermissionsBaselinedAt on completion. |
| 1 — Load delta timestamp | Reads `permissions_delta_timestamp`. If not found (first run): saves current time, exits. Next run will have a valid window. |
| 2 — Load mailbox lookup | Builds SMTP/UPN → `{ExchangeGuid, MailboxType}` hashtable from Mailboxes table. |
| 3 — CloudAppEvents query | KQL query against CloudAppEvents for permission-change operations in the time window (with overlap). Via Invoke-GraphHuntingQuery. |
| 4 — Full Access + Send-As | Extracts unique affected mailboxes from FA/SA events. Per-mailbox: re-fetches current state from EXO, replaces in DB (soft-delete + insert). One connection for phase. |
| 5 — Send-on-Behalf | Filters Set-Mailbox events with GrantSendOnBehalfTo parameter. Per-mailbox: replaces all SoB from the event payload (soft-delete + insert). |
| 6 — Folder Permissions | Processes Add/Set/Remove-MailboxFolderPermission events in timestamp order. Surgical INSERT or soft-delete per event. |
| 7 — Advance timestamp | Saves current UTC only after all phases complete. |

**KQL query:**
```kql
CloudAppEvents
| where Timestamp > datetime(<stored - overlap>)
| where Application == 'Microsoft Exchange Online'
| where ActionType in (
    'Add-MailboxPermission', 'Remove-MailboxPermission', 'Set-MailboxPermission',
    'Add-RecipientPermission', 'Remove-RecipientPermission',
    'Set-Mailbox',
    'Add-MailboxFolderPermission', 'Set-MailboxFolderPermission', 'Remove-MailboxFolderPermission'
)
| project Timestamp, ActionType, ObjectName, RawEventData
| take 100000
| order by Timestamp asc
```

**Private helper functions:**

- `Get-AuditParameter` — extracts a named parameter from `RawEventData.Parameters`
- `Resolve-MailboxGuid` — resolves SMTP/UPN to ExchangeGuid via lookup hashtable
- `Replace-MailboxPermissions` — shared between Phase 0 and Phase 4: re-fetches Full Access + Send-As from EXO and replaces in DB

---

## 7. Cross-Cutting Design Patterns

### 7.1 Soft-Delete

All primary tables (except MfcGroupMembers) use soft-delete: `IsDeleted BIT DEFAULT 0` + `DeletedAt DATETIME2 NULL`. Active rows have `IsDeleted = 0`. Deletion sets `IsDeleted = 1, DeletedAt = SYSUTCDATETIME()`. This preserves audit history.

Exception: MfcGroupMembers uses full-replace (DELETE + INSERT per group per sync). No historical record of past membership.

### 7.2 Crash-Safe Token Advancement

Every delta sync script follows the same pattern: the delta token is saved ONLY after ALL processing completes successfully. If a script crashes mid-processing, the next run re-reads the old token and reprocesses the same window. This means events may be processed twice — all SQL operations are idempotent (MERGE for upserts, filtered unique indexes prevent duplicates).

### 7.3 Overlap Windows

Delta syncs query slightly further back than the stored timestamp (configurable via `Sync.DeltaOverlapMinutes`, default 30). This ensures events/changes that appeared just before the boundary are caught even with ingestion delays.

### 7.4 SyncLog Tracking

Every script inserts a SyncLog row at startup (`Status = 'Running'`) and updates it at completion with final status, counters, and the advanced token value. This provides operational visibility without external monitoring infrastructure.

### 7.5 Run ID Correlation

Every script generates a `RunId` (GUID) at startup. This ID is written to every SQL row touched during that run (`LastSyncRunId` column). This enables tracing any row back to the exact script execution that created or modified it.

---

## 8. Configuration (config.xml)

| Section | Key Settings |
|---------|-------------|
| General | Environment name, Organisation domain |
| Authentication | TenantId, AppId, CertificateThumbprint, CertificateStoreLocation. Also ClientSecret (for non-cert auth modes — should be moved to Key Vault). |
| Database | Server, Name, TargetTenantId (for cross-tenant SQL auth), ConnectionTimeoutSec, CommandTimeoutSec |
| Logging | Directory, RetentionDays, Level (Info/Warning/Error), EventLog integration |
| Graph | BaseUrl, TokenScope, PageSize, MaxRetries, ThrottleBackoffMaxSec |
| ExchangeOnline | Organisation, MailboxTypes (comma-separated), SqlBatchSize |
| AdminApi | BaseUrl, MaxRetries, ThrottleBackoffMaxSec |
| Sync | DeltaOverlapMinutes |
| Dependencies | Module definitions with Name and MinimumVersion |

---

## 9. Deployment Guide

### 9.1 Prerequisites

1. Azure SQL Database provisioned (zone-redundant, Serverless GP tier recommended)
2. App registration with all permissions from Section 2.4 granted (admin consent)
3. Certificate installed in the local machine cert store
4. PowerShell 7+ installed on the server
5. Local `\Modules` folder with Az.Accounts (≥ 3.0.0) and ExchangeOnlineManagement (≥ 3.4.0)
6. config.xml populated with all settings (no REPLACE-WITH-* placeholders)

### 9.2 First-Run Order

```
1. Run SQL migrations in order: 01 → 02 → 03 → 04 → 05 → 06
2. Run Invoke-UserBaselineLoad.ps1
3. Run Invoke-MailboxBaselineLoad.ps1
4. Run Invoke-PermissionBaselineLoad.ps1
5. Enable all four 15-minute delta sync scheduled tasks
6. Schedule Invoke-MailboxBaselineLoad weekly
7. Schedule Invoke-PermissionBaselineLoad monthly
```

Dependencies: 2 must complete before 3 (Mailboxes.UserId FK → Users). 3 must complete before 4 (PermissionBaselineLoad reads mailbox data from DB).

### 9.3 Scheduled Task Configuration

| Script | Trigger | Execution Time Limit | Multiple Instances |
|--------|---------|---------------------|--------------------|
| Invoke-UserDeltaSync | Daily, repeat every 15 min | 30 min | Do not start new |
| Invoke-MailboxDeltaSync | Daily, repeat every 15 min | 30 min | Do not start new |
| Invoke-MfcGroupDeltaSync | Daily, repeat every 15 min | 30 min | Do not start new |
| Invoke-PermissionDeltaSync | Daily, repeat every 15 min | 30 min | Do not start new |
| Invoke-MailboxBaselineLoad | Weekly Sunday 03:00 | 4 hours | Do not start new |
| Invoke-PermissionBaselineLoad | 1st of month 00:00 | 48 hours | Do not start new |
| Invoke-UserBaselineLoad | 1st of month 02:00 | 2 hours | Do not start new |

Program: `pwsh.exe`
Arguments: `-NonInteractive -File "C:\M365PermSync\<ScriptName>.ps1"`

---

## 10. Known Issues and Gap Review

### 10.1 First Review Findings

| # | Severity | Finding | Status |
|---|----------|---------|--------|
| 1 | Critical | Client secret in plaintext in config.xml (appears twice) | Pending — move to Key Vault |
| 2 | Critical | Phase 4 trustee resolution not built (ResolvedUserId NULL) | Pending — build Invoke-TrusteeResolution.ps1 |
| 3 | Important | Mailboxes.UserId FK race condition — new mailbox for user not yet in Users table | Pending — insert UserId=NULL when user absent |
| 4 | Important | vw_AllMailboxPermissions doesn't filter m.IsDeleted=0 — returns permissions for deleted mailboxes | Pending — add filter |
| 5 | Important | DeltaTokens.TokenValue and SyncLog columns may be too short — ALTER to NVARCHAR(MAX) | Pending — create 06_column_size_fixes.sql |
| 6 | Operational | DependencyHelpers requires undocumented local \Modules folder | Documented in this guide |
| 7 | Operational | Add-SP.sql contains bare DELETE statements — dangerous | Pending — rename/remove |
| 8 | Operational | No deployment guide | Addressed in this guide Section 9 |
| 9 | Operational | App permissions never consolidated | Addressed in this guide Section 2.4 |
| 10 | Operational | SyncLog column mismatch — PermissionDeltaSync stores NewMailboxes in MfcMembersProcessed | Pending — add NewMailboxesBaselined column |
| 11 | Minor | Soft-delete on mailbox doesn't cascade to permission tables | Future |
| 12 | Minor | EXO connects on every PermissionDeltaSync run even when unneeded | Future |
| 13 | Minor | No sanity check for CloudAppEvents returning 0 results consistently | Future |

### 10.2 Second Review Findings

| # | Severity | Finding | Status |
|---|----------|---------|--------|
| A | Critical | permissions_delta_timestamp not seeded by baseline — gap between baseline completion and first delta run | Pending fix |
| B | Critical | Phase 6 Add-MailboxFolderPermission bare INSERT violates unique index on overlap reprocessing | Pending fix |
| C | Important | EXO session timeout during 12-24h baseline — no reconnect logic | Pending fix |
| D | Important | PermissionsBaselinedAt set only at end of baseline — failed baseline triggers Phase 0 stampede | Pending fix |
| E | Important | AccessRights format mismatch between baseline (effective rights) and delta (cmdlet parameter) | Pending |
| F | Important | RawEventData from hunting API may be JSON string not object — needs defensive parsing | Pending fix |
| G | Confirmed | vw_AllMailboxPermissions missing m.IsDeleted=0 on Mailboxes joins (both direct and MFC) | Same as #4 |
| H | Compliance | MfcGroupMembers full-replace = no historical membership record | Future |
| I | Edge case | Mailbox type change (Shared↔User) leaves stale permissions from old collection strategy | Future |
| J | Edge case | Folder identity parse breaks if folder name contains literal ':\' | Pending guard |
| K | Operational | Auth mode inconsistency in config (Secret vs Certificate) | Pending — standardise on cert |
| L | Future | No HA run-lock at DB level | Future |
| M | Maintenance | Duplicated folder-collection logic between baseline Phase 5 and delta Phase 0 | Future |
| N | Governance | Log files contain sensitive permission data (access map) — needs ACLs | Future |

---

## 11. Deliverables Inventory

### SQL Migrations
- `01_shared_tables.sql`
- `02_users_table.sql`
- `03_mailboxes_table.sql`
- `04_permissions_tables.sql`
- `05_add_permissions_baselined.sql`

### Shared Modules
- `ConfigHelpers.psm1`
- `LoggingHelpers.psm1`
- `DependencyHelpers.psm1`
- `GraphHelpers.psm1`
- `SqlHelpers.psm1`
- `ExoHelpers.psm1`
- `AdminApiHelpers.psm1`

### Scripts
- `Invoke-UserBaselineLoad.ps1`
- `Invoke-UserDeltaSync.ps1`
- `Invoke-MailboxBaselineLoad.ps1`
- `Invoke-MailboxDeltaSync.ps1`
- `Invoke-PermissionBaselineLoad.ps1`
- `Invoke-MfcGroupDeltaSync.ps1`
- `Invoke-PermissionDeltaSync.ps1`

### Configuration
- `config.xml`

### Documentation
- `M365-PermSync-Solution-Guide.md` (this document)

### Not Yet Built
- `06_column_size_fixes.sql`
- `Invoke-TrusteeResolution.ps1` (Phase 4)
