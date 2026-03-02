# User SOA Switch (Entra ID)

PowerShell 7 + WPF tool for **switching user Source of Authority (SOA)** from on-premises to cloud in Microsoft Entra ID, with built-in **backup** and **on-prem synced attribute clearing**.

This project is designed for controlled user lifecycle operations where synced users are being moved to cloud-managed identity.

## What this tool is for

Primary use case:
- Select one or more synced users
- Back up their synced/on-premises attribute state to JSON
- Switch SOA to cloud (`isCloudManaged = true`)
- Clear synced on-premises attributes on the Entra user object

Also included:
- Restore attributes from a backup JSON file
- Roll back SOA to on-prem (`isCloudManaged = false`)
- Load either synced users or cloud-managed users for targeted operations

## Core capabilities

- Connect/disconnect to Microsoft Graph interactively
- Load all synced users (`onPremisesSyncEnabled eq true`)
- Load all cloud users (`onPremisesSyncEnabled ne true`)
- Filter by display name (3+ characters)
- Multi-select users with checkbox DataGrid
- Backup selected users in two formats:
  - single combined JSON
  - per-user JSON files + `_BackupIndex.json`
- SOA operations:
  - **Switch SOA to Cloud** (2-phase process)
    1. PATCH `onPremisesSyncBehavior` with `isCloudManaged=true`
    2. Clear on-premises sync attributes via `ADSyncTools`
  - **Clear On-Prem Attributes** (standalone operation)
  - **Rollback SOA to On-Prem** (`isCloudManaged=false`)
  - **Restore from Backup** (using backup JSON)
- Progress/status UI and timestamped debug logs

## Prerequisites

### Platform
- Windows 10/11 or Windows Server
- PowerShell 7.0+

### Modules
The script installs/loads what it needs:
- `Microsoft.Graph.Authentication` (v2.35.0)
- `Microsoft.Graph.Users` (v2.35.0)
- `ADSyncTools` (required for clear/restore operations)

### Graph scopes / permissions
The app requests:
- `User.Read.All`
- `User-OnPremisesSyncBehavior.ReadWrite.All`

Use an account/role that can perform these user operations in your tenant.

## Run

From PowerShell 7 in the repo folder:

```powershell
.\User-SOA-Switch.ps1
```

Optional launcher (starts a fresh PowerShell process):

```powershell
.\Launch.ps1
```

## Recommended workflow (Synced users → Cloud SOA)

1. Connect to Graph
2. Click **Load Synced Users**
3. Filter/select target users
4. Click **Backup Selected Users** and save JSON
5. Click **Switch SOA to Cloud**
   - This button runs both phases (SOA switch, then attribute clear)
6. Review completion summary and any per-user errors

## Notes on operations

- **Switch SOA to Cloud** is the primary action for migration and includes attribute clearing for successfully switched users.
- **Clear On-Prem Attributes** can be run separately when needed (for targeted cleanup).
- **Rollback SOA to On-Prem** changes SOA state only; it does not automatically restore all user data.
- **Restore from Backup** reads:
  - a single-user backup JSON
  - a combined multi-user backup JSON
  - a `_BackupIndex.json` folder index

## Backup content

Backups include user identity data and on-prem sync attributes used by this tool, such as:
- `OnPremisesDistinguishedName`
- `OnPremisesDomainName`
- `OnPremisesSamAccountName`
- `OnPremisesUserPrincipalName`
- `OnPremisesSecurityIdentifier`
- `OnPremisesImmutableId`
- `OnPremisesLastSyncDateTime`
- `ProxyAddresses`

## Logging and troubleshooting

- Each run creates a timestamped `debug_log_*.txt` in the repo folder.
- Use these logs to troubleshoot Graph/API or permission issues.
- A basic forms/WPF check script is included at:
  - `Troubleshooting/Test-WindowsForms.ps1`

## Safety guidance

- Always back up before SOA changes.
- Users can remain in on-prem sync scope after SOA switch.
- After SOA switch, AD DS changes for those users do not flow to Entra ID.
- Validate in a pilot group first.
- Ensure change control/rollback planning is in place for production operations.
