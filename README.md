# Entra ID User Sync Attribute Backup Tool - Multi-User Edition

A PowerShell 7+ WPF application for backing up on-premises synced user attributes from Microsoft Entra ID using Microsoft Graph API with support for bulk operations, filtering, and flexible export formats.

## Overview

This tool provides a comprehensive interface to:
- Connect to Microsoft Graph with interactive authentication
- **Load ALL on-premises synced users** from your Entra ID tenant
- **Filter users by display name** (3+ characters) for quick searches
- **Select multiple users** via checkboxes (individual, filtered, or all)
- **Export to JSON** in single file or individual files per user
- **Clear on-premises attributes** for selected users (SOA prerequisite)
- **Switch Source of Authority (SOA)** from on-premises to cloud for selected users
- **Restore user attributes** from a previous JSON backup
- View real-time counts of total, filtered, and selected users

## Features

✅ **DataGrid Multi-User Interface** - Modern tabular view with sortable columns  
✅ **Bulk User Loading** - Load all on-premises synced users with pagination support  
✅ **Client-Side Filtering** - Filter by display name (minimum 3 characters) for fast searches  
✅ **Flexible Selection** - Individual checkboxes, Select All Visible, or Clear All  
✅ **Multiple Export Formats**:
   - **Single JSON file** - All users in one array
   - **Individual files** - One JSON per user in timestamped folder

✅ **SOA Management** - Full Source of Authority management workflow:
   - **Clear On-Prem Attributes** - Nulls on-premises sync attributes for selected users
   - **Switch SOA to Cloud** - Sets `onPremisesSyncEnabled` to `false` for selected users
   - **Restore from Backup** - Restores on-premises attributes from a previous JSON backup
   - **Warning banner** - Prominent prerequisite guidance before running SOA operations

✅ **Comprehensive Attribute Capture** - Backs up all on-premises sync attributes:
   - Distinguished Name, Domain Name, SAM Account Name
   - Security Identifier, Immutable ID
   - Last Sync DateTime, UPN (on-prem)
   - Proxy Addresses, Account Status
   - User Type, Mail, Object ID

✅ **Progress Indication** - Real-time progress bar with user count during load  
✅ **Selection Persistence** - Selections maintained when filtering  
✅ **Audit Trail** - Backup metadata includes timestamp, operator, and user count  
✅ **Debug Logging** - Comprehensive logging for troubleshooting
✅ **Error Handling** - Robust error handling with user-friendly messages

## Prerequisites

### Required Software
- **PowerShell 7.0 or higher** (PowerShell Core)
  - Download: https://github.com/PowerShell/PowerShell/releases
- **Windows 10/11** or **Windows Server 2016+**

### Required PowerShell Modules
The script will automatically install these if missing:
- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Users`

The following module must be installed manually (included with Azure AD Connect / Entra Cloud Sync):
- `ADSyncTools` — required for **Clear On-Prem Attributes** and **Restore from Backup** operations

### Required Permissions
You must have one of the following Entra ID roles or permissions:
- Global Administrator
- Global Reader
- User Administrator
- Or custom role with the following Graph API permissions:
  - `User.Read.All`
  - `User-OnPremisesSyncBehavior.ReadWrite.All`

> **Note:** `User-OnPremisesSyncBehavior.ReadWrite.All` is required for clearing on-premises attributes and switching SOA. Without this permission those operations will return Access Denied.

## Installation

1. **Clone or download this repository**
   ```powershell
   git clone https://github.com/yourusername/User-SOA-Switch.git
   cd User-SOA-Switch
   ```

2. **Verify PowerShell version**
   ```powershell
   $PSVersionTable.PSVersion
   # Should be 7.0 or higher
   ```

3. **Set execution policy** (if needed)
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## Usage

### Starting the Application

1. Open PowerShell 7
2. Navigate to the project directory
3. Run the script:
   ```powershell
   .\User-SOA-Switch.ps1
   ```

### Workflow

1. **Connect to Microsoft Graph**
   - Click the "Connect to Graph" button
   - Sign in with your credentials in the browser window
   - Consent to the required permissions (first run only)
   - Wait for "Connected" status
   - "Load Synced Users" button will become enabled

2. **Load All Synced Users**
   - Click "Load Synced Users" button
   - Progress bar displays: "Loading synced users: X of Y"
   - Wait for load to complete (time varies by tenant size)
   - DataGrid populates with all on-premises synced users
   - Counts display: Total Users, Filtered, Selected

3. **Filter Users** (Optional)
   - Type 3 or more characters in the "Filter by Display Name" box
   - DataGrid updates in real-time to show matching users
   - Filtered count updates automatically
   - Click "Clear Filter" to show all users again
   - **Note**: Selections persist even when users are filtered out

4. **Select Users for Backup**
   - **Individual Selection**: Click checkbox next to each user
   - **Select All Visible**: Click "Select All Visible" (only selects filtered users)
   - **Clear All**: Click "Clear All Selections" (clears all, including hidden)
   - Selected count updates in real-time
   - Backup button shows count: "Backup Selected Users (N)"

5. **Backup Selected Users**
   - Click "Backup Selected Users (N)" button
   - Choose backup format:
     - **Single JSON file**: All users in one array
     - **Individual files**: One JSON per user in a timestamped folder
   - Select save location
   - Wait for confirmation message
   - **Single file**: Choose filename and location
   - **Multiple files**: Choose folder (subfolder created automatically)

### SOA Management Workflow

The SOA section (highlighted in amber/yellow) provides the tools to move selected users from on-premises-managed (synced) to cloud-managed. **Always follow this sequence:**

1. **Backup** — Use "Backup Selected Users" to capture all on-premises attributes first.
2. **Clear On-Prem Attributes** — Click "Clear On-Prem Attributes" to null out the on-premises sync attributes for the selected users. This is a prerequisite before switching SOA.
3. **Switch SOA to Cloud** — Click "Switch SOA to Cloud" to set `onPremisesSyncEnabled` to `false` for the selected users, handing authority to Entra ID (cloud).
4. **Restore from Backup** *(if needed)* — If something goes wrong, use "Restore from Backup" to re-apply the saved on-premises attributes from a JSON backup file.

> ⚠ **Warning:** Ensure that each user has been **removed from the on-premises AD sync scope** (e.g., excluded from Azure AD Connect / Entra Cloud Sync) **before** switching SOA. Failure to do so may cause the sync engine to overwrite the changes on the next sync cycle.

6. **Refresh or Continue**
   - Click "Refresh" to reload data from Entra ID
   - Filter and select different users
   - Create additional backups as needed

7. **Close Application**
   - Click "Close" button
   - Application automatically disconnects from Graph

### Example User Scenarios

**Scenario 1: Bulk Backup Before Cloud Migration**
- Load all synced users (1000+ users)
- Select all with "Select All Visible"
- Choose "Individual files" format
- Preserve all on-premises attributes before breaking sync

**Scenario 2: Backup Specific Department**
- Load all synced users
- Filter by display name (e.g., "Sales")
- Select all filtered results
- Create single JSON file for department backup

**Scenario 3: Selective User Backup**
- Load all synced users
- Use filter to find specific users quickly
- Manually select 5-10 users of interest
- Backup to single file for easy review

**Scenario 4: Regular Audit Trail**
- Weekly/monthly: Load all users
- Select all
- Backup to timestamped folder
- Maintain historical snapshots for compliance

**Scenario 5: Switch SOA to Cloud**
- Backup selected users to JSON
- Clear on-premises attributes for those users
- Switch SOA to Cloud
- Verify in Entra ID that `onPremisesSyncEnabled` is `false`

## Output Formats

The tool supports two backup formats to suit different use cases:

### Format 1: Single JSON File (Multiple Users)

When selecting "Single JSON file", all users are combined in an array:

```json
{
  "BackupMetadata": {
    "BackupDateTime": "2026-02-24 10:30:45",
    "BackupDateTimeUTC": "2026-02-24 15:30:45",
    "BackedUpBy": "admin@contoso.com",
    "ToolVersion": "1.0.1",
    "GraphEndpoint": "Global",
    "TotalUsers": 150
  },
  "Users": [
    {
      "UserData": {
        "Id": "00000000-0000-0000-0000-000000000001",
        "DisplayName": "John Doe",
        "UserPrincipalName": "john.doe@contoso.com",
        "Mail": "john.doe@contoso.com",
        "AccountEnabled": true,
        "UserType": "Member"
      },
      "OnPremisesSyncAttributes": {
        "OnPremisesSyncEnabled": true,
        "OnPremisesLastSyncDateTime": "2026-02-24T09:15:22Z",
        "OnPremisesDistinguishedName": "CN=John Doe,OU=Users,DC=contoso,DC=com",
        "OnPremisesDomainName": "contoso.com",
        "OnPremisesSamAccountName": "jdoe",
        "OnPremisesUserPrincipalName": "jdoe@contoso.local",
        "OnPremisesSecurityIdentifier": "S-1-5-21-...",
        "OnPremisesImmutableId": "abc123==",
        "ProxyAddresses": [
          "SMTP:john.doe@contoso.com",
          "smtp:jdoe@contoso.onmicrosoft.com"
        ]
      }
    },
    {
      "UserData": { "..." },
      "OnPremisesSyncAttributes": { "..." }
    }
  ]
}
```

**Use Cases:**
- Small to medium user sets (1-100 users)
- Quick review and comparison
- Easy to parse programmatically
- Single file for transport or archival

### Format 2: Individual Files in Folder

When selecting "Individual JSON files", structure:

```
UserBackup_20260224_103045/
  ├── _BackupIndex.json              (Index of all files with metadata)
  ├── jdoe_john.doe_20260224_103045.json
  ├── jsmith_jane.smith_20260224_103045.json
  └── ... (one file per user)
```

**Individual User File Structure:**
```json
{
  "BackupMetadata": {
    "BackupDateTime": "2026-02-24 10:30:45",
    "BackupDateTimeUTC": "2026-02-24 15:30:45",
    "BackedUpBy": "admin@contoso.com",
    "ToolVersion": "1.0.1",
    "GraphEndpoint": "Global",
    "UserNumber": 1,
    "TotalInBatch": 150
  },
  "UserData": {
    "Id": "00000000-0000-0000-0000-000000000001",
    "DisplayName": "John Doe",
    "UserPrincipalName": "john.doe@contoso.com",
    "Mail": "john.doe@contoso.com",
    "AccountEnabled": true,
    "UserType": "Member"
  },
  "OnPremisesSyncAttributes": {
    "OnPremisesSyncEnabled": true,
    "OnPremisesLastSyncDateTime": "2026-02-24T09:15:22Z",
    "OnPremisesDistinguishedName": "CN=John Doe,OU=Users,DC=contoso,DC=com",
    "OnPremisesDomainName": "contoso.com",
    "OnPremisesSamAccountName": "jdoe",
    "OnPremisesUserPrincipalName": "jdoe@contoso.local",
    "OnPremisesSecurityIdentifier": "S-1-5-21-...",
    "OnPremisesImmutableId": "abc123==",
    "ProxyAddresses": [
      "SMTP:john.doe@contoso.com",
      "smtp:jdoe@contoso.onmicrosoft.com"
    ]
  }
}
```

**Index File Structure (_BackupIndex.json):**
```json
{
  "BackupMetadata": {
    "BackupDateTime": "2026-02-24 10:30:45",
    "BackupDateTimeUTC": "2026-02-24 15:30:45",
    "BackedUpBy": "admin@contoso.com",
    "ToolVersion": "1.0.1",
    "GraphEndpoint": "Global",
    "TotalUsers": 150,
    "BackupFormat": "IndividualFiles"
  },
  "Files": [
    {
      "FileName": "jdoe_john.doe_20260224_103045.json",
      "UserPrincipalName": "john.doe@contoso.com",
      "DisplayName": "John Doe",
      "SamAccountName": "jdoe"
    }
  ]
}
```

**Use Cases:**
- Large user sets (100+ users)
- Easier per-user review and restoration
- Better for version control (git)
- Distributed processing scenarios
- Lower memory footprint when parsing

## Troubleshooting

### ADSyncTools Module Not Found

**Problem:** "ADSyncTools Not Available" error when clicking Clear On-Prem Attributes or Restore from Backup  
**Solution:**
- The `ADSyncTools` module is installed as part of **Microsoft Azure AD Connect** or **Entra Cloud Sync**
- Run this script on the server where Azure AD Connect or Entra Cloud Sync is installed
- To verify the module is available:
```powershell
Get-Module -ListAvailable -Name ADSyncTools
```
- If it is available but not loading, import it manually:
```powershell
Import-Module ADSyncTools
```

### Module Installation Issues

**Problem:** Script fails to install Microsoft.Graph modules  
**Solution:**
```powershell
# Manually install modules with specific version
Install-Module Microsoft.Graph.Authentication -RequiredVersion 2.35.0 -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Users -RequiredVersion 2.35.0 -Scope CurrentUser -Force
```

### Module Version Conflicts

**Problem:** "Could not load file or assembly 'Microsoft.Graph.Authentication'" error  
**Solution:**
- The script requires version 2.35.0 exactly for both modules
- Remove conflicting versions:
```powershell
Get-Module Microsoft.Graph.* -ListAvailable | Where-Object Version -ne "2.35.0" | Uninstall-Module
```

### Authentication Failures

**Problem:** "Access Denied" or consent errors  
**Solution:**
- Ensure you have appropriate admin permissions
- Ask Global Administrator to pre-consent to the app permissions
- Check if Conditional Access policies are blocking sign-in

### XAML File Not Found

**Problem:** "XAML file not found" error  
**Solution:**
- Ensure `UserBackupUI.xaml` is in the same directory as the script
- Verify file name spelling
- Check file permissions

### Slow User Loading

**Problem:** Loading users takes too long (5+ minutes for 1000+ users)  
**Solution:**
- This is normal for large tenants - Graph API has pagination limits
- Progress bar shows current count
- Consider filtering users by department/location if possible (future enhancement)
- Large tenants may take 10-20 minutes for initial load

### Filter Not Working

**Problem:** Filter doesn't show results  
**Solution:**
- Filter requires minimum 3 characters
- Filter is case-insensitive but checks DisplayName only
- Try broader search terms (e.g., "Smi" instead of "Smith, John")
- Click "Clear Filter" and try again

### DataGrid Performance Issues

**Problem:** UI freezes or is slow with many users  
**Solution:**
- WPF DataGrid virtualizes rows for performance
- If experiencing issues with 5000+ users, use filter to narrow results
- Close and reopen application if UI becomes unresponsive
- Check debug log for errors

### Backup Fails with "Access Denied"

**Problem:** Cannot save backup file  
**Solution:**
- Ensure you have write permissions to selected folder
- Try saving to Documents or Desktop
- Run PowerShell as Administrator if needed (not recommended)
- Check antivirus isn't blocking file creation

### No Users Appear After Loading

**Problem:** "Load Synced Users" completes but DataGrid is empty  
**Solution:**
- Verify your tenant has on-premises synced users
- Check connection status (should say "Connected")
- Review debug log (debug_log_*.txt) for errors
- Try disconnecting and reconnecting to Graph

### "Select All" Not Working

**Problem:** "Select All Visible" doesn't select users  
**Solution:**
- Ensure users are loaded (Total count > 0)
- If filter is active, only visible/filtered users are selected
- Try "Clear All Selections" then try again
- Check debug log for binding errors

## File Structure

```
User-SOA-Switch/
│
├── User-SOA-Switch.ps1      # Main PowerShell script
├── UserBackupUI.xaml         # WPF user interface definition
├── README.md                 # This file
└── .gitignore               # Git ignore rules
```

## Security Considerations

- **Credentials:** The tool uses interactive authentication - never stores passwords
- **Tokens:** Graph API tokens are session-based and cleared on exit
- **Backup Files:** JSON backups may contain sensitive data - store securely
- **Permissions:** Uses least-privilege permissions (read-only for users)
- **Audit:** All operations are logged by Microsoft Graph for audit purposes

## Best Practices

1. **Run with least privilege** - Don't use Global Admin unless required
2. **Secure backup storage** - Encrypt or restrict access to JSON backups
3. **Test first** - Test with a non-critical user before bulk operations
4. **Document usage** - Keep records of when/why backups were created
5. **Regular updates** - Keep Microsoft.Graph modules updated

## Technical Details

### Microsoft Graph API Endpoints

This tool uses the **Microsoft Graph v1.0 API** for all operations:
- **User queries**: `https://graph.microsoft.com/v1.0/users`
- **SOA management**: `https://graph.microsoft.com/v1.0/users/{id}/onPremisesSyncBehavior`

The `onPremisesSyncBehavior` endpoint is used to manage the Source of Authority:
- **Switch SOA to Cloud**: Sets `isCloudManaged = true` (Phase 1) and clears on-premises attributes using ADSyncTools (Phase 2)
- **Rollback SOA to On-Prem**: Sets `isCloudManaged = false`

### ADSyncTools Cmdlets

On-premises user attributes are **read-only** via the Microsoft Graph API. To clear or restore them, this tool uses the **ADSyncTools** PowerShell module (included with Azure AD Connect / Entra Cloud Sync):

- **Clear attributes**: `Clear-ADSyncToolsOnPremisesAttribute`
  ```
  Clear-ADSyncToolsOnPremisesAttribute [-Id] <String>
      [[-onPremisesDistinguishedName]] [[-onPremisesDomainName]]
      [[-onPremisesImmutableId]] [[-onPremisesSamAccountName]]
      [[-onPremisesSecurityIdentifier]] [[-onPremisesUserPrincipalName]]
  ```

- **Restore attributes**: `Set-ADSyncToolsOnPremisesAttribute`
  ```
  Set-ADSyncToolsOnPremisesAttribute [-Id] <String>
      [[-onPremisesDistinguishedName] <String>] [[-onPremisesDomainName] <String>]
      [[-onPremisesImmutableId] <String>] [[-onPremisesSamAccountName] <String>]
      [[-onPremisesSecurityIdentifier] <String>] [[-onPremisesUserPrincipalName] <String>]
  ```

### SOA Operation Flow

**Switch to Cloud (Two-Phase):**
1. **Phase 1**: PATCH to `/onPremisesSyncBehavior` with `{"isCloudManaged": true}`
2. **Verification**: GET to `/onPremisesSyncBehavior` confirms `isCloudManaged = true`
3. **Phase 2**: `Clear-ADSyncToolsOnPremisesAttribute` clears on-premises sync attributes

**Rollback to On-Premises (Single-Phase):**
1. **Verification**: GET to `/onPremisesSyncBehavior` confirms `isCloudManaged = true`
2. **Rollback**: PATCH to `/onPremisesSyncBehavior` with `{"isCloudManaged": false}`

All SOA management operations use the stable v1.0 endpoint (not beta) and require the `User-OnPremisesSyncBehavior.ReadWrite.All` permission.

## Known Limitations

- Only retrieves one user at a time (no bulk operations)
- Requires PowerShell 7+ (not compatible with Windows PowerShell 5.1)
- Windows-only (due to WPF dependency)
- SOA switch and rollback operations require `User-OnPremisesSyncBehavior.ReadWrite.All` permission
- Clear On-Prem Attributes and Restore from Backup require the `ADSyncTools` module (installed with Azure AD Connect / Entra Cloud Sync)

## Version History

**Version 1.3** (February 2026)
- Replaced read-only Graph API PATCH calls with `Clear-ADSyncToolsOnPremisesAttribute` for clearing on-premises attributes
- Replaced Graph API PATCH calls with `Set-ADSyncToolsOnPremisesAttribute` for restoring on-premises attributes from backup
- Added `Initialize-ADSyncToolsModule` function for ADSyncTools module management
- Updated README to document ADSyncTools module dependency

**Version 1.2** (February 2026)
- Updated all SOA operations to use Microsoft Graph v1.0 endpoint (previously beta)
- Added `Rollback-UserSOA` function to revert SOA from Cloud back to On-Premises
- Enhanced verification steps for SOA operations with detailed logging
- Two-phase SOA switch process with verification between phases

**Version 1.1** (February 2026)
- Added SOA management workflow: Clear On-Prem Attributes, Switch SOA to Cloud, Restore from Backup
- Added SOA warning banner with prerequisite sequence guidance
- Added `Switch-UserSOA` function using Microsoft Graph v1.0 API

**Version 1.0** (February 23, 2026)
- Initial release
- Basic user search and backup functionality
- Microsoft Graph integration
- WPF-based GUI

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Submit a pull request with detailed description

## Support

For issues or questions:
- Check the Troubleshooting section above
- Review PowerShell error output
- Check Microsoft Graph API documentation

## License

This project is provided as-is for educational and administrative purposes.

## Acknowledgments

- Built with Microsoft Graph PowerShell SDK
- WPF interface design using XAML
- PowerShell 7+ Core features

---

**Disclaimer:** Always test in a non-production environment first. Ensure you have proper authorization before accessing user data.
