#Requires -Version 7.0

<#
.SYNOPSIS
    Entra ID User Sync Attribute Backup Tool
.DESCRIPTION
    WPF-based GUI tool to backup on-premises synced attributes from Entra ID users
    using Microsoft Graph API with User-OnPremisesSyncBehavior.ReadWrite.All permission.
.NOTES
    Author: User-SOA-Switch Project
    Date: February 23, 2026
    Requires: PowerShell 7+, Microsoft.Graph module
#>

# Load required assemblies for WPF
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Global variables for multi-user selection
$script:IsConnected = $false
$script:AllUsers = @()  # Master list of all loaded users
$script:UserCollection = $null  # ObservableCollection for DataGrid binding
$script:IsUsersLoaded = $false
$script:CurrentUserType = $null  # 'OnPrem' or 'Cloud' based on loaded user set
$script:ADSyncToolsAvailable = $false  # Tracks whether ADSyncTools module is loaded

# Debug logging setup
$script:LogFile = Join-Path $PSScriptRoot "debug_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$script:DebugEnabled = $true

function Write-DebugLog {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )
    
    if ($script:DebugEnabled) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $logMessage = "[$timestamp] [$Level] $Message"
        
        # Write to log file
        Add-Content -Path $script:LogFile -Value $logMessage -ErrorAction SilentlyContinue
        
        # Also write to console with color
        $color = switch ($Level) {
            'SUCCESS' { 'Green' }
            'WARNING' { 'Yellow' }
            'ERROR' { 'Red' }
            default { 'Cyan' }
        }
        Write-Host $logMessage -ForegroundColor $color
    }
}

# Function to check and install Microsoft.Graph module
function Initialize-GraphModule {
    param()
    
    try {
        Write-DebugLog "Starting Microsoft.Graph module initialization" -Level INFO
        Write-Host "Checking for Microsoft.Graph module..." -ForegroundColor Cyan
        
        # Define required module version to ensure compatibility
        $targetVersion = "2.35.0"
        $requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
        Write-DebugLog "Required modules: $($requiredModules -join ', ') @ v$targetVersion" -Level INFO
        
        foreach ($moduleName in $requiredModules) {
            Write-DebugLog "Processing module: $moduleName" -Level INFO
            
            # Check if correct version is already imported
            $loadedModule = Get-Module -Name $moduleName
            if ($loadedModule -and $loadedModule.Version.ToString().StartsWith($targetVersion)) {
                Write-DebugLog "$moduleName v$($loadedModule.Version) is already loaded" -Level SUCCESS
                Write-Host "$moduleName is already loaded." -ForegroundColor Green
                continue
            }
            
            # Check if target version is available
            $availableModule = Get-Module -ListAvailable -Name $moduleName | 
                               Where-Object { $_.Version.ToString().StartsWith($targetVersion) } | 
                               Select-Object -First 1
            
            if (-not $availableModule) {
                Write-DebugLog "$moduleName v$targetVersion not found. Installing..." -Level WARNING
                Write-Host "Installing $moduleName v$targetVersion..." -ForegroundColor Yellow
                Install-Module -Name $moduleName -RequiredVersion $targetVersion -Scope CurrentUser -Force -AllowClobber
                Write-DebugLog "$moduleName v$targetVersion installed successfully" -Level SUCCESS
            }
            else {
                Write-DebugLog "$moduleName found. Version: $($availableModule.Version)" -Level INFO
            }
            
            # Import the specific version
            Write-DebugLog "Importing $moduleName v$targetVersion..." -Level INFO
            Write-Host "Importing $moduleName..." -ForegroundColor Cyan
            Import-Module $moduleName -RequiredVersion $targetVersion -ErrorAction Stop -Force
            $imported = Get-Module -Name $moduleName
            Write-DebugLog "$moduleName v$($imported.Version) imported successfully" -Level SUCCESS
        }
        
        Write-Host "Microsoft.Graph modules loaded successfully." -ForegroundColor Green
        Write-DebugLog "All Microsoft.Graph modules initialized successfully" -Level SUCCESS
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugLog "ERROR initializing Graph modules: $errorMsg" -Level ERROR
        Write-DebugLog "Exception Type: $($_.Exception.GetType().FullName)" -Level ERROR
        Write-DebugLog "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
        Write-Host "Error initializing Graph modules: $_" -ForegroundColor Red
        return $false
    }
}

# Function to check and install ADSyncTools module
function Initialize-ADSyncToolsModule {
    param()
    
    try {
        Write-DebugLog "Starting ADSyncTools module initialization" -Level INFO
        Write-Host "Checking for ADSyncTools module..." -ForegroundColor Cyan
        
        $moduleName = 'ADSyncTools'
        
        # Check if already imported
        $loadedModule = Get-Module -Name $moduleName
        if ($loadedModule) {
            Write-DebugLog "$moduleName v$($loadedModule.Version) is already loaded" -Level SUCCESS
            Write-Host "$moduleName is already loaded." -ForegroundColor Green
            $script:ADSyncToolsAvailable = $true
            return $true
        }
        
        # Check if available on this system
        $availableModule = Get-Module -ListAvailable -Name $moduleName | Select-Object -First 1
        
        if (-not $availableModule) {
            Write-DebugLog "$moduleName not found. Installing from PowerShell Gallery..." -Level WARNING
            Write-Host "$moduleName not found. Installing from PowerShell Gallery..." -ForegroundColor Yellow
            
            try {
                Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-DebugLog "$moduleName installed successfully from PowerShell Gallery" -Level SUCCESS
                Write-Host "$moduleName installed successfully." -ForegroundColor Green
                
                # Re-check for the newly installed module
                $availableModule = Get-Module -ListAvailable -Name $moduleName | Select-Object -First 1
            }
            catch {
                $installError = $_.Exception.Message
                Write-DebugLog "Failed to install $moduleName from PowerShell Gallery: $installError" -Level ERROR
                Write-Host "ERROR: Failed to install $moduleName from PowerShell Gallery." -ForegroundColor Red
                Write-Host "Error: $installError" -ForegroundColor Red
                Write-Host "Clear On-Prem Attributes and Restore from Backup features will be unavailable." -ForegroundColor Yellow
                $script:ADSyncToolsAvailable = $false
                return $false
            }
        }
        
        if ($availableModule) {
            Write-DebugLog "$moduleName found. Version: $($availableModule.Version)" -Level INFO
            Import-Module $moduleName -ErrorAction Stop
            $imported = Get-Module -Name $moduleName
            Write-DebugLog "$moduleName v$($imported.Version) imported successfully" -Level SUCCESS
            Write-Host "$moduleName loaded successfully." -ForegroundColor Green
            $script:ADSyncToolsAvailable = $true
            return $true
        }
        else {
            Write-DebugLog "$moduleName still not available after installation attempt" -Level ERROR
            $script:ADSyncToolsAvailable = $false
            return $false
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugLog "ERROR initializing ADSyncTools module: $errorMsg" -Level ERROR
        Write-Host "Error initializing ADSyncTools module: $_" -ForegroundColor Red
        $script:ADSyncToolsAvailable = $false
        return $false
    }
}

# Function to connect to Microsoft Graph
function Connect-ToGraph {
    param(
        [System.Windows.Controls.Label]$StatusLabel,
        [System.Windows.Controls.Label]$UserLabel,
        [System.Windows.Controls.TextBlock]$StatusBar
    )
    
    try {
        $StatusBar.Text = "Connecting to Microsoft Graph..."
        
        # Define required scopes
        $scopes = @(
            'User.Read.All',
            'User-OnPremisesSyncBehavior.ReadWrite.All'
        )
        
        # Connect with interactive authentication
        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
        
        # Get connection context
        $context = Get-MgContext
        
        if ($context) {
            $script:IsConnected = $true
            $StatusLabel.Content = "Connected"
            $StatusLabel.Foreground = "Green"
            $UserLabel.Content = $context.Account
            $UserLabel.Foreground = "Green"
            $StatusBar.Text = "Successfully connected to Microsoft Graph as $($context.Account)"
            
            return $true
        }
        else {
            throw "Failed to establish Graph connection"
        }
    }
    catch {
        $script:IsConnected = $false
        $StatusLabel.Content = "Connection Failed"
        $StatusLabel.Foreground = "Red"
        $StatusBar.Text = "Error connecting to Graph: $($_.Exception.Message)"
        
        [System.Windows.MessageBox]::Show(
            "Failed to connect to Microsoft Graph:`n`n$($_.Exception.Message)",
            "Connection Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        
        return $false
    }
}

# Function to load all synced users from Entra ID
function Load-AllSyncedUsers {
    param(
        [System.Windows.Controls.ProgressBar]$ProgressBar,
        [System.Windows.Controls.TextBlock]$ProgressStatus,
        [System.Windows.Controls.TextBlock]$ProgressDetail,
        [System.Windows.Controls.TextBlock]$StatusBar
    )
    
    try {
        $startTime = Get-Date
        Write-DebugLog "Starting to load all synced users" -Level INFO
        $StatusBar.Text = "Loading all on-premises synced users from Entra ID..."
        
        # Query all users with OnPremisesSyncEnabled filter
        Write-DebugLog "Querying Graph API with filter: onPremisesSyncEnabled eq true" -Level INFO
        
        $allUsers = @()
        $properties = @(
            'Id',
            'DisplayName',
            'UserPrincipalName',
            'Mail',
            'OnPremisesSyncEnabled',
            'OnPremisesDistinguishedName',
            'OnPremisesDomainName',
            'OnPremisesSamAccountName',
            'OnPremisesUserPrincipalName',
            'OnPremisesSecurityIdentifier',
            'OnPremisesImmutableId',
            'OnPremisesLastSyncDateTime',
            'ProxyAddresses',
            'AccountEnabled',
            'UserType'
        )
        
        # Get all synced users with pagination support
        $users = Get-MgUser -Filter "onPremisesSyncEnabled eq true" `
                           -All `
                           -Property $properties `
                           -ConsistencyLevel eventual `
                           -CountVariable userCount `
                           -ErrorAction Stop
        
        $total = ($users | Measure-Object).Count
        Write-DebugLog "Query returned $total synced users" -Level SUCCESS
        
        if ($total -eq 0) {
            $StatusBar.Text = "No on-premises synced users found in this tenant"
            Write-DebugLog "No synced users found" -Level WARNING
            return @()
        }
        
        # Update progress bar setup
        $ProgressBar.Maximum = $total
        $ProgressBar.Value = 0
        $current = 0
        
        # Process users and add IsSelected property
        foreach ($user in $users) {
            $current++
            
            # Update progress every 50 users or on last user
            if (($current % 50 -eq 0) -or ($current -eq $total)) {
                $ProgressBar.Value = $current
                $ProgressStatus.Text = "Loading synced users: $current of $total"
                $percent = [math]::Round(($current / $total) * 100, 1)
                $ProgressDetail.Text = "$percent% complete"
                
                # Allow UI to update
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            # Query isCloudManaged from onPremisesSyncBehavior endpoint
            $isCloudManaged = "Unknown"
            try {
                $syncBehaviorUri = "https://graph.microsoft.com/v1.0/users/$($user.Id)/onPremisesSyncBehavior"
                $syncBehavior = Invoke-MgGraphRequest -Method GET -Uri $syncBehaviorUri -ErrorAction SilentlyContinue
                if ($null -ne $syncBehavior -and $null -ne $syncBehavior.isCloudManaged) {
                    $isCloudManaged = if ($syncBehavior.isCloudManaged -eq $true) { "True" } else { "False" }
                }
            }
            catch {
                # If endpoint doesn't exist or errors, leave as Unknown
                $isCloudManaged = "Unknown"
            }
            
            # Create enriched user object with IsSelected property
            $userObject = [PSCustomObject]@{
                IsSelected = $false
                Id = $user.Id
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Mail = $user.Mail
                IsCloudManaged = $isCloudManaged
                OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
                OnPremisesDistinguishedName = $user.OnPremisesDistinguishedName
                OnPremisesDomainName = $user.OnPremisesDomainName
                OnPremisesSamAccountName = $user.OnPremisesSamAccountName
                OnPremisesUserPrincipalName = $user.OnPremisesUserPrincipalName
                OnPremisesSecurityIdentifier = $user.OnPremisesSecurityIdentifier
                OnPremisesImmutableId = $user.OnPremisesImmutableId
                OnPremisesLastSyncDateTime = $user.OnPremisesLastSyncDateTime
                ProxyAddresses = $user.ProxyAddresses
                AccountEnabled = $user.AccountEnabled
                UserType = $user.UserType
            }
            
            $allUsers += $userObject
        }
        
        $endTime = Get-Date
        $duration = ($endTime - $startTime).TotalSeconds
        Write-DebugLog "Loaded $total users in $([math]::Round($duration, 2)) seconds" -Level SUCCESS
        $StatusBar.Text = "Successfully loaded $total on-premises synced users"
        
        return $allUsers
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugLog "ERROR loading users: $errorMsg" -Level ERROR
        Write-DebugLog "Exception Type: $($_.Exception.GetType().FullName)" -Level ERROR
        $StatusBar.Text = "Error loading users: $errorMsg"
        
        [System.Windows.MessageBox]::Show(
            "Failed to load synced users:`n`n$errorMsg",
            "Load Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        
        return @()
    }
}

# Function to load all cloud-managed (non-synced) users from Entra ID
function Load-AllCloudUsers {
    param(
        [System.Windows.Controls.ProgressBar]$ProgressBar,
        [System.Windows.Controls.TextBlock]$ProgressStatus,
        [System.Windows.Controls.TextBlock]$ProgressDetail,
        [System.Windows.Controls.TextBlock]$StatusBar
    )
    
    try {
        $startTime = Get-Date
        Write-DebugLog "Starting to load all cloud-managed (non-synced) users" -Level INFO
        $StatusBar.Text = "Loading all cloud-managed (non-synced) users from Entra ID..."
        
        $properties = @(
            'Id',
            'DisplayName',
            'UserPrincipalName',
            'Mail',
            'OnPremisesSyncEnabled',
            'OnPremisesDistinguishedName',
            'OnPremisesDomainName',
            'OnPremisesSamAccountName',
            'OnPremisesUserPrincipalName',
            'OnPremisesSecurityIdentifier',
            'OnPremisesImmutableId',
            'OnPremisesLastSyncDateTime',
            'ProxyAddresses',
            'AccountEnabled',
            'UserType'
        )
        
        # Get all non-synced users via Graph API v1.0
        Write-DebugLog "Querying Graph API with filter: onPremisesSyncEnabled ne true" -Level INFO
        $users = Get-MgUser -Filter "onPremisesSyncEnabled ne true" `
                            -All `
                            -Property $properties `
                            -ConsistencyLevel eventual `
                            -CountVariable userCount `
                            -ErrorAction Stop
        
        $total = ($users | Measure-Object).Count
        Write-DebugLog "Query returned $total cloud-managed users" -Level SUCCESS
        
        if ($total -eq 0) {
            $StatusBar.Text = "No cloud-managed users found in this tenant"
            Write-DebugLog "No cloud-managed users found" -Level WARNING
            return @()
        }
        
        $ProgressBar.Maximum = $total
        $ProgressBar.Value = 0
        $current = 0
        
        $allUsers = @()
        foreach ($user in $users) {
            $current++
            
            if (($current % 50 -eq 0) -or ($current -eq $total)) {
                $ProgressBar.Value = $current
                $ProgressStatus.Text = "Loading cloud users: $current of $total"
                $percent = [math]::Round(($current / $total) * 100, 1)
                $ProgressDetail.Text = "$percent% complete"
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            # Query isCloudManaged from onPremisesSyncBehavior endpoint
            $isCloudManaged = "Unknown"
            try {
                $syncBehaviorUri = "https://graph.microsoft.com/v1.0/users/$($user.Id)/onPremisesSyncBehavior"
                $syncBehavior = Invoke-MgGraphRequest -Method GET -Uri $syncBehaviorUri -ErrorAction SilentlyContinue
                if ($null -ne $syncBehavior -and $null -ne $syncBehavior.isCloudManaged) {
                    $isCloudManaged = if ($syncBehavior.isCloudManaged -eq $true) { "True" } else { "False" }
                }
            }
            catch {
                # If endpoint doesn't exist or errors, leave as Unknown
                $isCloudManaged = "Unknown"
            }
            
            $userObject = [PSCustomObject]@{
                IsSelected = $false
                Id = $user.Id
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Mail = $user.Mail
                IsCloudManaged = $isCloudManaged
                OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
                OnPremisesDistinguishedName = $user.OnPremisesDistinguishedName
                OnPremisesDomainName = $user.OnPremisesDomainName
                OnPremisesSamAccountName = $user.OnPremisesSamAccountName
                OnPremisesUserPrincipalName = $user.OnPremisesUserPrincipalName
                OnPremisesSecurityIdentifier = $user.OnPremisesSecurityIdentifier
                OnPremisesImmutableId = $user.OnPremisesImmutableId
                OnPremisesLastSyncDateTime = $user.OnPremisesLastSyncDateTime
                ProxyAddresses = $user.ProxyAddresses
                AccountEnabled = $user.AccountEnabled
                UserType = $user.UserType
            }
            
            $allUsers += $userObject
        }
        
        $endTime = Get-Date
        $duration = ($endTime - $startTime).TotalSeconds
        Write-DebugLog "Loaded $total cloud users in $([math]::Round($duration, 2)) seconds" -Level SUCCESS
        $StatusBar.Text = "Successfully loaded $total cloud-managed (non-synced) users"
        
        return $allUsers
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugLog "ERROR loading cloud users: $errorMsg" -Level ERROR
        Write-DebugLog "Exception Type: $($_.Exception.GetType().FullName)" -Level ERROR
        $StatusBar.Text = "Error loading cloud users: $errorMsg"
        
        [System.Windows.MessageBox]::Show(
            "Failed to load cloud-managed users:`n`n$errorMsg",
            "Load Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        
        return @()
    }
}

# Function to update the DataGrid with user collection
function Update-UserGrid {
    param(
        [System.Windows.Controls.DataGrid]$DataGrid,
        [Array]$Users
    )
    
    try {
        Write-DebugLog "Updating DataGrid with $($Users.Count) users" -Level INFO
        
        # Create observable collection for data binding
        $collection = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
        
        foreach ($user in $Users) {
            $collection.Add($user)
        }
        
        $DataGrid.ItemsSource = $collection
        $script:UserCollection = $collection
        
        Write-DebugLog "DataGrid updated successfully" -Level SUCCESS
    }
    catch {
        Write-DebugLog "ERROR updating DataGrid: $($_.Exception.Message)" -Level ERROR
    }
}

# Function to apply filter to users based on display name
function Apply-UserFilter {
    param(
        [string]$FilterText,
        [System.Windows.Controls.DataGrid]$DataGrid,
        [System.Windows.Controls.TextBlock]$FilteredCountLabel
    )
    
    try {
        $filterStart = Get-Date
        
        # If filter text is less than 3 characters, show all users
        if ([string]::IsNullOrWhiteSpace($FilterText) -or $FilterText.Length -lt 3) {
            Write-DebugLog "Filter text too short or empty. Showing all users" -Level INFO
            Update-UserGrid -DataGrid $DataGrid -Users $script:AllUsers
            $FilteredCountLabel.Text = $script:AllUsers.Count
            return
        }
        
        Write-DebugLog "Applying filter: '$FilterText'" -Level INFO
        
        # Filter users by display name (case-insensitive)
        $filteredUsers = $script:AllUsers | Where-Object {
            $_.DisplayName -like "*$FilterText*"
        }
        
        $filterDuration = ((Get-Date) - $filterStart).TotalMilliseconds
        Write-DebugLog "Filter applied. Found $($filteredUsers.Count) of $($script:AllUsers.Count) users in $([math]::Round($filterDuration, 0))ms" -Level SUCCESS
        
        # Update DataGrid with filtered results
        Update-UserGrid -DataGrid $DataGrid -Users $filteredUsers
        $FilteredCountLabel.Text = $filteredUsers.Count
    }
    catch {
        Write-DebugLog "ERROR applying filter: $($_.Exception.Message)" -Level ERROR
    }
}

# Function to get all selected users
function Get-SelectedUsers {
    try {
        $selected = $script:AllUsers | Where-Object { $_.IsSelected -eq $true }
        Write-DebugLog "Get-SelectedUsers returned $($selected.Count) users" -Level INFO
        return $selected
    }
    catch {
        Write-DebugLog "ERROR getting selected users: $($_.Exception.Message)" -Level ERROR
        return @()
    }
}

# Function to update selection counts in UI
function Update-SelectionCounts {
    param(
        [System.Windows.Controls.TextBlock]$TotalCountLabel,
        [System.Windows.Controls.TextBlock]$FilteredCountLabel,
        [System.Windows.Controls.TextBlock]$SelectedCountLabel,
        [System.Windows.Controls.Button]$BackupButton
    )
    
    try {
        $totalCount = $script:AllUsers.Count
        $selectedUsers = Get-SelectedUsers
        $selectedCount = $selectedUsers.Count
        
        $TotalCountLabel.Text = $totalCount
        $SelectedCountLabel.Text = $selectedCount
        
        # Update backup button
        $BackupButton.Content = "Backup Selected Users ($selectedCount)"
        $BackupButton.IsEnabled = ($selectedCount -gt 0)
        
        # Update SOA action buttons if they have been bound
        if ($script:btnClearAttributes) {
            $script:btnClearAttributes.IsEnabled = ($selectedCount -gt 0)
        }
        if ($script:btnSwitchSOA) {
            $script:btnSwitchSOA.IsEnabled = ($selectedCount -gt 0)
        }
        if ($script:btnRollbackSOA) {
            $script:btnRollbackSOA.IsEnabled = ($selectedCount -gt 0)
        }
        
        Write-DebugLog "Updated counts - Total: $totalCount, Selected: $selectedCount" -Level INFO
    }
    catch {
        Write-DebugLog "ERROR updating selection counts: $($_.Exception.Message)" -Level ERROR
    }
}

# Function to show/hide SOA action buttons based on the type of users currently loaded
function Update-UIForUserType {
    param(
        [string]$UserType
    )
    
    $script:CurrentUserType = $UserType
    Write-DebugLog "Updating UI for user type: $UserType" -Level INFO
    
    if ($UserType -eq 'Cloud') {
        # Cloud-managed users: show ClearAttributes, RollbackSOA and RestoreBackup; hide SwitchSOA
        $script:btnClearAttributes.Visibility = [System.Windows.Visibility]::Visible
        $script:btnSwitchSOA.Visibility = [System.Windows.Visibility]::Collapsed
        $script:btnRollbackSOA.Visibility = [System.Windows.Visibility]::Visible
        $script:btnRestoreBackup.Visibility = [System.Windows.Visibility]::Visible
    }
    else {
        # On-premises synced users or default: show all SOA action buttons
        $script:btnClearAttributes.Visibility = [System.Windows.Visibility]::Visible
        $script:btnSwitchSOA.Visibility = [System.Windows.Visibility]::Visible
        $script:btnRollbackSOA.Visibility = [System.Windows.Visibility]::Visible
        $script:btnRestoreBackup.Visibility = [System.Windows.Visibility]::Visible
    }
}

# Function to backup selected users to JSON
function Backup-SelectedUsersToJson {
    param(
        [System.Windows.Controls.TextBlock]$StatusBar
    )
    
    try {
        $selectedUsers = Get-SelectedUsers
        
        if ($selectedUsers.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No users selected for backup.`n`nPlease select at least one user from the list.",
                "No Selection",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return $false
        }
        
        Write-DebugLog "Starting backup for $($selectedUsers.Count) selected users" -Level INFO
        $StatusBar.Text = "Preparing backup for $($selectedUsers.Count) users..."
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $context = Get-MgContext
        
        # For single user, skip format dialog and use individual file format
        if ($selectedUsers.Count -eq 1) {
            Write-DebugLog "Single user selected - defaulting to individual file format" -Level INFO
            $useSingleFile = $false
        }
        else {
            # Show format selection dialog for multiple users
            $formatDialog = New-Object System.Windows.Window
            $formatDialog.Title = "Choose Backup Format"
            $formatDialog.Width = 450
            $formatDialog.Height = 250
            $formatDialog.WindowStartupLocation = "CenterScreen"
            $formatDialog.ResizeMode = "NoResize"
            
            $grid = New-Object System.Windows.Controls.Grid
            $grid.Margin = "15"
            
            $stackPanel = New-Object System.Windows.Controls.StackPanel
            
            $labelText = New-Object System.Windows.Controls.TextBlock
            $labelText.Text = "Select backup format for $($selectedUsers.Count) users:"
            $labelText.FontWeight = "Bold"
            $labelText.Margin = "0,0,0,15"
            $stackPanel.AddChild($labelText)
            
            $radioSingle = New-Object System.Windows.Controls.RadioButton
            $radioSingle.Content = "Single JSON file (all users in one array)"
            $radioSingle.IsChecked = $true
            $radioSingle.Margin = "0,0,0,10"
            $stackPanel.AddChild($radioSingle)
            
            $radioMultiple = New-Object System.Windows.Controls.RadioButton
            $radioMultiple.Content = "Individual JSON files (one file per user in a folder)"
            $radioMultiple.Margin = "0,0,0,20"
            $stackPanel.AddChild($radioMultiple)
            
            $buttonPanel = New-Object System.Windows.Controls.StackPanel
            $buttonPanel.Orientation = "Horizontal"
            $buttonPanel.HorizontalAlignment = "Right"
            
            $btnOK = New-Object System.Windows.Controls.Button
            $btnOK.Content = "Continue"
            $btnOK.Width = 80
            $btnOK.Height = 30
            $btnOK.Margin = "0,0,10,0"
            $btnOK.IsDefault = $true
            $btnOK.Add_Click({
                $formatDialog.DialogResult = $true
                $formatDialog.Close()
            })
            $buttonPanel.AddChild($btnOK)
            
            $btnCancel = New-Object System.Windows.Controls.Button
            $btnCancel.Content = "Cancel"
            $btnCancel.Width = 80
            $btnCancel.Height = 30
            $btnCancel.IsCancel = $true
            $btnCancel.Add_Click({
                $formatDialog.DialogResult = $false
                $formatDialog.Close()
            })
            $buttonPanel.AddChild($btnCancel)
            
            $stackPanel.AddChild($buttonPanel)
            $grid.AddChild($stackPanel)
            $formatDialog.Content = $grid
            
            $result = $formatDialog.ShowDialog()
            
            if (-not $result) {
                $StatusBar.Text = "Backup cancelled"
                Write-DebugLog "Backup cancelled by user" -Level INFO
                return $false
            }
            
            $useSingleFile = $radioSingle.IsChecked
        }
        
        if ($useSingleFile) {
            # Single JSON file for all users
            Write-DebugLog "User chose single JSON file format" -Level INFO
            
            $usersArray = @()
            foreach ($user in $selectedUsers) {
                $userBackup = [PSCustomObject]@{
                    UserData = [PSCustomObject]@{
                        Id = $user.Id
                        DisplayName = $user.DisplayName
                        UserPrincipalName = $user.UserPrincipalName
                        Mail = $user.Mail
                        AccountEnabled = $user.AccountEnabled
                        UserType = $user.UserType
                    }
                    OnPremisesSyncAttributes = [PSCustomObject]@{
                        OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
                        OnPremisesLastSyncDateTime = $user.OnPremisesLastSyncDateTime
                        OnPremisesDistinguishedName = $user.OnPremisesDistinguishedName
                        OnPremisesDomainName = $user.OnPremisesDomainName
                        OnPremisesSamAccountName = $user.OnPremisesSamAccountName
                        OnPremisesUserPrincipalName = $user.OnPremisesUserPrincipalName
                        OnPremisesSecurityIdentifier = $user.OnPremisesSecurityIdentifier
                        OnPremisesImmutableId = $user.OnPremisesImmutableId
                        ProxyAddresses = $user.ProxyAddresses
                    }
                }
                $usersArray += $userBackup
            }
            
            $backup = [PSCustomObject]@{
                BackupMetadata = [PSCustomObject]@{
                    BackupDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    BackupDateTimeUTC = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
                    BackedUpBy = $context.Account
                    ToolVersion = "1.0.1"
                    GraphEndpoint = $context.Environment
                    TotalUsers = $selectedUsers.Count
                }
                Users = $usersArray
            }
            
            $json = $backup | ConvertTo-Json -Depth 10
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            $saveDialog.DefaultExt = ".json"
            $saveDialog.FileName = "SyncedUsers_Backup_${timestamp}.json"
            $saveDialog.Title = "Save Combined User Backup"
            
            $dialogResult = $saveDialog.ShowDialog()
            
            if ($dialogResult -eq $true) {
                $filePath = $saveDialog.FileName
                $json | Out-File -FilePath $filePath -Encoding UTF8 -Force
                
                Write-DebugLog "Backup saved to: $filePath" -Level SUCCESS
                $StatusBar.Text = "Backup saved successfully: $filePath"
                
                [System.Windows.MessageBox]::Show(
                    "Successfully backed up $($selectedUsers.Count) users!`n`nFile: $filePath",
                    "Backup Complete",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                )
                
                return $true
            }
            else {
                $StatusBar.Text = "Backup cancelled"
                return $false
            }
        }
        else {
            # Individual file(s)
            
            if ($selectedUsers.Count -eq 1) {
                # Single user - save directly to chosen location without creating folder
                Write-DebugLog "Single user - saving individual file" -Level INFO
                
                $user = $selectedUsers[0]
                $safeSam = $user.OnPremisesSamAccountName -replace '[\\/:*?"<>|]', '_'
                $safeUpn = ($user.UserPrincipalName -split '@')[0] -replace '[\\/:*?"<>|]', '_'
                $defaultFileName = "${safeSam}_${safeUpn}_${timestamp}.json"
                
                $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
                $saveDialog.DefaultExt = ".json"
                $saveDialog.FileName = $defaultFileName
                $saveDialog.Title = "Save User Backup"
                
                $dialogResult = $saveDialog.ShowDialog()
                
                if ($dialogResult -eq $true) {
                    $filePath = $saveDialog.FileName
                    
                    $userBackup = [PSCustomObject]@{
                        BackupMetadata = [PSCustomObject]@{
                            BackupDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            BackupDateTimeUTC = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
                            BackedUpBy = $context.Account
                            ToolVersion = "1.0.1"
                            GraphEndpoint = $context.Environment
                        }
                        UserData = [PSCustomObject]@{
                            Id = $user.Id
                            DisplayName = $user.DisplayName
                            UserPrincipalName = $user.UserPrincipalName
                            Mail = $user.Mail
                            AccountEnabled = $user.AccountEnabled
                            UserType = $user.UserType
                        }
                        OnPremisesSyncAttributes = [PSCustomObject]@{
                            OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
                            OnPremisesLastSyncDateTime = $user.OnPremisesLastSyncDateTime
                            OnPremisesDistinguishedName = $user.OnPremisesDistinguishedName
                            OnPremisesDomainName = $user.OnPremisesDomainName
                            OnPremisesSamAccountName = $user.OnPremisesSamAccountName
                            OnPremisesUserPrincipalName = $user.OnPremisesUserPrincipalName
                            OnPremisesSecurityIdentifier = $user.OnPremisesSecurityIdentifier
                            OnPremisesImmutableId = $user.OnPremisesImmutableId
                            ProxyAddresses = $user.ProxyAddresses
                        }
                    }
                    
                    $json = $userBackup | ConvertTo-Json -Depth 10
                    $json | Out-File -FilePath $filePath -Encoding UTF8 -Force
                    
                    Write-DebugLog "Backup saved to: $filePath" -Level SUCCESS
                    $StatusBar.Text = "Backup saved successfully: $filePath"
                    
                    [System.Windows.MessageBox]::Show(
                        "Successfully backed up user: $($user.DisplayName)!`n`nFile: $filePath",
                        "Backup Complete",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information
                    )
                    
                    return $true
                }
                else {
                    $StatusBar.Text = "Backup cancelled"
                    return $false
                }
            }
            else {
                # Multiple users - create folder with individual files and index
                Write-DebugLog "Multiple users - creating folder with individual files" -Level INFO
                
                Add-Type -AssemblyName System.Windows.Forms
                $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                $folderDialog.Description = "Select folder for user backup files"
                $folderDialog.ShowNewFolderButton = $true
                
                if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $backupFolderName = "UserBackup_${timestamp}"
                    $backupPath = Join-Path $folderDialog.SelectedPath $backupFolderName
                    New-Item -ItemType Directory -Path $backupPath -Force | Out-Null
                    
                    Write-DebugLog "Created backup folder: $backupPath" -Level INFO
                    
                    $filesCreated = @()
                    $index = 0
                    
                    foreach ($user in $selectedUsers) {
                        $index++
                        $safeSam = $user.OnPremisesSamAccountName -replace '[\\/:*?"<>|]', '_'
                        $safeUpn = ($user.UserPrincipalName -split '@')[0] -replace '[\\/:*?"<>|]', '_'
                        $fileName = "${safeSam}_${safeUpn}_${timestamp}.json"
                        $filePath = Join-Path $backupPath $fileName
                        
                        $userBackup = [PSCustomObject]@{
                            BackupMetadata = [PSCustomObject]@{
                                BackupDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                                BackupDateTimeUTC = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
                                BackedUpBy = $context.Account
                                ToolVersion = "1.0.1"
                                GraphEndpoint = $context.Environment
                                UserNumber = $index
                                TotalInBatch = $selectedUsers.Count
                            }
                            UserData = [PSCustomObject]@{
                                Id = $user.Id
                                DisplayName = $user.DisplayName
                                UserPrincipalName = $user.UserPrincipalName
                                Mail = $user.Mail
                                AccountEnabled = $user.AccountEnabled
                                UserType = $user.UserType
                            }
                            OnPremisesSyncAttributes = [PSCustomObject]@{
                                OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
                                OnPremisesLastSyncDateTime = $user.OnPremisesLastSyncDateTime
                                OnPremisesDistinguishedName = $user.OnPremisesDistinguishedName
                                OnPremisesDomainName = $user.OnPremisesDomainName
                                OnPremisesSamAccountName = $user.OnPremisesSamAccountName
                                OnPremisesUserPrincipalName = $user.OnPremisesUserPrincipalName
                                OnPremisesSecurityIdentifier = $user.OnPremisesSecurityIdentifier
                                OnPremisesImmutableId = $user.OnPremisesImmutableId
                                ProxyAddresses = $user.ProxyAddresses
                            }
                        }
                        
                        $json = $userBackup | ConvertTo-Json -Depth 10
                        $json | Out-File -FilePath $filePath -Encoding UTF8 -Force
                        $filesCreated += $fileName
                    }
                    
                    # Create index file
                    $indexPath = Join-Path $backupPath "_BackupIndex.json"
                    $indexData = [PSCustomObject]@{
                        BackupDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        TotalUsers = $selectedUsers.Count
                        BackedUpBy = $context.Account
                        Files = $filesCreated
                        Users = $selectedUsers | Select-Object DisplayName, UserPrincipalName, OnPremisesSamAccountName
                    }
                    $indexData | ConvertTo-Json -Depth 5 | Out-File -FilePath $indexPath -Encoding UTF8 -Force
                    
                    Write-DebugLog "Created $($filesCreated.Count) backup files in $backupPath" -Level SUCCESS
                    $StatusBar.Text = "Backup completed: $($filesCreated.Count) files in $backupFolderName"
                    
                    [System.Windows.MessageBox]::Show(
                        "Successfully backed up $($selectedUsers.Count) users!`n`nLocation: $backupPath`nFiles: $($filesCreated.Count)",
                        "Backup Complete",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information
                    )
                    
                    return $true
                }
                else {
                    $StatusBar.Text = "Backup cancelled"
                    return $false
                }
            }
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugLog "Backup failed: $errorMsg" -Level ERROR
        Write-DebugLog "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
        $StatusBar.Text = "Backup failed: $errorMsg"
        
        [System.Windows.MessageBox]::Show(
            "Failed to save backup:`n`n$errorMsg",
            "Backup Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        
        return $false
    }
}

function Clear-OnPremisesAttributes {
    <#
    .SYNOPSIS
    Clears specific on-premises attributes for selected users using Graph API
    
    .PARAMETER SelectedUsers
    Array of selected users to clear attributes for
    
    .PARAMETER StatusBar
    Status bar control to update with progress
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$SelectedUsers,
        
        [Parameter(Mandatory=$true)]
        $StatusBar
    )
    
    try {
        if ($SelectedUsers.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "Please select at least one user to clear on-premises attributes.",
                "No Users Selected",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return $false
        }
        
        # Check ADSyncTools availability
        if (-not $script:ADSyncToolsAvailable) {
            [System.Windows.MessageBox]::Show(
                "The ADSyncTools module is not available on this system.`n`n" +
                "ADSyncTools is required for clearing on-premises attributes.`n" +
                "The script attempted to install it automatically but was unsuccessful.`n`n" +
                "Please try installing manually: Install-Module ADSyncTools -Scope CurrentUser",
                "ADSyncTools Not Available",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return $false
        }
        
        # Confirm action
        $result = [System.Windows.MessageBox]::Show(
            "This will clear the following on-premises attributes for $($SelectedUsers.Count) selected user(s):`n`n" +
            "- OnPremisesDistinguishedName`n" +
            "- OnPremisesDomainName`n" +
            "- OnPremisesSamAccountName`n" +
            "- OnPremisesUserPrincipalName`n" +
            "- OnPremisesSecurityIdentifier`n" +
            "- OnPremisesImmutableId`n`n" +
            "This action cannot be undone. Continue?",
            "Clear On-Premises Attributes",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
            $StatusBar.Text = "Clear operation cancelled"
            return $false
        }
        
        Write-DebugLog "Starting clear on-premises attributes for $($SelectedUsers.Count) users" -Level INFO
        $StatusBar.Text = "Clearing on-premises attributes..."
        
        $successCount = 0
        $failureCount = 0
        $errors = @()
        
        foreach ($user in $SelectedUsers) {
            try {
                Write-DebugLog "Clearing attributes for user: $($user.UserPrincipalName)" -Level INFO
                
                # Use Clear-ADSyncToolsOnPremisesAttribute to clear on-premises attributes
                Clear-ADSyncToolsOnPremisesAttribute -Id $user.Id `
                    -onPremisesDistinguishedName `
                    -onPremisesDomainName `
                    -onPremisesImmutableId `
                    -onPremisesSamAccountName `
                    -onPremisesSecurityIdentifier `
                    -onPremisesUserPrincipalName
                
                Write-DebugLog "Successfully cleared attributes for: $($user.UserPrincipalName)" -Level SUCCESS
                $successCount++
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-DebugLog "Failed to clear attributes for $($user.UserPrincipalName): $errorMsg" -Level ERROR
                $errors += "$($user.UserPrincipalName): $errorMsg"
                $failureCount++
            }
        }
        
        # Show result summary
        $summaryMsg = "Clear operation completed:`n`n"
        $summaryMsg += "Successful: $successCount`n"
        $summaryMsg += "Failed: $failureCount"
        
        if ($errors.Count -gt 0) {
            $summaryMsg += "`n`nErrors:`n" + ($errors -join "`n")
        }
        
        $msgType = if ($failureCount -eq 0) { 
            [System.Windows.MessageBoxImage]::Information 
        } elseif ($successCount -eq 0) { 
            [System.Windows.MessageBoxImage]::Error 
        } else { 
            [System.Windows.MessageBoxImage]::Warning 
        }
        
        [System.Windows.MessageBox]::Show(
            $summaryMsg,
            "Clear Attributes Complete",
            [System.Windows.MessageBoxButton]::OK,
            $msgType
        )
        
        $StatusBar.Text = "Clear completed: $successCount successful, $failureCount failed"
        Write-DebugLog "Clear operation completed: $successCount successful, $failureCount failed" -Level INFO
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugLog "Clear operation failed: $errorMsg" -Level ERROR
        $StatusBar.Text = "Clear operation failed: $errorMsg"
        
        [System.Windows.MessageBox]::Show(
            "Failed to clear on-premises attributes:`n`n$errorMsg",
            "Clear Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        
        return $false
    }
}

function Restore-UserAttributesFromBackup {
    <#
    .SYNOPSIS
    Restores user attributes from a backup JSON file or _BackupIndex
    
    .PARAMETER StatusBar
    Status bar control to update with progress
    #>
    param(
        [Parameter(Mandatory=$true)]
        $StatusBar
    )
    
    try {
        Write-DebugLog "Starting restore from backup" -Level INFO
        
        # Check ADSyncTools availability
        if (-not $script:ADSyncToolsAvailable) {
            [System.Windows.MessageBox]::Show(
                "The ADSyncTools module is not available on this system.`n`n" +
                "ADSyncTools is required for restoring on-premises attributes from backup.`n" +
                "The script attempted to install it automatically but was unsuccessful.`n`n" +
                "Please try installing manually: Install-Module ADSyncTools -Scope CurrentUser",
                "ADSyncTools Not Available",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return $false
        }
        
        # Show file selection dialog
        $openDialog = New-Object Microsoft.Win32.OpenFileDialog
        $openDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $openDialog.Title = "Select Backup File to Restore"
        $openDialog.Multiselect = $false
        
        $dialogResult = $openDialog.ShowDialog()
        
        if ($dialogResult -ne $true) {
            $StatusBar.Text = "Restore cancelled"
            return $false
        }
        
        $backupFile = $openDialog.FileName
        Write-DebugLog "Selected backup file: $backupFile" -Level INFO
        $StatusBar.Text = "Loading backup file..."
        
        # Load and parse JSON
        $backupContent = Get-Content -Path $backupFile -Raw | ConvertFrom-Json
        
        $usersToRestore = @()
        
        # Determine backup file type
        if ($backupContent.PSObject.Properties.Name -contains 'Files' -and 
            $backupContent.PSObject.Properties.Name -contains 'Users') {
            # This is a _BackupIndex file
            Write-DebugLog "Detected _BackupIndex file with $($backupContent.TotalUsers) users" -Level INFO
            
            $backupFolder = Split-Path -Parent $backupFile
            
            foreach ($file in $backupContent.Files) {
                $userFile = Join-Path $backupFolder $file
                if (Test-Path $userFile) {
                    $userData = Get-Content -Path $userFile -Raw | ConvertFrom-Json
                    $usersToRestore += $userData
                }
                else {
                    Write-DebugLog "Warning: Backup file not found: $userFile" -Level WARNING
                }
            }
        }
        elseif ($backupContent.PSObject.Properties.Name -contains 'BackupMetadata' -and 
                $backupContent.PSObject.Properties.Name -contains 'Users') {
            # Single combined backup file with multiple users
            Write-DebugLog "Detected combined backup file with $($backupContent.Users.Count) users" -Level INFO
            $usersToRestore = $backupContent.Users
        }
        elseif ($backupContent.PSObject.Properties.Name -contains 'BackupMetadata' -and 
                $backupContent.PSObject.Properties.Name -contains 'UserData') {
            # Single user backup file
            Write-DebugLog "Detected single user backup file" -Level INFO
            $usersToRestore = @($backupContent)
        }
        else {
            throw "Unrecognized backup file format"
        }
        
        if ($usersToRestore.Count -eq 0) {
            throw "No users found in backup file"
        }
        
        # Confirm restore
        $result = [System.Windows.MessageBox]::Show(
            "Found $($usersToRestore.Count) user(s) in backup file.`n`n" +
            "This will restore the following on-premises attributes:`n" +
            "- OnPremisesDistinguishedName`n" +
            "- OnPremisesDomainName`n" +
            "- OnPremisesSamAccountName`n" +
            "- OnPremisesUserPrincipalName`n" +
            "- OnPremisesSecurityIdentifier`n" +
            "- OnPremisesImmutableId`n" +
            "- ProxyAddresses`n`n" +
            "Continue with restore?",
            "Restore User Attributes",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
            $StatusBar.Text = "Restore cancelled"
            return $false
        }
        
        $StatusBar.Text = "Restoring user attributes..."
        $successCount = 0
        $failureCount = 0
        $errors = @()
        
        foreach ($userBackup in $usersToRestore) {
            try {
                $userPrincipalName = $userBackup.UserData.UserPrincipalName
                Write-DebugLog "Restoring attributes for: $userPrincipalName" -Level INFO
                
                # Find user in Graph
                $graphUser = Get-MgUser -Filter "userPrincipalName eq '$userPrincipalName'" -ErrorAction Stop
                
                if (-not $graphUser) {
                    throw "User not found in Entra ID: $userPrincipalName"
                }
                
                # Restore on-premises attributes from backup using Set-ADSyncToolsOnPremisesAttribute
                $syncAttrs = $userBackup.OnPremisesSyncAttributes
                Set-ADSyncToolsOnPremisesAttribute -Id $graphUser.Id `
                    -onPremisesDistinguishedName $syncAttrs.OnPremisesDistinguishedName `
                    -onPremisesDomainName $syncAttrs.OnPremisesDomainName `
                    -onPremisesImmutableId $syncAttrs.OnPremisesImmutableId `
                    -onPremisesSamAccountName $syncAttrs.OnPremisesSamAccountName `
                    -onPremisesSecurityIdentifier $syncAttrs.OnPremisesSecurityIdentifier `
                    -onPremisesUserPrincipalName $syncAttrs.OnPremisesUserPrincipalName
                
                Write-DebugLog "Successfully restored attributes for: $userPrincipalName" -Level SUCCESS
                $successCount++
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-DebugLog "Failed to restore $($userPrincipalName): $errorMsg" -Level ERROR
                $errors += "$($userPrincipalName): $errorMsg"
                $failureCount++
            }
        }
        
        # Show result summary
        $summaryMsg = "Restore operation completed:`n`n"
        $summaryMsg += "Successful: $successCount`n"
        $summaryMsg += "Failed: $failureCount"
        
        if ($errors.Count -gt 0) {
            $summaryMsg += "`n`nErrors:`n" + ($errors -join "`n")
        }
        
        $msgType = if ($failureCount -eq 0) { 
            [System.Windows.MessageBoxImage]::Information 
        } elseif ($successCount -eq 0) { 
            [System.Windows.MessageBoxImage]::Error 
        } else { 
            [System.Windows.MessageBoxImage]::Warning 
        }
        
        [System.Windows.MessageBox]::Show(
            $summaryMsg,
            "Restore Complete",
            [System.Windows.MessageBoxButton]::OK,
            $msgType
        )
        
        $StatusBar.Text = "Restore completed: $successCount successful, $failureCount failed"
        Write-DebugLog "Restore operation completed: $successCount successful, $failureCount failed" -Level INFO
        
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugLog "Restore operation failed: $errorMsg" -Level ERROR
        Write-DebugLog "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
        $StatusBar.Text = "Restore operation failed: $errorMsg"
        
        [System.Windows.MessageBox]::Show(
            "Failed to restore from backup:`n`n$errorMsg",
            "Restore Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        
        return $false
    }
}

function Switch-UserSOA {
    <#
    .SYNOPSIS
    Switches the Source of Authority (SOA) for selected users from On-Premises to Cloud
    This function performs a two-step process:
    1. Sets isCloudManaged = true via the onPremisesSyncBehavior endpoint (v1.0 API)
    2. Clears on-premises sync attributes after verifying the SOA switch
    
    .PARAMETER SelectedUsers
    Array of selected users whose SOA will be switched
    
    .PARAMETER StatusBar
    Status bar control to update with progress
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$SelectedUsers,
        
        [Parameter(Mandatory=$true)]
        $StatusBar
    )
    
    try {
        if ($SelectedUsers.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "Please select at least one user to switch Source of Authority.",
                "No Users Selected",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return $false
        }
        
        # Strong warning before SOA switch
        $result = [System.Windows.MessageBox]::Show(
            "⚠ CRITICAL WARNING ⚠`n`n" +
            "You are about to switch the Source of Authority (SOA) for $($SelectedUsers.Count) user(s) from On-Premises to Cloud.`n`n" +
            "CRITICAL SEQUENCE:`n" +
            "1. Set isCloudManaged = true (SOA switch to cloud)`n" +
            "2. Verify SOA switch succeeded`n" +
            "3. Clear on-premises sync attributes`n`n" +
            "PREREQUISITES:`n" +
            "✓ Users should be excluded from on-premises AD sync scope`n" +
            "✓ A backup of user attributes should exist before proceeding`n" +
            "✓ Ensure groups are cloud-managed BEFORE switching users`n`n" +
            "This operation will use the Microsoft Graph v1.0 endpoint.`n`n" +
            "Are you sure you want to proceed?",
            "Switch Source of Authority",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
            $StatusBar.Text = "SOA switch cancelled"
            return $false
        }
        
        Write-DebugLog "========================================" -Level INFO
        Write-DebugLog "Starting SOA switch process for $($SelectedUsers.Count) user(s)" -Level INFO
        Write-DebugLog "========================================" -Level INFO
        
        $totalUsers = $SelectedUsers.Count
        $soaSwitchSuccess = @()
        $soaSwitchFailed = @()
        $attributeClearSuccess = @()
        $attributeClearFailed = @()
        $errors = @()
        
        # === PHASE 1: Switch SOA to Cloud (isCloudManaged = true) ===
        $StatusBar.Text = "Phase 1/2: Switching SOA to cloud for $totalUsers user(s)..."
        Write-DebugLog "PHASE 1: Switching Source of Authority to Cloud" -Level INFO
        
        $currentUser = 0
        foreach ($user in $SelectedUsers) {
            $currentUser++
            try {
                Write-DebugLog "[$currentUser/$totalUsers] Switching SOA for: $($user.UserPrincipalName)" -Level INFO
                $StatusBar.Text = "Phase 1/2: Switching SOA [$currentUser/$totalUsers] - $($user.UserPrincipalName)"
                
                # Use v1.0 endpoint with onPremisesSyncBehavior
                $updateBody = @{
                    isCloudManaged = $true
                }
                
                $uri = "https://graph.microsoft.com/v1.0/users/$($user.Id)/onPremisesSyncBehavior"
                Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $updateBody -ContentType "application/json"
                
                Write-DebugLog "✓ Successfully set isCloudManaged=true for: $($user.UserPrincipalName)" -Level SUCCESS
                $soaSwitchSuccess += $user
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-DebugLog "✗ Failed to switch SOA for $($user.UserPrincipalName): $errorMsg" -Level ERROR
                $soaSwitchFailed += $user
                $errors += "SOA Switch - $($user.UserPrincipalName): $errorMsg"
            }
        }
        
        Write-DebugLog "PHASE 1 Complete: $($soaSwitchSuccess.Count) successful, $($soaSwitchFailed.Count) failed" -Level INFO
        
        # === PHASE 2: Clear On-Premises Attributes (only for successful SOA switches) ===
        if ($soaSwitchSuccess.Count -gt 0) {
            $StatusBar.Text = "Phase 2/2: Clearing attributes for $($soaSwitchSuccess.Count) cloud-managed user(s)..."
            Write-DebugLog "PHASE 2: Clearing on-premises attributes for successfully switched users" -Level INFO
            
            # Small delay to allow SOA change to propagate
            Start-Sleep -Milliseconds 500
            
            $currentUser = 0
            foreach ($user in $soaSwitchSuccess) {
                $currentUser++
                try {
                    Write-DebugLog "[$currentUser/$($soaSwitchSuccess.Count)] Clearing attributes for: $($user.UserPrincipalName)" -Level INFO
                    $StatusBar.Text = "Phase 2/2: Clearing attributes [$currentUser/$($soaSwitchSuccess.Count)] - $($user.UserPrincipalName)"
                    
                    # Verify SOA switch by checking current state
                    $verifyUri = "https://graph.microsoft.com/v1.0/users/$($user.Id)/onPremisesSyncBehavior"
                    try {
                        $syncBehavior = Invoke-MgGraphRequest -Method GET -Uri $verifyUri
                        if ($syncBehavior.isCloudManaged -ne $true) {
                            throw "SOA verification failed: isCloudManaged is not true"
                        }
                        Write-DebugLog "  Verified: isCloudManaged = true" -Level INFO
                    }
                    catch {
                        Write-DebugLog "  Warning: Could not verify SOA state, proceeding with caution: $($_.Exception.Message)" -Level WARNING
                    }
                    
                    # Clear on-premises sync attributes using Clear-ADSyncToolsOnPremisesAttribute
                    Clear-ADSyncToolsOnPremisesAttribute -Id $user.Id `
                        -onPremisesDistinguishedName `
                        -onPremisesDomainName `
                        -onPremisesImmutableId `
                        -onPremisesSamAccountName `
                        -onPremisesSecurityIdentifier `
                        -onPremisesUserPrincipalName
                    
                    Write-DebugLog "✓ Successfully cleared on-premises attributes for: $($user.UserPrincipalName)" -Level SUCCESS
                    $attributeClearSuccess += $user
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-DebugLog "✗ Failed to clear attributes for $($user.UserPrincipalName): $errorMsg" -Level ERROR
                    $attributeClearFailed += $user
                    $errors += "Attribute Clear - $($user.UserPrincipalName): $errorMsg"
                }
            }
            
            Write-DebugLog "PHASE 2 Complete: $($attributeClearSuccess.Count) successful, $($attributeClearFailed.Count) failed" -Level INFO
        }
        else {
            Write-DebugLog "PHASE 2 Skipped: No users successfully switched SOA" -Level WARNING
        }
        
        # === FINAL SUMMARY ===
        Write-DebugLog "========================================" -Level INFO
        Write-DebugLog "SOA Switch Process Complete" -Level INFO
        Write-DebugLog "Phase 1 (SOA Switch): $($soaSwitchSuccess.Count) successful, $($soaSwitchFailed.Count) failed" -Level INFO
        Write-DebugLog "Phase 2 (Attribute Clear): $($attributeClearSuccess.Count) successful, $($attributeClearFailed.Count) failed" -Level INFO
        Write-DebugLog "========================================" -Level INFO
        
        $summaryMsg = "SOA Switch Process Completed`n`n"
        $summaryMsg += "=== PHASE 1: Switch SOA to Cloud ===`n"
        $summaryMsg += "Successful: $($soaSwitchSuccess.Count)`n"
        $summaryMsg += "Failed: $($soaSwitchFailed.Count)`n`n"
        
        if ($soaSwitchSuccess.Count -gt 0) {
            $summaryMsg += "=== PHASE 2: Clear On-Premises Attributes ===`n"
            $summaryMsg += "Successful: $($attributeClearSuccess.Count)`n"
            $summaryMsg += "Failed: $($attributeClearFailed.Count)`n`n"
        }
        
        $summaryMsg += "=== OVERALL ===`n"
        $summaryMsg += "Fully Completed: $($attributeClearSuccess.Count) user(s)`n"
        $summaryMsg += "Partial/Failed: $(($soaSwitchFailed.Count + $attributeClearFailed.Count)) user(s)"
        
        if ($errors.Count -gt 0) {
            $summaryMsg += "`n`n=== ERRORS ===`n" + ($errors -join "`n")
        }
        
        $msgType = if ($errors.Count -eq 0) { 
            [System.Windows.MessageBoxImage]::Information 
        } elseif ($attributeClearSuccess.Count -eq 0) { 
            [System.Windows.MessageBoxImage]::Error 
        } else { 
            [System.Windows.MessageBoxImage]::Warning 
        }
        
        [System.Windows.MessageBox]::Show(
            $summaryMsg,
            "SOA Switch Process Complete",
            [System.Windows.MessageBoxButton]::OK,
            $msgType
        )
        
        $StatusBar.Text = "SOA switch completed: $($attributeClearSuccess.Count) fully completed, $($errors.Count) errors"
        
        return $attributeClearSuccess.Count -gt 0
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugLog "CRITICAL ERROR in SOA switch process: $errorMsg" -Level ERROR
        Write-DebugLog "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
        $StatusBar.Text = "SOA switch process failed: $errorMsg"
        
        [System.Windows.MessageBox]::Show(
            "Failed to complete SOA switch process:`n`n$errorMsg",
            "SOA Switch Process Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        
        return $false
    }
}

function Rollback-UserSOA {
    <#
    .SYNOPSIS
    Rolls back the Source of Authority (SOA) for selected users from Cloud to On-Premises
    This function sets isCloudManaged = false via the onPremisesSyncBehavior endpoint (v1.0 API)
    
    .PARAMETER SelectedUsers
    Array of selected users whose SOA will be rolled back
    
    .PARAMETER StatusBar
    Status bar control to update with progress
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$SelectedUsers,
        
        [Parameter(Mandatory=$true)]
        $StatusBar
    )
    
    try {
        if ($SelectedUsers.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "Please select at least one user to rollback Source of Authority.",
                "No Users Selected",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return $false
        }
        
        # Strong warning before SOA rollback
        $result = [System.Windows.MessageBox]::Show(
            "⚠ CRITICAL WARNING ⚠`n`n" +
            "You are about to ROLLBACK the Source of Authority (SOA) for $($SelectedUsers.Count) user(s) from Cloud to On-Premises.`n`n" +
            "WHAT THIS DOES:`n" +
            "• Sets isCloudManaged = false (SOA reverts to on-premises)`n" +
            "• Users will be subject to on-premises AD sync again`n" +
            "• Does NOT restore on-premises attributes (use Restore from Backup separately)`n`n" +
            "PREREQUISITES:`n" +
            "✓ Users must currently be cloud-managed (isCloudManaged = true)`n" +
            "✓ Understand implications of returning to on-premises management`n" +
            "✓ Ensure on-premises AD is configured to sync these users`n`n" +
            "USE CASE:`n" +
            "This rollback is typically used when SOA switch succeeded but attribute clearing failed,`n" +
            "or when you need to revert a cloud SOA decision.`n`n" +
            "This operation will use the Microsoft Graph v1.0 endpoint.`n`n" +
            "Are you sure you want to proceed?",
            "Rollback Source of Authority",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
            $StatusBar.Text = "SOA rollback cancelled"
            return $false
        }
        
        Write-DebugLog "========================================" -Level INFO
        Write-DebugLog "Starting SOA rollback process for $($SelectedUsers.Count) user(s)" -Level INFO
        Write-DebugLog "========================================" -Level INFO
        
        $totalUsers = $SelectedUsers.Count
        $rollbackSuccess = @()
        $rollbackFailed = @()
        $skippedUsers = @()
        $errors = @()
        
        $StatusBar.Text = "Rolling back SOA to on-premises for $totalUsers user(s)..."
        
        $currentUser = 0
        foreach ($user in $SelectedUsers) {
            $currentUser++
            try {
                Write-DebugLog "[$currentUser/$totalUsers] Rolling back SOA for: $($user.UserPrincipalName)" -Level INFO
                $StatusBar.Text = "Rolling back SOA [$currentUser/$totalUsers] - $($user.UserPrincipalName)"
                
                # Verify current SOA state - user must be cloud-managed to rollback
                $verifyUri = "https://graph.microsoft.com/v1.0/users/$($user.Id)/onPremisesSyncBehavior"
                try {
                    $syncBehavior = Invoke-MgGraphRequest -Method GET -Uri $verifyUri
                    
                    if ($syncBehavior.isCloudManaged -ne $true) {
                        Write-DebugLog "⊘ Skipping $($user.UserPrincipalName) - already on-premises managed (isCloudManaged = false)" -Level WARNING
                        $skippedUsers += $user
                        continue
                    }
                    
                    Write-DebugLog "  Verified: Currently cloud-managed (isCloudManaged = true)" -Level INFO
                }
                catch {
                    Write-DebugLog "  Warning: Could not verify SOA state: $($_.Exception.Message)" -Level WARNING
                    # Continue with rollback attempt even if verification fails
                }
                
                # Rollback SOA by setting isCloudManaged = false
                $updateBody = @{
                    isCloudManaged = $false
                }
                
                $uri = "https://graph.microsoft.com/v1.0/users/$($user.Id)/onPremisesSyncBehavior"
                Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $updateBody -ContentType "application/json"
                
                Write-DebugLog "✓ Successfully set isCloudManaged=false for: $($user.UserPrincipalName)" -Level SUCCESS
                $rollbackSuccess += $user
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-DebugLog "✗ Failed to rollback SOA for $($user.UserPrincipalName): $errorMsg" -Level ERROR
                $rollbackFailed += $user
                $errors += "$($user.UserPrincipalName): $errorMsg"
            }
        }
        
        # === FINAL SUMMARY ===
        Write-DebugLog "========================================" -Level INFO
        Write-DebugLog "SOA Rollback Process Complete" -Level INFO
        Write-DebugLog "Successful: $($rollbackSuccess.Count)" -Level INFO
        Write-DebugLog "Failed: $($rollbackFailed.Count)" -Level INFO
        Write-DebugLog "Skipped (already on-prem): $($skippedUsers.Count)" -Level INFO
        Write-DebugLog "========================================" -Level INFO
        
        $summaryMsg = "SOA Rollback Process Completed`n`n"
        $summaryMsg += "Successful: $($rollbackSuccess.Count)`n"
        $summaryMsg += "Failed: $($rollbackFailed.Count)`n"
        $summaryMsg += "Skipped (already on-premises): $($skippedUsers.Count)"
        
        if ($skippedUsers.Count -gt 0) {
            $summaryMsg += "`n`nSkipped Users:`n" + (($skippedUsers | ForEach-Object { $_.UserPrincipalName }) -join "`n")
        }
        
        if ($errors.Count -gt 0) {
            $summaryMsg += "`n`n=== ERRORS ===`n" + ($errors -join "`n")
        }
        
        $msgType = if ($errors.Count -eq 0 -and $skippedUsers.Count -eq 0) { 
            [System.Windows.MessageBoxImage]::Information 
        } elseif ($rollbackSuccess.Count -eq 0) { 
            [System.Windows.MessageBoxImage]::Error 
        } else { 
            [System.Windows.MessageBoxImage]::Warning 
        }
        
        [System.Windows.MessageBox]::Show(
            $summaryMsg,
            "SOA Rollback Process Complete",
            [System.Windows.MessageBoxButton]::OK,
            $msgType
        )
        
        $StatusBar.Text = "SOA rollback completed: $($rollbackSuccess.Count) successful, $($rollbackFailed.Count) failed, $($skippedUsers.Count) skipped"
        
        return $rollbackSuccess.Count -gt 0
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DebugLog "CRITICAL ERROR in SOA rollback process: $errorMsg" -Level ERROR
        Write-DebugLog "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
        $StatusBar.Text = "SOA rollback process failed: $errorMsg"
        
        [System.Windows.MessageBox]::Show(
            "Failed to complete SOA rollback process:`n`n$errorMsg",
            "SOA Rollback Process Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        
        return $false
    }
}

# Main script execution
try {
    # Initialize debug logging
    Write-DebugLog "=================================================" -Level INFO
    Write-DebugLog "Entra ID User Sync Attribute Backup Tool Started" -Level INFO
    Write-DebugLog "PowerShell Version: $($PSVersionTable.PSVersion)" -Level INFO
    Write-DebugLog "Script Path: $PSScriptRoot" -Level INFO
    Write-DebugLog "Log File: $script:LogFile" -Level INFO
    Write-DebugLog "=================================================" -Level INFO
    
    # Initialize Graph modules
    Write-Host "\n=== Entra ID User Sync Attribute Backup Tool ===" -ForegroundColor Cyan
    Write-Host "Initializing..." -ForegroundColor Cyan
    Write-DebugLog "Starting application initialization" -Level INFO
    
    if (-not (Initialize-GraphModule)) {
        throw "Failed to initialize Microsoft Graph modules"
    }
    
    # Initialize ADSyncTools module (non-fatal if not available)
    Initialize-ADSyncToolsModule | Out-Null
    
    # Load XAML
    $xamlPath = Join-Path $PSScriptRoot "UserBackupUI.xaml"
    Write-DebugLog "XAML Path: $xamlPath" -Level INFO
    
    if (-not (Test-Path $xamlPath)) {
        Write-DebugLog "XAML file not found at path: $xamlPath" -Level ERROR
        throw "XAML file not found: $xamlPath"
    }
    
    Write-DebugLog "Loading XAML file..." -Level INFO
    [xml]$xaml = Get-Content $xamlPath
    Write-DebugLog "XAML file loaded successfully" -Level SUCCESS
    
    Write-DebugLog "Creating XML reader..." -Level INFO
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    Write-DebugLog "Loading WPF window from XAML..." -Level INFO
    $window = [Windows.Markup.XamlReader]::Load($reader)
    Write-DebugLog "WPF window loaded successfully" -Level SUCCESS
    
    # Get UI elements
    Write-DebugLog "Binding UI elements..." -Level INFO
    $btnConnect = $window.FindName("btnConnect")
    $btnDisconnect = $window.FindName("btnDisconnect")
    $btnLoadUsers = $window.FindName("btnLoadUsers")
    $btnLoadCloudUsers = $window.FindName("btnLoadCloudUsers")
    $lblConnectionStatus = $window.FindName("lblConnectionStatus")
    $lblConnectedUser = $window.FindName("lblConnectedUser")
    $txtFilterUsers = $window.FindName("txtFilterUsers")
    $btnClearFilter = $window.FindName("btnClearFilter")
    $txtTotalCount = $window.FindName("txtTotalCount")
    $txtFilteredCount = $window.FindName("txtFilteredCount")
    $txtSelectedCount = $window.FindName("txtSelectedCount")
    $btnSelectAll = $window.FindName("btnSelectAll")
    $btnSelectNone = $window.FindName("btnSelectNone")
    $dgUsers = $window.FindName("dgUsers")
    $pnlProgress = $window.FindName("pnlProgress")
    $txtProgressStatus = $window.FindName("txtProgressStatus")
    $pbProgress = $window.FindName("pbProgress")
    $txtProgressDetail = $window.FindName("txtProgressDetail")
    $btnBackup = $window.FindName("btnBackup")
    $btnClose = $window.FindName("btnClose")
    $txtStatusBar = $window.FindName("txtStatusBar")
    $script:btnClearAttributes = $window.FindName("btnClearAttributes")
    $script:btnSwitchSOA = $window.FindName("btnSwitchSOA")
    $script:btnRollbackSOA = $window.FindName("btnRollbackSOA")
    $script:btnRestoreBackup = $window.FindName("btnRestoreBackup")
    
    Write-DebugLog "All UI elements bound successfully" -Level SUCCESS
    
    # Connect button click event
    Write-DebugLog "Attaching event handlers..." -Level INFO
    $btnConnect.Add_Click({
        Write-DebugLog "Connect button clicked" -Level INFO
        if (Connect-ToGraph -StatusLabel $lblConnectionStatus -UserLabel $lblConnectedUser -StatusBar $txtStatusBar) {
            $btnLoadUsers.IsEnabled = $true
            $btnLoadCloudUsers.IsEnabled = $true
            $btnDisconnect.IsEnabled = $true
            $script:btnRestoreBackup.IsEnabled = $true
            $txtStatusBar.Text = "Connected to Graph. Click 'Load Synced Users' or 'Load Cloud Users' to begin."
        }
    })
    
    # Disconnect/Switch User button click event
    $btnDisconnect.Add_Click({
        Write-DebugLog "Disconnect/Switch User button clicked" -Level INFO
        
        $result = [System.Windows.MessageBox]::Show(
            "This will disconnect from Microsoft Graph and clear all loaded data. Do you want to continue?",
            "Switch User",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
                Write-DebugLog "Disconnecting from Microsoft Graph" -Level INFO
                Disconnect-MgGraph | Out-Null
                Write-DebugLog "Disconnected successfully" -Level SUCCESS
                
                # Reset UI state
                $script:IsConnected = $false
                $script:IsUsersLoaded = $false
                $script:AllUsers = @()
                
                # Clear data grid
                if ($null -ne $script:UserCollection) {
                    $script:UserCollection.Clear()
                }
                
                # Reset UI controls
                $lblConnectionStatus.Content = "Not Connected"
                $lblConnectionStatus.Foreground = "Red"
                $lblConnectedUser.Content = "N/A"
                $lblConnectedUser.Foreground = "Gray"
                $btnLoadUsers.IsEnabled = $false
                $btnLoadCloudUsers.IsEnabled = $false
                $btnDisconnect.IsEnabled = $false
                $btnBackup.IsEnabled = $false
                $txtFilterUsers.IsEnabled = $false
                $btnClearFilter.IsEnabled = $false
                $btnSelectAll.IsEnabled = $false
                $btnSelectNone.IsEnabled = $false
                $script:btnClearAttributes.IsEnabled = $false
                $script:btnSwitchSOA.IsEnabled = $false
                $script:btnRollbackSOA.IsEnabled = $false
                $script:btnRestoreBackup.IsEnabled = $false
                $txtTotalCount.Text = "0"
                $txtFilteredCount.Text = "0"
                $txtSelectedCount.Text = "0"
                $btnBackup.Content = "Backup Selected Users (0)"
                $txtFilterUsers.Text = ""
                $script:CurrentUserType = $null
                Update-UIForUserType -UserType 'Default'
                
                $txtStatusBar.Text = "Disconnected. Click 'Connect to Graph' to sign in with a different account."
                
                Write-Host "Disconnected successfully. You can now connect with a different account." -ForegroundColor Green
                Write-DebugLog "UI reset completed" -Level SUCCESS
                
            } catch {
                Write-Host "Error during disconnect: $_" -ForegroundColor Red
                Write-DebugLog "Error during disconnect: $_" -Level ERROR
                [System.Windows.MessageBox]::Show(
                    "Failed to disconnect: $_",
                    "Disconnect Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            }
        }
    })
    
    # Load Users button click event
    $btnLoadUsers.Add_Click({
        Write-DebugLog "Load Users button clicked" -Level INFO
        
        if (-not $script:IsConnected) {
            [System.Windows.MessageBox]::Show(
                "Please connect to Microsoft Graph first.",
                "Not Connected",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }
        
        # Disable button and show progress
        $btnLoadUsers.IsEnabled = $false
        $pnlProgress.Visibility = "Visible"
        
        try {
            # Load all synced users
            $users = Load-AllSyncedUsers -ProgressBar $pbProgress `
                                         -ProgressStatus $txtProgressStatus `
                                         -ProgressDetail $txtProgressDetail `
                                         -StatusBar $txtStatusBar
            
            if ($users.Count -gt 0) {
                $script:AllUsers = $users
                $script:IsUsersLoaded = $true
                
                # Update DataGrid
                Update-UserGrid -DataGrid $dgUsers -Users $script:AllUsers
                
                # Update counts
                $txtTotalCount.Text = $script:AllUsers.Count
                $txtFilteredCount.Text = $script:AllUsers.Count
                $txtSelectedCount.Text = "0"
                
                # Enable controls
                $txtFilterUsers.IsEnabled = $true
                $btnClearFilter.IsEnabled = $true
                $btnSelectAll.IsEnabled = $true
                $btnSelectNone.IsEnabled = $true
                $script:btnRestoreBackup.IsEnabled = $true
                
                # Update UI for on-premises synced user type
                Update-UIForUserType -UserType 'OnPrem'
                
                Write-DebugLog "User load completed. $($users.Count) users loaded" -Level SUCCESS
            }
        }
        finally {
            # Hide progress and re-enable buttons
            $pnlProgress.Visibility = "Collapsed"
            $btnLoadUsers.IsEnabled = $true
            $btnLoadCloudUsers.IsEnabled = $true
        }
    })
    
    # Load Cloud Users button click event
    $btnLoadCloudUsers.Add_Click({
        Write-DebugLog "Load Cloud Users button clicked" -Level INFO
        
        if (-not $script:IsConnected) {
            [System.Windows.MessageBox]::Show(
                "Please connect to Microsoft Graph first.",
                "Not Connected",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }
        
        # Disable buttons and show progress
        $btnLoadCloudUsers.IsEnabled = $false
        $btnLoadUsers.IsEnabled = $false
        $pnlProgress.Visibility = "Visible"
        $txtProgressStatus.Text = "Loading cloud-managed users..."
        
        try {
            # Load all cloud-managed (non-synced) users
            $users = Load-AllCloudUsers -ProgressBar $pbProgress `
                                        -ProgressStatus $txtProgressStatus `
                                        -ProgressDetail $txtProgressDetail `
                                        -StatusBar $txtStatusBar
            
            if ($users.Count -gt 0) {
                $script:AllUsers = $users
                $script:IsUsersLoaded = $true
                
                # Update DataGrid
                Update-UserGrid -DataGrid $dgUsers -Users $script:AllUsers
                
                # Update counts
                $txtTotalCount.Text = $script:AllUsers.Count
                $txtFilteredCount.Text = $script:AllUsers.Count
                $txtSelectedCount.Text = "0"
                
                # Enable controls
                $txtFilterUsers.IsEnabled = $true
                $btnClearFilter.IsEnabled = $true
                $btnSelectAll.IsEnabled = $true
                $btnSelectNone.IsEnabled = $true
                $script:btnRestoreBackup.IsEnabled = $true
                
                # Update UI for cloud-managed user type
                Update-UIForUserType -UserType 'Cloud'
                
                Write-DebugLog "Cloud user load completed. $($users.Count) users loaded" -Level SUCCESS
            }
        }
        finally {
            # Hide progress and re-enable buttons
            $pnlProgress.Visibility = "Collapsed"
            $btnLoadCloudUsers.IsEnabled = $true
            $btnLoadUsers.IsEnabled = $true
        }
    })
    
    # Filter TextBox text changed event
    $txtFilterUsers.Add_TextChanged({
        $filterText = $txtFilterUsers.Text
        
        if ([string]::IsNullOrWhiteSpace($filterText) -or $filterText.Length -lt 3) {
            # Show all users if filter is too short
            if ($script:AllUsers.Count -gt 0) {
                Update-UserGrid -DataGrid $dgUsers -Users $script:AllUsers
                $txtFilteredCount.Text = $script:AllUsers.Count
            }
        }
        else {
            # Apply filter
            Apply-UserFilter -FilterText $filterText `
                             -DataGrid $dgUsers `
                             -FilteredCountLabel $txtFilteredCount
        }
        
        # Update counts after filter
        Update-SelectionCounts -TotalCountLabel $txtTotalCount `
                               -FilteredCountLabel $txtFilteredCount `
                               -SelectedCountLabel $txtSelectedCount `
                               -BackupButton $btnBackup
    })
    
    # Clear Filter button click event
    $btnClearFilter.Add_Click({
        $txtFilterUsers.Text = ""
        Write-DebugLog "Filter cleared" -Level INFO
    })
    
    # Select All button click event
    $btnSelectAll.Add_Click({
        Write-DebugLog "Select All button clicked" -Level INFO
        
        # Select all visible users in the current DataGrid view
        if ($script:UserCollection) {
            foreach ($user in $script:UserCollection) {
                $user.IsSelected = $true
            }
            
            # Refresh DataGrid
            $dgUsers.Items.Refresh()
            
            # Update counts
            Update-SelectionCounts -TotalCountLabel $txtTotalCount `
                                   -FilteredCountLabel $txtFilteredCount `
                                   -SelectedCountLabel $txtSelectedCount `
                                   -BackupButton $btnBackup
            
            Write-DebugLog "Selected all visible users" -Level SUCCESS
        }
    })
    
    # Select None button click event
    $btnSelectNone.Add_Click({
        Write-DebugLog "Clear All Selections button clicked" -Level INFO
        
        # Deselect all users (including filtered out ones)
        foreach ($user in $script:AllUsers) {
            $user.IsSelected = $false
        }
        
        # Refresh DataGrid
        $dgUsers.Items.Refresh()
        
        # Update counts
        Update-SelectionCounts -TotalCountLabel $txtTotalCount `
                               -FilteredCountLabel $txtFilteredCount `
                               -SelectedCountLabel $txtSelectedCount `
                               -BackupButton $btnBackup
        
        Write-DebugLog "Cleared all selections" -Level SUCCESS
    })
    
# DataGrid TargetUpdated event - fires when checkbox binding updates
$dgUsers.AddHandler(
    [System.Windows.Data.Binding]::TargetUpdatedEvent,
    [System.Windows.RoutedEventHandler]{
        param($sender, $e)
        # Update counts immediately when checkbox value changes
        Update-SelectionCounts -TotalCountLabel $txtTotalCount `
                               -FilteredCountLabel $txtFilteredCount `
                               -SelectedCountLabel $txtSelectedCount `
                               -BackupButton $btnBackup
    }
)

# Also handle PreviewMouseUp on DataGrid to catch checkbox clicks
$dgUsers.Add_PreviewMouseUp({
    param($sender, $e)
    # Small delay to let the click complete and binding update
    $dgUsers.Dispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [System.Action]{
            Update-SelectionCounts -TotalCountLabel $txtTotalCount `
                                   -FilteredCountLabel $txtFilteredCount `
                                   -SelectedCountLabel $txtSelectedCount `
                                   -BackupButton $btnBackup
        }
    )
})    
    # Backup button click event
    $btnBackup.Add_Click({
        Write-DebugLog "Backup button clicked" -Level INFO
        Backup-SelectedUsersToJson -StatusBar $txtStatusBar
    })
    
    # Clear On-Premises Attributes button click event
    $script:btnClearAttributes.Add_Click({
        Write-DebugLog "Clear On-Premises Attributes button clicked" -Level INFO
        $selectedUsers = Get-SelectedUsers
        Clear-OnPremisesAttributes -SelectedUsers $selectedUsers -StatusBar $txtStatusBar
    })
    
    # Switch SOA to Cloud button click event
    $script:btnSwitchSOA.Add_Click({
        Write-DebugLog "Switch SOA button clicked" -Level INFO
        $selectedUsers = Get-SelectedUsers
        Switch-UserSOA -SelectedUsers $selectedUsers -StatusBar $txtStatusBar
    })
    
    # Rollback SOA to On-Premises button click event
    $script:btnRollbackSOA.Add_Click({
        Write-DebugLog "Rollback SOA button clicked" -Level INFO
        $selectedUsers = Get-SelectedUsers
        Rollback-UserSOA -SelectedUsers $selectedUsers -StatusBar $txtStatusBar
    })
    
    # Restore from Backup button click event
    $script:btnRestoreBackup.Add_Click({
        Write-DebugLog "Restore from Backup button clicked" -Level INFO
        Restore-UserAttributesFromBackup -StatusBar $txtStatusBar
    })
    
    # Close button click event
    $btnClose.Add_Click({
        Write-DebugLog "Close button clicked" -Level INFO
        $window.Close()
    })
    
    Write-DebugLog "All event handlers attached successfully" -Level SUCCESS
    
    # Show window
    Write-Host "Launching UI..." -ForegroundColor Green
    Write-DebugLog "Displaying WPF window..." -Level INFO
    $window.ShowDialog() | Out-Null
    Write-DebugLog "Window closed by user" -Level INFO
    
    # Cleanup on exit
    if ($script:IsConnected) {
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
        Write-DebugLog "Disconnecting from Microsoft Graph" -Level INFO
        Disconnect-MgGraph | Out-Null
        Write-DebugLog "Disconnected successfully" -Level SUCCESS
    }
    
    Write-Host "Application closed." -ForegroundColor Cyan
    Write-DebugLog "Application terminated normally" -Level SUCCESS
    Write-DebugLog "Log file saved at: $script:LogFile" -Level INFO
}
catch {
    $errorMsg = $_.Exception.Message
    Write-DebugLog "FATAL ERROR: $errorMsg" -Level ERROR
    Write-DebugLog "Exception Type: $($_.Exception.GetType().FullName)" -Level ERROR
    Write-DebugLog "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
    Write-DebugLog "Position: Line $($_.InvocationInfo.ScriptLineNumber), Column $($_.InvocationInfo.OffsetInLine)" -Level ERROR
    
    Write-Host "Fatal error: $($_.Exception.Message)" -ForegroundColor Red
    [System.Windows.MessageBox]::Show(
        "Application error:`n`n$($_.Exception.Message)",
        "Error",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
    
    Write-DebugLog "Log file saved at: $script:LogFile" -Level ERROR
}
