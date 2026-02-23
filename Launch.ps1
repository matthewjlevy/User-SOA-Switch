#Requires -Version 7.0

<#
.SYNOPSIS
    Launcher for Entra ID User Sync Attribute Backup Tool
.DESCRIPTION
    Launches the main application in a fresh PowerShell process to avoid module conflicts
#>

Write-Host "Launching Entra ID User Sync Attribute Backup Tool..." -ForegroundColor Cyan
Write-Host "Starting in new PowerShell window to avoid module version conflicts..." -ForegroundColor Yellow

$scriptPath = Join-Path $PSScriptRoot "User-SOA-Switch.ps1"

# Launch in a completely new PowerShell process
$pinfo = New-Object System.Diagnostics.ProcessStartInfo
$pinfo.FileName = "pwsh"
$pinfo.Arguments = "-NoProfile -File `"$scriptPath`""
$pinfo.UseShellExecute = $false
$pinfo.CreateNoWindow = $false

$process = New-Object System.Diagnostics.Process
$process.StartInfo = $pinfo
$process.Start() | Out-Null

Write-Host "Application launched in new window (PID: $($process.Id))" -ForegroundColor Green
Write-Host "You can close this window." -ForegroundColor Gray
