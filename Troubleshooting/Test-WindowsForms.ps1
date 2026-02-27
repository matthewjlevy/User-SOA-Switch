# Test if Windows Forms is accessible
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    [System.Windows.Forms.Application]::DoEvents()
    Write-Host "✓ Windows Forms available" -ForegroundColor Green
}
catch {
    Write-Host "✗ Windows Forms NOT available: $_" -ForegroundColor Red
}

Write-Host "`nPowerShell: $($PSVersionTable.PSVersion)"
Write-Host ".NET: $([System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)"