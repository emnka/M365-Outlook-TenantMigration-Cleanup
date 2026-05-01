# Outlook Tenant-to-Tenant Migration Cleanup Script
# Author: E.M Nimesh Kavinda Amarasinghe
# Web: www.edirisooriya.com
# Purpose: Automates closing apps, profile removal, registry cleanup, credential purge, and reboot prompt

# Admin Check
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as Administrator." -ForegroundColor Red
    pause
    exit
}

Write-Host "Starting Outlook Tenant-to-Tenant Cleanup..." -ForegroundColor Cyan

# 1. Kill Microsoft background apps (expanded list incl. WAM Broker)
$apps = "outlook","teams","onedrive","winword","excel","powerpnt",
        "OfficeClickToRun","MSOSYNC","MSOUC","Microsoft.AAD.BrokerPlugin"
foreach ($app in $apps) {
    Get-Process -Name $app -ErrorAction SilentlyContinue | Stop-Process -Force
    Write-Host "Closed $app if running." -ForegroundColor Yellow
}

# 2. Delete Outlook Profiles
$profilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
if (Test-Path $profilePath) {
    Remove-Item -Path $profilePath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Outlook profiles removed." -ForegroundColor Yellow
}

# 3�6. Registry Cleanup
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover\RedirectServers" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Identity" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Roaming\Identities" -Recurse -Force -ErrorAction SilentlyContinue

# 7. Credential Manager Cleanup (optimized loop)
$targets = @("MicrosoftOffice","Outlook","Exchange","ADAL","MicrosoftIdentity","msteams")
$creds = cmdkey /list
foreach ($line in $creds) {
    foreach ($target in $targets) {
        if ($line -match $target) {
            $cred = ($line -split ":")[1].Trim()
            if ($cred) {
                Write-Host "Removing credential: $cred" -ForegroundColor Yellow
                cmdkey /delete:$cred | Out-Null
            }
        }
    }
}

Write-Host "Cleanup complete." -ForegroundColor Green

# 8. Prompt for reboot
$choice = Read-Host "Do you want to reboot now? (Y/N)"
if ($choice -match "^[Yy]$") {
    Write-Host "Rebooting system..." -ForegroundColor Cyan
    Restart-Computer -Force
} else {
    Write-Host "Please reboot manually before reconfiguring Outlook." -ForegroundColor Red
}