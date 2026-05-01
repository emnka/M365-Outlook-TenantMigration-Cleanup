<#
.SYNOPSIS
M365 Tenant-to-Tenant Migration Full Cleanup Tool

.DESCRIPTION
Performs deep cleanup of Outlook, Office identity, and Windows AAD/WAM cache
to resolve sign-in and Autodiscover issues after tenant migration.

Author: Nimesh Kavinda
Role: Senior M365 Administrator
Last Updated: 2026
#>

# =========================
# 🔷 INITIAL SETUP
# =========================

$LogFile = "$env:TEMP\M365-TenantCleanup.log"

function Write-Log {
    param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp - $Message"
    Write-Host $entry
    Add-Content -Path $LogFile -Value $entry
}

Write-Log "===== Starting M365 Tenant Migration Cleanup ====="

# =========================
# 🔷 CHECK ADMIN RIGHTS
# =========================

if (-NOT ([Security.Principal.WindowsPrincipal] `
[Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole] "Administrator")) {

    Write-Log "ERROR: Script must be run as Administrator!"
    Exit
}

# =========================
# 🔷 STOP PROCESSES
# =========================

Write-Log "Stopping Office and related processes..."

$processes = @(
    "Outlook","Teams","msedge","OneDrive",
    "OfficeClickToRun","Microsoft.AAD.BrokerPlugin"
)

foreach ($proc in $processes) {
    Get-Process -Name $proc -ErrorAction SilentlyContinue | Stop-Process -Force
}

# =========================
# 🔷 REMOVE OUTLOOK PROFILES (REGISTRY)
# =========================

Write-Log "Removing Outlook Profiles..."

Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles" `
-Recurse -Force -ErrorAction SilentlyContinue

# =========================
# 🔷 AUTODISCOVER CLEANUP
# =========================

Write-Log "Cleaning Autodiscover cache..."

Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover\RedirectServers" `
-Recurse -Force -ErrorAction SilentlyContinue

Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover" `
-Recurse -Force -ErrorAction SilentlyContinue

# =========================
# 🔷 OFFICE IDENTITY CLEANUP
# =========================

Write-Log "Removing Office Identity cache..."

Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Identity" `
-Recurse -Force -ErrorAction SilentlyContinue

Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Roaming\Identities" `
-Recurse -Force -ErrorAction SilentlyContinue

# =========================
# 🔷 WINDOWS CREDENTIAL MANAGER CLEANUP
# =========================

Write-Log "Clearing Credential Manager entries..."

cmdkey /list | ForEach-Object {
    if ($_ -match "Microsoft|Office|Outlook|Exchange|ADAL|Teams|Identity") {
        $target = ($_ -split ":")[1].Trim()
        cmdkey /delete:$target
        Write-Log "Removed credential: $target"
    }
}

# =========================
# 🔷 AZURE AD DEVICE UNJOIN
# =========================

Write-Log "Checking Azure AD Join status..."

$dsreg = dsregcmd /status

if ($dsreg -match "AzureAdJoined\s*:\s*YES") {
    Write-Log "Device is Azure AD joined. Leaving..."
    Start-Process -FilePath "dsregcmd.exe" -ArgumentList "/leave" -Wait -NoNewWindow
    Write-Log "Device removed from Azure AD"
} else {
    Write-Log "Device not Azure AD joined"
}

# =========================
# 🔷 WAM / AAD BROKER CACHE CLEANUP (CRITICAL)
# =========================

Write-Log "Clearing WAM / AAD Broker cache..."

$aadPath = "$env:LOCALAPPDATA\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy"

if (Test-Path $aadPath) {
    Remove-Item -Path $aadPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "AAD Broker cache removed"
} else {
    Write-Log "AAD Broker cache not found"
}

# =========================
# 🔷 ADDITIONAL TOKEN CLEANUP
# =========================

Write-Log "Removing additional identity token caches..."

Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\TokenBroker" `
-Recurse -Force -ErrorAction SilentlyContinue

Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\IdentityCache" `
-Recurse -Force -ErrorAction SilentlyContinue

# =========================
# 🔷 WORKPLACE JOIN CLEANUP
# =========================

Write-Log "Cleaning Workplace Join registry..."

Remove-Item -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin" `
-Recurse -Force -ErrorAction SilentlyContinue

# =========================
# 🔷 FINAL STEP
# =========================

Write-Log "Cleanup complete."

Write-Host ""
Write-Host "=========================================" -ForegroundColor Yellow
Write-Host "⚠️  IMPORTANT: REBOOT REQUIRED" -ForegroundColor Red
Write-Host "=========================================" -ForegroundColor Yellow
Write-Host ""

$reboot = Read-Host "Do you want to reboot now? (Y/N)"

if ($reboot -eq "Y") {
    Write-Log "Reboot initiated by user"
    Restart-Computer -Force
} else {
    Write-Log "User skipped reboot (manual restart required)"
}