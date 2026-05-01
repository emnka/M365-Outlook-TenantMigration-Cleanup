# M365-Outlook-TenantMigration-Cleanup
After tenant Migration, Outlook is not getting the new tenant and it is trying to move back and search on Old Tenant since it use the old cache to authenticate and find the path for the Email

# M365 Outlook Tenant-to-Tenant Migration Cleanup Script

## Overview
This repository contains a PowerShell automation script and batch launcher to fix **Outlook sign-in / profile creation issues** after a Microsoft 365 tenant-to-tenant migration.  
It automates the manual runbook steps required to clear cached Autodiscover redirects, identity remnants, and Windows WAM/AAD tokens that often block modern authentication.

## Features
- ✅ Closes all Microsoft background apps (Outlook Classic/New, Teams, OneDrive, Office apps, Click-to-Run, AAD Broker).
- ✅ Deletes Outlook profiles from registry.
- ✅ Cleans Autodiscover redirect cache and Autodiscover values.
- ✅ Removes Office Identity and Roaming Identities registry keys.
- ✅ Purges stale credentials from Windows Credential Manager.
- ✅ Prompts for reboot (mandatory for WAM token flush).
- ✅ Admin rights check to prevent silent failures.

## Repository Structure


## Usage
1. Clone or download this repository.
2. Run `RunOutlookTenantCleanup.bat` as **Administrator**.
3. Follow the reboot prompt (mandatory).
4. After reboot, open Outlook → add account → complete modern authentication → Outlook will recreate profile and OST.

## Requirements
- Windows 10/11
- Office 2016 or later (registry paths assume 16.0)
- Administrator rights

## Best Practices
- Always reboot after cleanup to flush WAM tokens.
- OST deletion alone is not sufficient — registry and identity cleanup is required.
- RedirectServers cache is the most common hidden blocker.
- Optional: Use Microsoft Support and Recovery Assistant (SaRA) for validation.

## Author
**E.M Nimesh Kavinda Amarasinghe**
## www.edirisooriya.com  
Senior Microsoft 365 Administrator & Cloud Infrastructure Specialist

## License
This project is licensed under the [MIT License](LICENSE) — feel free to use, adapt, and share with attribution.
