# Migrate AD Groups to M365

PowerShell script to migrate Active Directory groups and their members to Microsoft 365.

The script exports on-premises groups and their membership, backs up any existing cloud groups,
creates missing groups in Microsoft 365, and adds members to the corresponding cloud groups.
Backups are stored in a `Backups` folder next to the script.

## Usage

```powershell
./MigrateADGroups.ps1 -OrgUnit "OU=Groups,DC=contoso,DC=com" -GroupScope Universal
```

Requires the **ActiveDirectory** module for on-premises queries and the
**Microsoft.Graph** modules for cloud operations. Run with appropriate administrative
permissions and connectivity to both environments.
