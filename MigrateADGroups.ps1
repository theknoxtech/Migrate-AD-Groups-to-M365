<#
.SYNOPSIS
Migrates Active Directory groups to Microsoft 365.

.DESCRIPTION
The script exports the on-premises AD groups and their members from the
specified OU, backs up any existing Microsoft 365 groups with matching
names, creates the groups in Microsoft 365 if they do not exist, and adds
all members to the cloud groups. Backups are written to a local Backups
folder in the script directory.

.EXAMPLE
./MigrateADGroups.ps1 -OrgUnit "OU=Groups,DC=contoso,DC=com" -GroupScope Universal
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OrgUnit,

    [Parameter(Mandatory)]
    [ValidateSet('Universal','Global','DomainLocal')]
    [string]$GroupScope,

    [switch]$WhatIf
)

function Get-ADGroupsWithMembers {
    param(
        [string]$SearchBase,
        [string]$Scope
    )
    $groups = Get-ADGroup -SearchBase $SearchBase -SearchScope Subtree -GroupScope $Scope -ErrorAction Stop
    foreach ($g in $groups) {
        $members = Get-ADGroupMember -Identity $g -Recursive -ErrorAction SilentlyContinue |
            Where-Object { $_.ObjectClass -eq 'user' }
        [pscustomobject]@{
            Group          = $g
            Name           = $g.Name
            SamAccountName = $g.SamAccountName
            Members        = $members
        }
    }
}

function Export-GroupBackup {
    param(
        [string]$Path,
        [array]$Groups
    )
    $Groups | ForEach-Object {
        foreach ($m in $_.Members) {
            [pscustomobject]@{
                GroupName              = $_.Name
                MemberSamAccountName   = $m.SamAccountName
                MemberUserPrincipalName = $m.UserPrincipalName
            }
        }
    } | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function Get-M365GroupByName {
    param([string]$Name)
    Get-MgGroup -Filter "displayName eq '$Name'"
}

function Ensure-M365Group {
    param([pscustomobject]$ADGroup)
    $cloudGroup = Get-M365GroupByName -Name $ADGroup.Name
    if (-not $cloudGroup) {
        $params = @{
            DisplayName     = $ADGroup.Name
            MailEnabled     = $false
            MailNickname    = $ADGroup.SamAccountName
            SecurityEnabled = $true
        }
        if (-not $WhatIf) {
            $cloudGroup = New-MgGroup @params
        } else {
            Write-Host "WhatIf: would create group $($ADGroup.Name)"
        }
    }
    $cloudGroup
}

function Sync-Members {
    param(
        [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphGroup]$CloudGroup,
        [array]$Members
    )
    foreach ($m in $Members) {
        $user = Get-MgUser -Filter "userPrincipalName eq '$($m.UserPrincipalName)'"
        if ($user) {
            if (-not $WhatIf) {
                New-MgGroupMember -GroupId $CloudGroup.Id -DirectoryObjectId $user.Id -ErrorAction SilentlyContinue
            } else {
                Write-Host "WhatIf: would add $($m.UserPrincipalName) to $($CloudGroup.DisplayName)"
            }
        }
    }
}

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Users
Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All" | Out-Null

$backupRoot = Join-Path $PSScriptRoot 'Backups'
if (-not (Test-Path $backupRoot)) { New-Item -ItemType Directory -Path $backupRoot | Out-Null }

$adGroups = Get-ADGroupsWithMembers -SearchBase $OrgUnit -Scope $GroupScope
Export-GroupBackup -Path (Join-Path $backupRoot 'ADGroups.csv') -Groups $adGroups

$cloudGroups = foreach ($g in $adGroups) { Get-M365GroupByName -Name $g.Name }
$cloudGroups | Export-Csv -Path (Join-Path $backupRoot 'M365Groups.csv') -NoTypeInformation -Encoding UTF8

foreach ($g in $adGroups) {
    $target = Ensure-M365Group -ADGroup $g
    if ($target) {
        Sync-Members -CloudGroup $target -Members $g.Members
    }
}

Write-Host "Migration complete. Backups stored in $backupRoot"
