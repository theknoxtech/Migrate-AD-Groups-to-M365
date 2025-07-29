[CmdletBinding()]
param(
    [Parameter(Mandatory=$true,Position=0,ValueFromPipeline,ValueFromPipelineByPropertyName,HelpMessage="Enter the OU that you want to process in double quotes")]
    [Alias("DistinguishedName")]
    [string]$OrgUnit,
    [Parameter(Mandatory=$true,Position=1,ValueFromPipeline,HelpMessage="Enter Universal, DomainLocal, or Global with NO double quotes")]
    [ValidateSet("Universal","DomainLocal","Global")]
    [string]$GroupScope,
    #[Parameter()]
    # [ValidateSet("Security","Distribution")]
    # [string]$GroupType,
    [Parameter()]
    [bool]$SavetoFile = $false
)


# Elevate session if not running with privileged rights
#[Security.Principal.WindowsIdentity]::GetCurrent()
# if (-not [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([security.principal.windowsbuiltinrole]::Administrator))
# {
#     Write-Host "Process requires elevated rights...`n`tElevating to administrator..."
#     $argsList = @(
#         "-NoProfile",
#         "-ExecutionPolicy Bypass",
#         "-File $PSCommandPath"
#     )

#     $PSBoundParameters.GetEnumerator() | ForEach-Object {
#         $argsList += "-$($_.Key) `"$($_.Value)`""

#         #$argsList += "-$($_.Key)"
#         #$argsList += "`"$($_.Value)`""
#     }

#     Write-Host $argsList

#     Start-Process -FilePath "powershell.exe" -ArgumentList $argsList -Verb runas -NoNewWindow:$true
#     Stop-Process -Id $PID -Force
# }


if (-not [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Host "Process requires elevated rights...`n`tElevating to administrator..."

    Start-Process -FilePath "powershell.exe" -ArgumentList @(
    "-File", "`"$PSCommandPath`"",
    "-OrgUnit", "`"$OrgUnit`"",
    "-GroupScope", "$GroupScope"
) -Verb RunAs

    #Start-Process -FilePath "powershell.exe" -ArgumentList @("-File", "`"$($PSCommandPath)`"", "-OrgUnit","`"$($OrgUnit)`"", "-GroupScope", "`"$($GroupScope)`"") -Verb runas
    #Stop-Process -Id $PID -Force
}



#Global Variables
$Global:PreMigrationReport = "$env:USERPROFILE\Documents\PreMigrationReport_ADGroups.csv"
$Global:PreCloudGroupRemovalReport = "$env:USERPROFILE\Documents\PreCloudRemovalReport_M365Groups.csv"
$global:TimeStamp = (Get-Date).ToString("MM/dd/yyyy HH:mm:ss")

Pause
Write-Host "Line 57"
function New-MigrationLog {
    param(
        [Logs]$Type,
        [string]$Message
    )

    enum Logs {
        Info
        Success
        Error
    }

    $LogPath = "$env:USERPROFILE\Documents\Migration-Log.txt"
    

    $console_logs = @{
        Info = "$($timestamp) : Line : $($MyInvocation.ScriptLineNumber) : $($message)"
        Success = "$($timestamp) : Line : $($MyInvocation.ScriptLineNumber) : $($message)"
        Error = "$($timestamp) : ERROR: An error occurred at Line: $($MyInvocation.ScriptLineNumber) with the following error message: `n$($Error[0])"
    }

    switch ($Type) {
        ([Logs]::Info) {$console_logs.Info | Tee-Object -FilePath $LogPath -Append ; break}
        ([Logs]::Success) {$console_logs.Success | Tee-Object -FilePath $LogPath -Append ; break }
        ([Logs]::Error) {$console_logs.Error | Tee-Object -FilePath $LogPath -Append; break} 

    }
}

function Stop-ScriptExecution {
    param (
        [switch]$ExitScript
    )
    $LogPath = "$env:USERPROFILE\Documents\Migration-Log.txt"
    $Failure =  "$($timestamp) : FAILURE: Script Halted at Line: $($MyInvocation.ScriptLineNumber) "

    if ($ExitScript){ 

        throw "$message`n  `n$($Failure)"  | Tee-Object -FilePath $LogPath -Append

    }
}

function Get-TargetADGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,Position=0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Alias("DistinguishedName")]
        [string]$OrgUnit,

        [Parameter(Mandatory = $true,Position=1)]
        [string]$GroupScope,

        [Parameter()]
        [switch]$SavetoFile


    )

    if ($OrgUnit -notmatch "^OU=.+m=,") {

        $DistinguishedName = (Get-ADOrganizationalUnit -Filter "Name -eq '$($OrgUnit)'").distinguishedname 

        $DistinguishedName.DistinguishedName
    }
   


    $TargetGroups = Get-ADGroup -Filter "GroupScope -eq '$($GroupScope)'" -SearchBase ($DistinguishedName) -Property Mail | Select-Object Name, Mail

    $Groups = @()

    foreach ($Group in $TargetGroups) {

        $GroupMembers = Get-ADGroupMember -Identity $Group.Name | Where-Object { $_.ObjectClass -eq "User" }

        foreach ($Member in $GroupMembers) {


            $Users = Get-ADUser -Identity $Member.SamAccountName -Properties Name, Mail
            
            foreach ($User in $Users) {
        
            $GroupObject = New-Object PSObject -Property @{
            "GroupName" = $Group.Name
            "GroupEmail" = $Group.Mail
            "MemberName" = $User.Name 
            "UserEmail" = $User.Mail # TODO Write in error handling at some point to handle users without email in the email field.
                }
                $Groups += $GroupObject
            }
            
        }
    }
    if ($SavetoFile){

        $Groups | Export-Csv -Path "$env:USERPROFILE\Documents\PreMigrationReport_ADGroups.csv" -NoTypeInformation

    }

    return $Groups
}

# Takes  [Array] from Get-TargetADGroups and return an [Array]
function Get-CloudGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline)]
        [PSCustomObject]$Groups,
        [Parameter()]
        [switch]$SavetoFile
    )
    
    Write-Verbose "Got group: $($Group | Out-String)"    

    $CloudGroups = @()

    if (!($Groups)){
        return 
    }
    else{
        foreach($Group in $Groups) {
        $TargetGroup = Get-DistributionGroup -Identity $Group.GroupEmail

        $GroupMembers = Get-DistributionGroupMember -Identity $TargetGroup.Identity
    
        foreach ($Member in ($GroupMembers)){

            $GroupObject = New-Object PSObject -Property @{
                "GroupDisplayName" = $Group.Name
                "GroupEmail" = $Group.Email
                "MemberEmail" = $Member.PrimarySMTPAddress
                "MemberDisplayName" = $Member.DisplayName
                }
                 $CloudGroups += $GroupObject
            }
    
            if ($SavetoFile){

                $CloudGroups | Export-Csv -Path "$env:USERPROFILE\Documents\PreMigrationReport_CloudGroups.csv" -NoTypeInformation
            
            
            }

            return $CloudGroups
        }
    }
}
function Remove-CloudGroups {
            [CmdletBinding(DefaultParameterSetName = "Remove")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Remove")]
        [switch]$Remove,
        [Parameter(Mandatory = $true, ParameterSetName = "Remove")]
        [string]$GroupID


    )

    foreach ($ID in $GroupID) {

        Remove-MgGroup -GroupID $ID

    }

}

function New-CloudGroups{
        [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupID

    )

}

function IsNugetInstalled {
        
    Get-PackageProvider -ListAvailable -Name Nuget -ErrorAction SilentlyContinue

}

function Restart-ScriptSession
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$OrgUnit,
        [Parameter()]
        [ValidateSet("Universal", "DomainLocal", "Global")]
        [string]$GroupScope = "Universal",
        [Parameter()]
        [bool]$SavetoFile
    )
    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$($PSCommandPath)`" -OrgUnit $($OrgUnit) -GroupScope $($GroupScope) -SavetoFile $($SavetoFile)"
    Stop-Process -Id $PID -Force
}
Write-Host "Line 293"
pause
#Get-TargetADGroups -OrgUnit "Migrated Distros - Unsynced Folder" -GroupScope "Universal"
######
## START SCRIPT
######
Write-Host "Line 299"
#$OUDistinguishedName = Get-OUDistinguishedNameDistinguisedName -OrgUnit $OrgUnit
Write-Host "Line 301"
$Groups = Get-TargetADGroups -OrgUnit $Orgunit.Trim('"') -GroupScope $GroupScope
$Groups | Where-Object {$_ -ne $null} | Get-CloudGroups -SavetoFile
Write-Host "Line 303"
pause

New-MigrationLog -Type Info -Message "Checking execution policy..."
if ((Get-ExecutionPolicy -Scope Process) -notin @("Bypass", "Unrestricted") -and (Get-ExecutionPolicy) -ne "Unrestricted") {
    New-MigrationLog -Type Info -Message "Policy current set to [$(Get-ExecutionPolicy -Scope Process)]. Bypass or Unrestricted required: Restarting script."
    pause
    #Restart-ScriptSession -OrgUnit $OrgUnit -GroupScope $GroupScope -SavetoFile $SavetoFile
}

pause

New-MigrationLog -Type Info -Message "Valid execution policy set"

# Install Excahnge Online module
# Spawns new session to ensure module is loaded because it is inconsistent otherwise
# Else shoud run when session restarts

New-MigrationLog -Type Info -Message "Verifying Exchange module is installed..." 

if ($null -eq (Get-Module -ListAvailable ExchangeOnlineManagement))
{
    try {

    New-MigrationLog -Type info  -Message "Checking for Nuget and installing if needed"
    
    if (!(IsNugetInstalled)) {

        Install-PackageProvider -Name Nuget -Force
    }

    New-MigrationLog -Type info  -Message "Missing Exchange module Installing module..."

    try{
        Install-Module -Name ExchangeOnlineManagement -Force
    } catch
    {
        New-MigrationLog -type info -message "Authentication failed. Script aborted!"
    }

    New-MigrationLog -Type info  -Message "Module installed Restarting script..."

    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$($PSCommandPath)`""

    New-MigrationLog -Info -Message "Terminating previous session and continuing execution in new session"

    Stop-Process -Id $PID -Force

    } catch {

        New-MigrationLog -Type Error
        
        Stop-ScriptExecution -ExitScript
    }

} else {

    New-MigrationLog -Type Info -Message "Importing module: [ExchangeOnlineManagement]"
    Write-Host "Imporint via 'Connect-365'"
    pause
    Import-Module -Name ExchangeOnlineManagement -Force
    Write-Host "Done"
    pause
}

pause
# Everything below this comment should start in a new session

New-MigrationLog -Info -Message "Connecting to Exchange Online..."
pause
Connect-ExchangeOnline

New-MigrationLog -Type Info -Message "Starting Active Directory group migration"

New-MigrationLog -Type Info -Message "Getting groups and group members from: $($OrgUnit)"

Get-TargetADGroups -OrgUnit $OrgUnit -GroupScope $GroupScope



if (!(Test-Path -Path $Global:PreMigrationReport)){

    try {
    
        throw "Pre-Migration Report NOT found, Exiting script"

    }
    catch {

        New-MigrationLog -Type Error
    }

} else{

    New-MigrationLog -Type Success -Message "AD Groups with Users has been backed up to $($PreMigrationReport)"

}

pause