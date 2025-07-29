[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$OrgUnit,
    [Parameter(Mandatory=$true)]
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

function Get-OUDistinguishedName{
    param(
        [string]$OrgUnit
    )
    $DistinguishedName = Get-ADOrganizationalUnit -Filter "Name -like '$($OrgUnit)'" | Select-Object -ExpandProperty DistinguishedName

    return $DistinguishedName
}  


function Get-TargetADGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$OrgUnit,

        [Parameter(Mandatory = $true)]
        [string]$GroupScope,

        [Parameter(Mandatory = $true)]
        [string]$DistinguishedName,

        [Parameter()]
        [switch]$SavetoFile


    )
    
    $TargetOU = (Get-OUDistinguishedName -OrgUnit $OrgUnit)

    $TargetGroups = Get-ADGroup -Filter "GroupScope -eq '$($GroupScope)'" -SearchBase ($TargetOU) -Property Mail | Select-Object Name, Mail

    $Groups = @()

    foreach ($Group in $TargetGroups) {
        <#
        $GroupObject = [pscustomobject]::new()
        
        $GroupObject | Add-Member NoteProperty GroupName $Group.Name
        $GroupObject | Add-Member NoteProperty GroupMail $Group.Mail
        $GroupObject | Add-Member NoteProperty GroupMembers ""
        #>
        

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

# function Connect-365Services {
#     [CmdletBinding()]
#     param (
#         [switch]$Install,
#         [switch]$Import,
#         [switch ]$Connect

#         # TODO Incorporate at a later time
#         #[switch]$Exchange,
#         #[switch]$Teams
#     )

    

#     if ($Install) {

#         Install-Module -Name ExchangeOnlineManagement -Force

#     }elseif ($Connect) {

#         Connect-ExchangeOnline

#     }elseif ($Import){

#         Import-Module -Name ExchangeOnlineManagement -Force
#     }

    
    
#     <# TODO Incorporate at a later time
#     elseif ($Exchange) {

#         ExchangeOnlineManagement
        
#     }elseif ($Teams) {

#         MicrosoftTeams
#     }
#     #>
# }

# Takes  [Array] from Get-TargetADGroups and return an [Array]
function Get-CloudGroups {
    [CmdletBinding()]
    param (
        [Parameter()]
        [Array]$Groups = (Get-TargetADGroups -OrgUnit $OrgUnit -GroupScope $GroupScope),
        [Parameter()]
        [switch]$SavetoFile
    )
    
    $ADGroups = $Groups

    $CloudGroups = @()

foreach($Group in $ADGroups) {
    $TargetGroup = Get-DistributionGroup  -Identity $Group.GroupMail
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
        [ValidateSet("Universal", "DomainLocal")]
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
$DistinguishedName = Get-ADOrganizationalUnit -Filter "Name -like '$($OrgUnit)'" | Select-Object -ExpandProperty DistinguishedName
Get-CloudGroups -Groups Get-TargetADGroups -OrgUnit $OrgUnit -GroupScope $GroupScope -DistinguishedName $DistinguishedName
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