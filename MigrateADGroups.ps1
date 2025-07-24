[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$OrgUnit,
    [Parameter()]
    [ValidateSet("Universal", "DomainLocal")]
    [string]$GroupType = "Universal",
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

pause

if (-not [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Host "Process requires elevated rights...`n`tElevating to administrator..."

    Start-Process -FilePath "powershell.exe" -ArgumentList "-File `"$($PSCommandPath)`" -OrgUnit `"$($OrgUnit)`" -GroupType $($GroupType)" -Verb runas
    Stop-Process -Id $PID -Force
}

pause

#Global Variables
$Global:PreMigrationReport = "$env:USERPROFILE\Documents\PreMigrationReport_ADGroups.csv"
$Global:PreCloudGroupRemovalReport = "$env:USERPROFILE\Documents\PreCloudRemovalReport_M365Groups.csv"
$global:TimeStamp = (Get-Date).ToString("MM/dd/yyyy HH:mm:ss")

pause

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
        [Parameter(Mandatory = $true)]
        [string]$OrgUnit,

        [Parameter(Mandatory = $true)]
        [string]$GroupType,

        [Parameter()]
        [switch]$SavetoFile


    )

    $TargetOU = Get-ADOrganizationalUnit -Filter "Name -like '$($OrgUnit)'" | Select-Object -ExpandProperty DistinguishedName

    $TargetGroups = Get-ADGroup -Filter "GroupScope -eq '$($GroupType)'"  -SearchBase $TargetOU -Property Mail | Select-Object Name, Mail

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
            "Group Name" = $Group.Name
            "Group Email" = $Group.Mail
            "Member Name" = $User.Name
            "User Email" = $User.Mail
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

function Connect-365Services {
    [CmdletBinding()]
    param (
        [switch]$Install,
        [switch]$Import,
        [switch ]$Connect

        # TODO Incorporate at a later time
        #[switch]$Exchange,
        #[switch]$Teams
    )

    

    if ($Install) {

        Install-Module -Name ExchangeOnlineManagement -Force

    }elseif ($Connect) {

        Connect-ExchangeOnline

    }elseif ($Import){

        Import-Module -Name ExchangeOnlineManagement -Force
    }

    
    
    <# TODO Incorporate at a later time
    elseif ($Exchange) {

        ExchangeOnlineManagement
        
    }elseif ($Teams) {

        MicrosoftTeams
    }
    #>
}

function Get-CloudGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$OrgUnit,

        [Parameter(Mandatory = $true)]
        [string]$GroupType,

        [Parameter()]
        [switch]$SavetoFile
    )
    
    $ADGroups = Get-TargetADGroups -OrgUnit $OrgUnit -GroupType $GroupType

    $CloudGroups = @()

foreach($Group in $ADGroups) {
    $TargetGroup = Get-DistributionGroup  -Identity $Group.GroupMail
    $GroupMembers = Get-DistributionGroupMember -Identity $TargetGroup.Identity

    foreach ($Member in ($GroupMembers)){

        $GroupObject = New-Object PSObject -Property @{
            "Group Display Name" = $Group.Name
            "Group Email" = $Group.Email
            "Member Email" = $Member.PrimarySMTPAddress
            "Member Display Name" = $Member.DisplayName
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
        [string]$GroupType = "Universal",
        [Parameter()]
        [bool]$SavetoFile
    )
    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$($PSCommandPath)`" -OrgUnit $($OrgUnit) -GroupType $($GroupType) -SavetoFile $($SavetoFile)"
    Stop-Process -Id $PID -Force
}

#Get-TargetADGroups -OrgUnit "Migrated Distros - Unsynced Folder" -GroupType "Universal"
######
## START SCRIPT
######

Write-Host "Main"
pause
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

New-MigrationLog -Type Info -Message "Checking execution policy..."
if ((Get-ExecutionPolicy -Scope Process) -notin @("Bypass", "Unrestricted") -and (Get-ExecutionPolicy) -ne "Unrestricted") {
    New-MigrationLog -Type Info -Message "Policy current set to [$(Get-ExecutionPolicy -Scope Process)]. Bypass or Unrestricted required: Restarting script."
    pause
    #Restart-ScriptSession -OrgUnit $OrgUnit -GroupType $GroupType -SavetoFile $SavetoFile
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
        Connect-365Services -Install
    } catch
    {
        New-MigrationLog -type info -message "Authentication failed. Script aborted!"
    }

    New-MigrationLog -Type info  -Message "Module installed Restarting script..."

    #Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$($PSCommandPath)`""

    New-MigrationLog -Info -Message "Terminating previous session and continuing execution in new session"

    #Stop-Process -Id $PID -Force

    } catch {

        New-MigrationLog -Type Error
        
        #Stop-ScriptExecution -ExitScript
    }

} else {

    New-MigrationLog -Type Info -Message "Importing module: [ExchangeOnlineManagement]"
    Write-Host "Imporint via 'Connect-365'"
    pause
    Connect-365Services -Import
    Write-Host "Done"
    pause
}

pause
# Everything below this comment should start in a new session

New-MigrationLog -Info -Message "Connecting to Exchange Online..."
pause
Connect-365Services -Connect

Write-Host 'at end'
pause