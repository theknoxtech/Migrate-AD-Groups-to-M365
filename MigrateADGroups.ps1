#Requires -RunAsAdministrator
param(
    [string]$OrgUnit,
    [string]$GroupType,
    [switch]$SavetoFile
)





#Global Variables
$Global:PreMigrationReport = "$env:USERPROFILE\Documents\PreMigrationReport_ADGroups.csv"
$Global:PreCloudGroupRemovalReport = "$env:USERPROFILE\Documents\PreCloudRemovalReport_M365Groups.csv"
$global:TimeStamp = (Get-Date).ToString("MM/dd/yyyy HH:mm:ss")



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

        Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All"

    }elseif ($ImportGraph){

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


# TODO refactor this for exchange
function IsGraphConnected {

$AdminUser = Read-Host "Enter the email of the admin account used to conenct to Graph" 

    If (Get-MGUser -UserId "$($AdminUser)"){

        return $true

    }else{

        return $false

    }


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


#Get-TargetADGroups -OrgUnit "Migrated Distros - Unsynced Folder" -GroupType "Universal"
######
## START SCRIPT
######


if (!(Test-Path -Path $Global:PreMigrationReport)){

    try {
    
        throw "Pre-Migration Report NOT found, Exiting script"

    }
    catch {

        New-MigrationLog -Type Error
    }

}else{

    New-MigrationLog -Type Success -Message "AD Groups with Users has been backed up to $($PreMigrationReport)"

}



New-MigrationLog -Type Info -Message "Checking execution policy..."
if ((Get-ExecutionPolicy -Scope Process) -notin @("Bypass", "Unrestricted") -and (Get-ExecutionPolicy) -ne "Unrestricted") {
    New-MigrationLog -Type Info -Message "Policy current set to [$(Get-ExecutionPolicy -Scope Process)]. Bypass or Unrestricted required: Restarting script."
    Restart-ScriptSession
}


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

    Connect-365Services -Install

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

    Connect-365Services -Import
}


# Everything below this comment should start in a new session

New-MigrationLog -Info -Message "Connecting to Exchange Online..."

Connect-365Services -Connect

if (IsGraphConnected) {
    New-MigrationLog -Type Success -Message "You are connected to Graph"

    New-MigrationLog -Type Info -Message "Starting backup of groups and users"


}

$ADGroups




