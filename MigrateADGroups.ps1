[CmdletBinding()]
param(
    [Parameter(Mandatory=$true,Position=0,ValueFromPipeline,ValueFromPipelineByPropertyName,HelpMessage="Enter the OU that you want to process in double quotes")]
    [Alias("DistinguishedName")]
    [string]$OrgUnit,
    [Parameter(Mandatory=$true,Position=1,ValueFromPipeline,HelpMessage="Enter Universal, DomainLocal, or Global with NO double quotes")]
    [ValidateSet("Universal","DomainLocal","Global")]
    [string]$GroupScope,
    [Parameter()]
    [bool]$SavetoFile

)

# Create temp folder for logs
if (!(Test-Path -Path "$env:TEMP\MigrateADGroups")){

    New-Item -Path "$env:TEMP" -ItemType Directory -Name "MigrateADGroups"
}

# Serializing the parameter values to be used in new pwoershell instances
    $ParentInstanceValues = @{
    OrgUnit = $OrgUnit
    GroupScope = $GroupScope
    SavetoFile = $SavetoFile
}

$ParentInstanceValues | ConvertTo-Json -Depth 3 | Set-Content "$env:TEMP\MigrateADGroups\ParameterValues.json"


# Elevating to administrator if not running as administrator   
if (-not [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Host "Process requires elevated rights...`n`tElevating to administrator..."

    Start-Process -FilePath "powershell.exe" -ArgumentList @(
    "-File", "`"$PSCommandPath`"",
    "-OrgUnit", "`"$($OrgUnit)`"",
    "-GroupScope", "$GroupScope"
) -Verb RunAs
Stop-Process $PID

}




#Global Variables
$Global:LogPath = "$env:TEMP\MigrateADGroups\"
$global:TimeStamp = (Get-Date).ToString("MM/dd/yyyy HH:mm:ss")



# Logs to console and to file
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

    $FileLocation = "$global:Logpath\Migration-Log.txt"
    

    $console_logs = @{
        Info = "$($timestamp) : Line : $($MyInvocation.ScriptLineNumber) : $($message)"
        Success = "$($timestamp) : Line : $($MyInvocation.ScriptLineNumber) : $($message)"
        Error = "$($timestamp) : ERROR: An error occurred at Line: $($MyInvocation.ScriptLineNumber) with the following error message: `n$($Error[0])"
    }

    switch ($Type) {
        ([Logs]::Info) {$console_logs.Info | Tee-Object -FilePath $FileLocation -Append ; break}
        ([Logs]::Success) {$console_logs.Success | Tee-Object -FilePath $FileLocation -Append ; break }
        ([Logs]::Error) {$console_logs.Error | Tee-Object -FilePath $FileLocation -Append; break} 

    }
}


# Stops stript execution and logs to console and file
function Stop-ScriptExecution {
    param (
        [switch]$ExitScript
    
    )
    $LogPath = "$env:TEMP\MigrateADGroups\Migration-Log.txt"
    $Failure =  "$($timestamp) : FAILURE: Script Halted at Line: $($MyInvocation.ScriptLineNumber) "

    if ($ExitScript){ 

        throw "$message`n  `n$($Failure)"  | Tee-Object -FilePath $LogPath -Append
    }
}

# Create a file based checkpoint
function New-ScriptCheckpoint {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [string]$FileName
    )
    $CheckPointLocation = "$env:TEMP\MigrateADGroups\"

    if ($FileName) {

        New-Item -Path $CheckPointLocation -ItemType File -Name $FileName

        New-MigrationLog -Type info -Message "$($FileName) has been created at: $($CheckPointLocation)"
    }


}   



# Checks for a file based checkpoint
function Get-ScriptCheckpoint {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$FileName
    )
    $CheckPointLocation = "$env:TEMP\MigrateADGroups\"

    if ($FileName) {

        if (Get-ChildItem -Path $CheckPointLocation -Name $FileName -ErrorAction SilentlyContinue)  {
            return "Exists"
        }
        else {
            return "Not Exists"
        }
    }
  
}   

# Takes [string] for the OU and gets the groups and members of the OU, returns [pscustomobject]
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

    if ($OrgUnit -notmatch "^OU=") {

        $DistinguishedName = (Get-ADOrganizationalUnit -Filter "Name -eq '$($OrgUnit)'").DistinguishedName
    }
    else {
        $DistinguishedName = $OrgUnit
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
            "UserName" = $User.Name 
            "UserEmail" = $User.Mail # TODO Write in error handling at some point to handle users without email in the email field.
                }
                $Groups += $GroupObject
            }
            
        }
    }

    $PreMigrationReport = "$Global:LogPath\PreMigrationReport_ADGroups.csv"

    if ($SavetoFile){

        $Groups | Export-Csv -Path $PreMigrationReport -NoTypeInformation

    }

    return $Groups
}

# Takes  [pscustomobject] from Get-TargetADGroups and queries Excahnge Online for the groups and members in the cloud
function Get-CloudGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline)]
        [PSCustomObject]$InputObject,
        [Parameter()]
        [switch]$SavetoFile
    )

    Begin {
        $CloudGroups = @()
    }

    Process {
        foreach ($Object in $InputObject) {
            $TargetGroup = Get-DistributionGroup -Identity $Object.GroupEmail

            foreach ($Group in $TargetGroup) {
                $GroupMembers = Get-DistributionGroupMember -Identity $Group.Identity

                foreach ($User in $GroupMembers) {
                    $CloudGroups += New-Object PSObject -Property @{
                        "GroupDisplayName" = $Object.GroupName
                        "GroupEmail"       = $Object.GroupEmail
                        "UserDisplayName"  = $User.DisplayName
                        "UserEmail"        = $User.PrimarySmtpAddress
                    }
                }
            }
        }
    }

    End {
        $PreCloudGroupRemovalReport = "$Global:LogPath\PreCloudRemovalReport_M365Groups.csv"

        if ($SavetoFile) {
            $CloudGroups | Export-Csv -Path $PreCloudGroupRemovalReport -NoTypeInformation
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
function Restart-ScriptSession
{
    # [CmdletBinding()]
    # Param(
    #     [Parameter(Mandatory=$true)]
    #     [string]$OrgUnit,
    #     [Parameter(Mandatory=$true)]
    #     [ValidateSet("Universal", "DomainLocal", "Global")]
    #     [string]$GroupScope,
    #     [Parameter()]
    #     [bool]$SavetoFile = $false
    # )

    # -OrgUnit $($OrgUnit.Trim()) -GroupScope $($GroupScope) -SavetoFile $($SavetoFile) 
Start-Process -FilePath "powershell.exe" -ArgumentList @(
    "-File", "`"$PSCommandPath`"",
    "-OrgUnit", "`"$OrgUnit`"",
    "-GroupScope", "`"$GroupScope`""
) -WindowStyle Normal

Stop-Process $PID

}



######
## START SCRIPT
######

# Checkpoint Switch 
pause 
Write-Host "305"
$IsFileExists = Get-ScriptCheckpoint -FileName "CheckPoint_1"
# TODO get switch working


if (($IsFileExists -eq "Not Exists")){

     New-MigrationLog -Type Info -Message "Checking execution policy..."

    if ((Get-ExecutionPolicy -Scope Process) -notin @("Bypass", "Unrestricted") -and (Get-ExecutionPolicy) -ne "Unrestricted") {
        New-MigrationLog -Type Info -Message "Policy current set to [$(Get-ExecutionPolicy -Scope Process)]. Bypass or Unrestricted required: Restarting script."
    }

Write-Host "Line 305"
Pause

    New-MigrationLog -Type Info -Message "Valid execution policy set"

    # Install Excahnge Online module

    New-MigrationLog -Type Info -Message "Verifying Exchange module is installed..." 
Write-Host "Pause 2"

    if (-not (Get-Module -ListAvailable ExchangeOnlineManagement))
    {  
        New-MigrationLog -Type info  -Message "Missing Exchange module Installing module..."

    try {

        New-MigrationLog -Type info  -Message "Checking for Nuget and installing if needed"
    
        if (-not (Get-PackageProvider -Name Nuget -ListAvailable  -ErrorAction SilentlyContinue)) {

            Install-PackageProvider -Name Nuget -Force
        }
    
    } catch {

        New-MigrationLog -type Error
        # Stop-ScriptExecution -ExitScript

    }
Write-Host "Line 334"
Pause
    try{
        Install-Module -Name ExchangeOnlineManagement -Force

        New-MigrationLog -Type info  -Message "ExchangeOnlineManagement Module installed"
    } catch {

        New-MigrationLog -type Error
        Stop-ScriptExecution -ExitScript
    }


    New-MigrationLog -Info -Message "Generating checkpoint at: $($Global:LogPath)"
    New-ScriptCheckpoint -FileName "CheckPoint_1"

Write-Host "Line 345"

    New-MigrationLog -Type Info -message  "Script is restarting" 
    try {
        if (Get-ScriptCheckpoint -FileName "Checkpoint_1"){
            Start-Process -FilePath "powershell.exe" -ArgumentList @(
                "-File", "`"$PSCommandPath`""
                "-OrgUnit", "`"$OrgUnit`"",
                "-GroupScope", "`"$GroupScope`"",
                "-SavetoFile", "`"$SavetoFile`""

            ) -WindowStyle Normal

            
        }
        Stop-Process $PID
    }
    catch {
        
        New-MigrationLog -type Error
    }
    }
}
elseif ($IsFileExists -eq "Exists") {
        New-MigrationLog -Type Info -Message "Importing module: [ExchangeOnlineManagement]"
        New-MigrationLog -Type Info -message "Importing Module after restart"

    Pause
    # Derializing the parameter values to be used in new pwoershell instances
    # Serialization is at the beginning of this script
        $JSONValues = @{
        $ParentInstanceValues=  Get-Content -Path "$env:TEMP\MigrateADGroups\ParameterValues.json" | ConvertFrom-Json  
        OrgUnit = $ParentInstanceValues.OrgUnit
        GroupScope = $ParentInstanceValues.GroupScope
        SavetoFile = $ParentInstanceValues.SavetoFile
    }

    $JSONValues.OrgUnit
    $JSONValues.GroupScope
    $JSONValues.SavetoFile

    Import-Module -Name ExchangeOnlineManagement -Verbose

Write-Host "Line 369"
Pause
    # Everything below this comment should start in a new session

    New-MigrationLog -Info -Message "Connecting to Exchange Online..."

    Connect-ExchangeOnline

    New-MigrationLog -Type Info -Message "Starting Active Directory group migration"

    New-MigrationLog -Type Info -Message "Getting AD groups and group members from: $($OrgUnit)"

    #$OrgUnit, $GroupScope = $OrgUnit, $GroupScope

    # Gather backup reports
    Get-TargetADGroups -OrgUnit $JSONValues.OrgUnit -GroupScope $JSONValues.GroupScope -SavetoFile $JSONValues.SavetoFile
    $CloudGroups = Get-TargetADGroups -OrgUnit $JSONValues.OrgUnit.Trim('"') -GroupScope $JSONValues.GroupScope -SavetoFile $JSONValues.SavetoFile
    $CloudGroups
    Pause

    # Verify backup reports
    if (!(Test-Path -Path "$Global:LogPath\PreMigrationADGroups_Backup.csv")){

    try {
    
        Get-TargetADGroups -OrgUnit $JSONValues.OrgUnit -GroupScope $JSONValues.GroupScope -SavetoFile $JSONValues.SavetoFile

        if (Test-Path -Path "$Global:LogPath\PreMigrationADGroups_Backup.csv"){

            New-MigrationLog -Type Info -message "AD Groups with Users has been backed up to $("$Global:LogPath\PreMigrationADGroups_Backup.csv")" 
        }
    }
    catch {

        New-MigrationLog -Type Error
    }

} else{

    New-MigrationLog -Type Success -Message "AD Groups with Users has been backed up to $("$Global:LogPath\PreMigrationADGroups_Backup.csv")"

}

if (!(Test-Path -Path "$Global:LogPath\PreMigrationCloudGroups_Backup.csv")){

    try {   

        $CloudGroups | Where-Object {$_ -ne $null} | Get-CloudGroups 

         if (Test-Path -Path ("$Global:LogPath\PreMigrationCloudGroups_Backup.csv")){
            
            New-MigrationLog -Type Info -message "AD Groups with Users has been backed up to $("$Global:LogPath\PreMigrationCloudGroups_Backup.csv")"
        }
    }
    catch {

        New-MigrationLog -Type Error
    }
} else{

    New-MigrationLog -Type Success -Message "Cloud Groups with Users has been backed up to $("$Global:LogPath\PreMigrationCloudGroups_Backup.csv")"
}

}






