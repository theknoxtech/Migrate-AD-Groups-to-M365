[CmdletBinding()]
param(
    [Parameter(Mandatory=$true,Position=0,ValueFromPipeline,ValueFromPipelineByPropertyName,HelpMessage="Enter the OU that you want to process in double quotes")]
    [Alias("DistinguishedName")]
    [string]$OrgUnit,
    [Parameter(Mandatory=$true,Position=1,ValueFromPipeline,HelpMessage="Enter Universal, DomainLocal, or Global with NO double quotes")]
    [ValidateSet("Universal","DomainLocal","Global")]
    [string]$GroupScope,
    [Parameter()]
    [bool]$SavetoFile,
    [int]$CurrentInstance = [System.Diagnostics.Process]::GetProcessById($PID)
)

#$CurrentPSSession = [System.Diagnostics.Process]::GetProcessById($this)

if (-not [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Host "Process requires elevated rights...`n`tElevating to administrator..."

    Start-Process -FilePath "powershell.exe" -ArgumentList @(
    "-File", "`"$PSCommandPath`"",
    "-OrgUnit", "`"$($OrgUnit)`"",
    "-GroupScope", "$GroupScope",
    "-CurrentInstacne", "$CurrentInstance"
) -Verb RunAs

    Stop-Process -Id $CurrentInstance -Force
    #Start-Process -FilePath "powershell.exe" -ArgumentList @("-File", "`"$($PSCommandPath)`"", "-OrgUnit","`"$($OrgUnit)`"", "-GroupScope", "`"$($GroupScope)`"") -Verb runas
    #Stop-Process -Id $PID -Force
}



#Global Variables
$Global:LogPath = "$env:USERPROFILE\Documents\"
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
    $LogPath = "$env:USERPROFILE\Documents\Migration-Log.txt"
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
    $CheckPointLocation = "$Global:LogPath\Documents\"

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
    $CheckPointLocation = "$Global:LogPath\Documents\"

    if ($FileName) {

        if (Get-Item -Path $CheckPointLocation -ItemType File -Name $FileName) {
            return $true
        }
        else {
            return $false
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
        foreach($Object in $InputObject) {
        $TargetGroup = Get-DistributionGroup -Identity $InputObject.GroupEmail

            foreach ($Group in $TargetGroup){

                $GroupMembers = Get-DistributionGroupMember -Identity $TargetGroup.Identity
    
                    foreach ($User in ($GroupMembers)){
                        $CloudGroups += New-Object PSObject -Property @{
                            "GroupDisplayName" = $InputObject.GroupName
                            "GroupEmail" = $InputObject.GroupEmail
                            "UserDisplayName" = $User.DisplayName
                            "UserEmail" = $User.PrimarySmtpAddress
                
                }
            }
        }
    }
}
    End {

        $PreCloudGroupRemovalReport = "$Global:LogPath\PreCloudRemovalReport_M365Groups.csv"

        if ($SavetoFile){
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

function IsNugetInstalled {
        
    Get-PackageProvider -ListAvailable -Name Nuget -ErrorAction SilentlyContinue

}

function Restart-ScriptSession
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$OrgUnit,
        [Parameter(Mandatory=$true)]
        [ValidateSet("Universal", "DomainLocal", "Global")]
        [string]$GroupScope,
        [Parameter()]
        [bool]$SavetoFile = $false
    )
        Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File |
        `"$($PSCommandPath)`" -OrgUnit $($OrgUnit.Trim()) -GroupScope $($GroupScope) -SavetoFile $($SavetoFile) "

       
}



######
## START SCRIPT
######


New-MigrationLog -Type Info -Message "Checking execution policy..."
if ((Get-ExecutionPolicy -Scope Process) -notin @("Bypass", "Unrestricted") -and (Get-ExecutionPolicy) -ne "Unrestricted") {
    New-MigrationLog -Type Info -Message "Policy current set to [$(Get-ExecutionPolicy -Scope Process)]. Bypass or Unrestricted required: Restarting script."
   
    #Restart-ScriptSession -OrgUnit $OrgUnit -GroupScope $GroupScope -SavetoFile $SavetoFile
}

$CurrentPSSession = [System.Diagnostics.Process]::GetProcessById($this)
Write-Host "The current session ID: $($CurrentPSSession)"
Write-Host "Pause 1"
Pause
New-MigrationLog -Type Info -Message "Valid execution policy set"

# Install Excahnge Online module
# Spawns new session to ensure module is loaded because it is inconsistent otherwise
# Else shoud run when session restarts

New-MigrationLog -Type Info -Message "Verifying Exchange module is installed..." 
Write-Host "Pause 2"

if ($null -eq (Get-Module -ListAvailable ExchangeOnlineManagement))
{  Write-Host "Pause 3"
    Pause
    New-MigrationLog -Type info  -Message "Missing Exchange module Installing module..."

    try {

        New-MigrationLog -Type info  -Message "Checking for Nuget and installing if needed"
    
        if (!(IsNugetInstalled)) {

            Install-PackageProvider -Name Nuget -Force
        }
    
    } catch {

        New-MigrationLog -type Error
        # Stop-ScriptExecution -ExitScript

    }

    try{
        
        Install-Module -Name ExchangeOnlineManagement -Force

        New-MigrationLog -Type info  -Message "Module installed Restarting script"

        Write-Host "Script Restart 1"

    } catch {

        New-MigrationLog -type Error
        Stop-ScriptExecution -ExitScript
    }
}


Write-Host "Script Restart"
Restart-ScriptSession -OrgUnit $OrgUnit -GroupScope $GroupScope  
Write-Host "382"
## TODO CheckPointOne Here

New-MigrationLog -Type Info -Message "Importing module: [ExchangeOnlineManagement]"
Write-Host "Importing Module  after restart"

Import-Module -Name ExchangeOnlineManagement

# Everything below this comment should start in a new session

New-MigrationLog -Info -Message "Connecting to Exchange Online..."

$ExchangeConnectionState = (Get-ConnectionInformation).state

switch ($ExchangeConnectionState) {
    ($_ -eq "Connected") {

        New-MigrationLog -Type info -Message "Exchange Online is connected"

        break}
    ($_ -eq "Broken") {

        Connect-ExchangeOnline 
    }

}


New-MigrationLog -Type Info -Message "Starting Active Directory group migration"

New-MigrationLog -Type Info -Message "Getting AD groups and group members from: $($OrgUnit)"

#$OrgUnit, $GroupScope = $OrgUnit, $GroupScope

# Gather backup reports
Get-TargetADGroups -OrgUnit $OrgUnit -GroupScope $GroupScope -SavetoFile
$CloudGroups = Get-TargetADGroups -OrgUnit $Orgunit.Trim('"') -GroupScope $GroupScope




if (!(Test-Path -Path $Global:PreMigrationReport)){

    try {
    
        Get-TargetADGroups -OrgUnit $OrgUnit -GroupScope $GroupScope -SavetoFile

        if (Test-Path -Path $Global:PreMigrationReport){

            New-MigrationLog -Type Info -message "AD Groups with Users has been backed up to $($PreMigrationReport)" 
        }
    }
    catch {

        New-MigrationLog -Type Error
    }

} else{

    New-MigrationLog -Type Success -Message "AD Groups with Users has been backed up to $($PreMigrationReport)"

}

if (!(Test-Path -Path $Global:PreCloudGroupRemovalReport)){

    try {   

        $CloudGroups | Where-Object {$_ -ne $null} | Get-CloudGroups -SavetoFile

         if (Test-Path -Path $Global:PreCloudGroupRemovalReport){
            
            New-MigrationLog -Type Info -message "AD Groups with Users has been backed up to $($PreCloudGroupRemovalReport)"
        }
    }
    catch {

        New-MigrationLog -Type Error
    }
} else{

    New-MigrationLog -Type Success -Message "Cloud Groups with Users has been backed up to $($PreCloudGroupRemovalReport)"
}

## TODO CheckPoinntTwo here
