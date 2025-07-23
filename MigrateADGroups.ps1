

function Get-TargetADGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$OrgUnit,

        [Parameter(Mandatory = $true)]
        [string]$GroupType
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
        

        $UserMembers = Get-ADGroupMember -Identity $Group.Name | Where-Object { $_.ObjectClass -eq "User" }

        foreach ($Member in $UserMembers) {


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
   return $Groups
}

function Connect-365Services {
       [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [switch]$Graph

        # TODO Incorporate at a later time
        #[switch]$Exchange,
        #[switch]$Teams

    )

    if ($Graph) {

        if ($null -eq (Get-Module -ListAvailable Microsoft.Graph)){

            Install-Module -Name Microsoft.Graph -Force

        }else {

            Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All"
        }

    }
    
    <# TODO Incorporate at a later time
    elseif ($Exchange) {

        ExchangeOnlineManagement
        
    }elseif ($Teams) {

        MicrosoftTeams
    }
    #>

}


function Get-GraphGroupIDs {
        [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupNames
    )

    foreach ($Group in $GroupNames){ 

        Get-MgGroup -Filter "displayName -eq '$($Group)'" | Select-Object DisplayName, ID

    }


}

function Get-GraphGroups {
        [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupID
    )

    foreach ($ID in $GroupID) {

        Get-MgGroup -GroupID $ID

    }
}


function Remove-GraphGroups {
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

function New-GraphGroups{
                [CmdletBinding(DefaultParameterSetName = "Remove")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Remove")]
        [switch]$Remove,
        [Parameter(Mandatory = $true, ParameterSetName = "Remove")]
        [string]$GroupID


    )

}

function Get-GraphUserIds {
                [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Remove")]
        [string]$UserEnail
    )

    foreach ($Email in $UserEmail) {

        Get-MgUser -Filter "mail -eq '$($Email)'" | Select-Object DisplayName, ID

    }

}


#Get-TargetADGroups -OrgUnit "Migrated Distros - Unsynced Folder" -GroupType "Universal"