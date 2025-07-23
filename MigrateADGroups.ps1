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
        
        $GroupObject = [pscustomobject]::new()
        
        $GroupObject | Add-Member NoteProperty GroupName $Group.Name
        $GroupObject | Add-Member NoteProperty GroupMail $Group.Mail
        $GroupObject | Add-Member NoteProperty GroupMembers ""

        

        $UserMembers = Get-ADGroupMember -Identity $Group.Name | Where-Object { $_.ObjectClass -eq "User" }

        foreach ($Member in $UserMembers) {


            $Users = Get-ADUser -Identity $Member.SamAccountName -Properties Name, Mail
            
            foreach ($User in $Users) {
            $MemberObject = [PSCustomObject]::new()
            $MemberObject | Add-Member NoteProperty Name $User.Name
            $MemberObject | Add-Member NoteProperty Mail $User.Mail

            }
            $GroupObject.GroupMembers += $MemberObject
        }
        $Groups += $GroupObject
    }

    $Groups
    $Groups.GroupMembers
}





Get-TargetADGroups -OrgUnit "Migrated Distros - Unsynced Folder" -GroupType "Universal"