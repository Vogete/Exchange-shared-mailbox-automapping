function Get-GroupBySID {
    param (
        [string]$GroupSID
    )

    try {
        $oGroup = Get-ADGroup -Filter {objectSid -eq $GroupSID} -properties * -ResultSetSize $null
    }
    catch {
        return $false
    }
    return $oGroup
}

function Get-GroupForSharedMailbox {
    param (
        $SharedMailbox,
        [string]$Prefix,
        [string]$Postfix
    )
    $SearchTerm = $Prefix + $SharedMailbox.DisplayName + $Postfix
    $Group = Get-GroupBySamAccountName -GroupName $SearchTerm

    return $Group
}

function Get-GroupBySamAccountName {
    param (
        [string]$GroupName
    )
    # $Group = Get-Group $GroupName -ErrorAction SilentlyContinue -ResultSize Unlimited
    try {
        $oGroup = Get-ADGroup -Filter {sAMAccountName -eq $GroupName} -properties * -ResultSetSize $null
    }
    catch {
        return $false
    }
    return $oGroup
}

function Get-UserBySamAccountName {
    param (
        [string]$Username
    )
    # $Group = Get-Group $Username -ErrorAction SilentlyContinue -ResultSize Unlimited
    try {
        $User = Get-ADUser -Filter {sAMAccountName -eq $Username} -properties * -ResultSetSize $null
    }
    catch {
        return $false
    }
    return $User
}

function Get-ADObjectByUserPrincipalName {
    param (
        [string]$Username
    )
    try {
        $ADObj = Get-ADObject -Filter {userPrincipalName -eq $Username} -properties * -ResultSetSize $null
    }
    catch {
        return $false
    }
    return $ADObj
}


function Get-RecursiveADGroupMembers {
    param (
        [Microsoft.ActiveDirectory.Management.ADGroup]$Group
    )

    [array]$GroupMembers = @()

    $GroupMembers = Get-ADGroupMember -Identity $Group.SamAccountName -Recursive
    # Write-Host ($GroupMembers | Format-Table | Out-String)

    return $GroupMembers
}
