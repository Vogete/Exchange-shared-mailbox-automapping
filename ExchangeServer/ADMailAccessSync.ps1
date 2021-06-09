$global:LogFolderName = $null
$global:LogFileName = $null

$LogFolderName = "logs\"
$LogFileName = $LogFolderName
$LogFileName += Get-Date -Format o | foreach {$_ -replace ":", "."}
$LogFileName += ".log"

function SetUpLogs {

    if (Test-Path $LogFolderName) {
        Write-Host "Log folder exists"
    } else {
        mkdir $LogFolderName
    }
}

function WriteLog {
    Param ([string]$logstring)
    # Write-Host $LogFileName

    $logstring += "`r`n"
    Add-content $LogFileName -value $logstring
}

function Get-GroupBySamAccountName {
    param([string]$GroupName)
    # $Group = Get-Group $GroupName -ErrorAction SilentlyContinue -ResultSize Unlimited
    try {
        $oGroup = Get-ADGroup -Filter {sAMAccountName -eq $GroupName} -properties * -ResultSetSize $null
    }
    catch {
        return $false
    }
    return $oGroup
}

function Get-GroupBySID {
    param([string]$GroupSID)
    try {
        $oGroup = Get-ADGroup -Filter {objectSid -eq $GroupSID} -properties * -ResultSetSize $null
    }
    catch {
        return $false
    }
    return $oGroup
}

function Get-UserBySamAccountName {
    param([string]$UserName)

    # $oUser = get-user $UserName -ResultSize Unlimited

    try {
        $oUser = Get-ADUser -Filter {sAMAccountName -eq $UserName} -properties * -ResultSetSize $null
    }
    catch {
        return $false
    }

    return $oUser
}

function Get-GroupMemberNamesRecursive {
    param($Group)

    [array]$GroupMembers = @()

    $GroupMembers = Get-ADGroupMember -Identity $Group.SamAccountName -Recursive
    # Write-Host ($GroupMembers | Format-Table | Out-String)

    return $GroupMembers
}

function Get-AllSharedMailboxes {
    $SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
    return $SharedMailboxes
}

function Get-MailboxFAGroups {
    param($Mailbox)
    [array]$MailboxUsersExch = @()
    # Get Full access permission users excluding SELF, inherited and denied users
    $MailboxUsersExch = Get-MailboxPermission -Identity $Mailbox | Where-Object {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false -and ($_.AccessRights -like "*FullAccess*") -and $_.Deny -eq $false}

    [array]$MailboxGroups = @()
    foreach ($MailboxUserExch in $MailboxUsersExch) {

        $MailboxGroup = Get-GroupBySID -GroupSID $MailboxUserExch.User.SecurityIdentifier.Value
        if ($MailboxGroup) {
            $isNotBlackListed = BlackListFilterForGroups -Group $MailboxGroup
            if ($isNotBlackListed) {
                $MailboxGroups += $MailboxGroup
            }
        }

        # Write-Host ($MailboxGroup | Format-List | Out-String)

    }
    return $MailboxGroups
}

function BlackListFilterForGroups {
    param($Group)
    [array]$BlackListedNames = @("Domain Admins", "Enterprise Admins")
    foreach ($blName in $BlackListedNames) {
        if ($Group.SamAccountName -eq $blName) {
            return $false
        }
    }
    return $true
}

function Get-MailboxSecurityGroupMembers {
    param($Mailbox)
    $MailboxFAGroups = Get-MailboxFAGroups -Mailbox $Mailbox

    [array]$GroupMembers = @()
    foreach ($Group in $MailboxFAGroups) {
        $GroupMembers += Get-GroupMemberNamesRecursive -Group $Group
    }
    # Remove redundancies
    $GroupMembers = $GroupMembers | Select-Object -Unique

    return $GroupMembers
}

function Set-MailboxAutoMapping {
    param($Mailbox, $UserDistinguishedName, $IsAutomapEnabled)

    # Test purposes
    $MailboxName = $Mailbox.SamAccountName

    if ($IsAutomapEnabled) {
        try {
            #### Comment this line to not actually do anything, just to see how it will behave
            Set-ADUser -Identity $Mailbox -Add @{msExchDelegateListLink=$UserDistinguishedName}
            #### Uncomment this line to see in the logs what (would have) happened
            ## Write-Host "Set-ADUser -Identity $MailboxName -Add @{msExchDelegateListLink=$UserDistinguishedName}"
            # Write-Host "$MailboxName -Add $UserDistinguishedName"
            WriteLog "$MailboxName -Add $UserDistinguishedName"
        }
        catch { return $false }
    } else {
        try {
            # Comment this line to not actually do anything, just to see how it will behave
            Set-ADUser -Identity $Mailbox -Remove @{msExchDelegateListLink=$UserDistinguishedName}
            #### Uncomment this line to see in the logs what (would have) happened
            ## Write-Host "Set-ADUser -Identity $MailboxName -Remove @{msExchDelegateListLink=$UserDistinguishedName}"
            # Write-Host "$MailboxName -Remove $UserDistinguishedName"
            WriteLog "$MailboxName -Remove $UserDistinguishedName"
        }
        catch { return $false }
    }
}

function CompareAccessToAutomapping {
    param($Mailbox, $MailboxSecurityGroupMembers)

    WriteLog -logstring "msExchDelegateListLink attribute content:"
    WriteLog -logstring ($Mailbox.msExchDelegateListLink | Format-Table | Out-String)
    WriteLog "`r`n`r`n"
    WriteLog -logstring "Mailbox Security Group Members: `r`n"
    WriteLog -logstring ($MailboxSecurityGroupMembers.distinguishedName | Format-Table | Out-String)
    WriteLog "`r`n`r`n"

    $AutomappedMembers=@()
    $AutomappedMembers = $Mailbox.msExchDelegateListLink

    # Write-Host "Actions:`r`n"
    WriteLog "Actions:`r`n"

    # Removing unneccesary outlook automaps
    foreach ($AutomappedMember in $AutomappedMembers) {
        $isFound = $false
        foreach ($MailboxSecurityGroupMember in $MailboxSecurityGroupMembers) {
            if ($MailboxSecurityGroupMember.distinguishedName -eq $AutomappedMember) {
                $isFound = $true
                break
            }
        }
        if ($isFound -eq $false) {
            Set-MailboxAutoMapping -Mailbox $Mailbox -UserDistinguishedName $AutomappedMember -IsAutomapEnabled $false
        }
    }

    # Adding missing outlook automaps
    foreach ($MailboxSecurityGroupMember in $MailboxSecurityGroupMembers) {
        $isFound = $false
        foreach ($AutomappedMember in $AutomappedMembers) {
            if ($MailboxSecurityGroupMember.distinguishedName -eq $AutomappedMember) {
                $isFound = $true
                break
            }
        }
        if ($isFound -eq $false) {
            Set-MailboxAutoMapping -Mailbox $Mailbox -UserDistinguishedName $MailboxSecurityGroupMember.distinguishedName -IsAutomapEnabled $true
        }
    }

}

function RunProgramForSingleMailbox {
    param($SharedMailbox)

    # Write-Host ($SharedMailbox.SamAccountName | Format-List | Out-String)
    WriteLog $SharedMailbox

    $GroupMembers = Get-MailboxSecurityGroupMembers -Mailbox $SharedMailbox
    # Write-Host ($GroupMembers | Format-Table | Out-String)

    $SharedMailboxUser = Get-UserBySamAccountName -UserName $SharedMailbox.SamAccountName

    CompareAccessToAutomapping -Mailbox $SharedMailboxUser -MailboxSecurityGroupMembers $GroupMembers
}

function RunProgramForAllMailboxes {
    param($AllSharedMailboxes)

    foreach ($SharedMailbox in $AllSharedMailboxes) {
        # Write-Host "----------------------"
        # Write-Host "`r`n"
        WriteLog "--------------------------------------"
        WriteLog "`r`n"

        RunProgramForSingleMailbox $SharedMailbox

        # Write-Host "`r`n`r`n"
        WriteLog "`r`n`r`n"
    }
}



function Main {
    SetUpLogs

    $AllSharedMailboxes = Get-AllSharedMailboxes
    # Write-Host ($AllSharedMailboxes | Format-Table | Out-String)
    WriteLog "Shared Mailbox List`r`n"
    WriteLog ($AllSharedMailboxes | Format-Table -Wrap -AutoSize | Out-String)

    RunProgramForAllMailboxes $AllSharedMailboxes

    ## This is for testing only on a single mailbox
    # RunProgramForSingleMailbox $AllSharedMailboxes[2]

}

# Start running our program
Main
