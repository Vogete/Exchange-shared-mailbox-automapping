Using module "./SyncModel.psm1"

Using module "./Logging.psm1"
Using module "./HelperMethods.psm1"
Using module "./OfficeOnlineMethods.psm1"
Using module "./ADMethods.psm1"

function Sync-DelegatesForMailbox {
    param (
        [Object]$SharedMailbox,
        [Microsoft.ActiveDirectory.Management.ADGroup]$PermissionGroup
    )
    Write-Host $SharedMailbox
    Write-Host $permissionGroup + "`r`n"
    WriteLog -_logstring ("Mailbox: " + $SharedMailbox)
    WriteLog -_logstring ("AD Access Group: " + $permissionGroup)

    $ADMembers = Get-RecursiveADGroupMembers -Group $PermissionGroup
    $ExchangeFullAccessMembers = Get-FullAccessPermissionForMailbox -SharedMailbox $SharedMailbox # Obj.User
    $ExchangeSendAsMembers = Get-SendAsPermissionForMailbox -SharedMailbox $SharedMailbox # Obj.Trustee

    # Convert that shitty list into ADObject lists so it's easier to work with it
    $ExchangeFullAccessADObjects = [System.Collections.ArrayList]@()
    foreach ($FullAccessMember in $ExchangeFullAccessMembers) {
        $FullAccessObject = Get-ADObjectByUserPrincipalName -Username $FullAccessMember.User
        $ExchangeFullAccessADObjects.Add($FullAccessObject)
    }
    # $ExchangeSendAsUsernames = [System.Collections.ArrayList]@()
    # foreach ($SendAsMember in $ExchangeSendAsMembers) {
    #     $SendAsObject = Get-ADObjectByUserPrincipalName -Username $SendAsMember.Trustee
    #     $ExchangeSendAsUsernames.Add($SendAsObject)
    # }

    # We only check for Full Access, since if they have that, they have SendAs perm as well
    $diffFullAccess = Compare-ExchangeADDifferences -ADGroupMembers $ADMembers -MailboxDelegates $ExchangeFullAccessADObjects

    foreach ($ObjectToRemove in $diffFullAccess.ObjectsToRemove) {
        if ($ObjectToRemove.sAMAccountName.length -le 0) {Continue}
        Write-Host $ObjectToRemove.sAMAccountName "is removed"
        WriteLog -_logstring ($ObjectToRemove.sAMAccountName + " is removed")
        Remove-PermissionsForMailbox -SharedMailbox $SharedMailbox -Username $ObjectToRemove.sAMAccountName
    }

    foreach ($ObjectToAdd in $diffFullAccess.ObjectsToAdd) {
        if ($ObjectToAdd.sAMAccountName.length -le 0) {Continue}
        Write-Host $ObjectToAdd.sAMAccountName "is Added"
        WriteLog -_logstring ($ObjectToAdd.sAMAccountName + " is added")
        Add-PermissionsForMailbox -SharedMailbox $SharedMailbox -Username $ObjectToAdd.sAMAccountName

    }

    WriteLog -_logstring "`r`n"
}

function Compare-ExchangeADDifferences {
    param (
        $ADGroupMembers,
        $MailboxDelegates
    )
    [SyncModel]$Differences = [SyncModel]::New()

    foreach ($ADGroupMember in $ADGroupMembers) {
        $shouldBeAdded = $true
        foreach ($MailboxDelegate in $MailboxDelegates) {
            if ($ADGroupMember.ObjectGuid -eq $MailboxDelegate.ObjectGuid) {
                $shouldBeAdded = $false;
                break;
            }
        }
        if ($shouldBeAdded) {
            $Differences.ObjectsToAdd.Add($ADGroupMember);
        }
    }

    foreach ($MailboxDelegate in $MailboxDelegates) {
        $shouldBeRemoved = $true
        foreach ($ADGroupMember in $ADGroupMembers) {
            if ($ADGroupMember.ObjectGuid -eq $MailboxDelegate.ObjectGuid) {
                $shouldBeRemoved = $false;
                break;
            }
        }
        if ($shouldBeRemoved) {
            $Differences.ObjectsToRemove.Add($MailboxDelegate);
        }
    }

    return $Differences;
}

function Add-PermissionsForMailbox {
    param (
        [Object]$SharedMailbox,
        [string]$Username
    )
    $User = Get-UserBySamAccountName -Username $Username

    Add-FullAccessPermissionForMailbox -SharedMailbox $SharedMailbox -User $User
    Add-SendAsPermissionForMailbox -SharedMailbox $SharedMailbox -User $User
}
function Remove-PermissionsForMailbox {
    param (
        [Object]$SharedMailbox,
        [string]$Username
    )
    $User = Get-UserBySamAccountName -Username $Username

    Remove-FullAccessPermissionForMailbox -SharedMailbox $SharedMailbox -User $User
    Remove-SendAsPermissionForMailbox -SharedMailbox $SharedMailbox -User $User
}

function Main {
    # Initialization
    SetUpLogs
    $global:LogFileName = $LogFolderName
    $global:LogFileName += Get-Date -Format o | foreach {$_ -replace ":", "."}
    $global:LogFileName += ".log"

    $config = Read-ConfigFile
    $credentials = Get-O365Credentials -username $config.o365User -password $config.o365UserPW
    $O365Session = Connect-Office365 -credentials $credentials
    if ($O365Session -eq $false) {return $false}
    # ######

    $SharedMailboxes = Get-AllSharedMailboxes

    # $testmailbox = $SharedMailboxes[0]
    # $PermissionGroup = Get-GroupForSharedMailbox -SharedMailbox $testmailbox -Prefix $config.mailboxGroupPrefix -Postfix $config.mailboxGroupPostfix
    # Sync-DelegatesForMailbox -SharedMailbox $testmailbox -PermissionGroup $PermissionGroup

    foreach ($SharedMailbox in $SharedMailboxes) {
        $PermissionGroup = Get-GroupForSharedMailbox -SharedMailbox $SharedMailbox -Prefix $config.mailboxGroupPrefix -Postfix $config.mailboxGroupPostfix
        Sync-DelegatesForMailbox -SharedMailbox $SharedMailbox -PermissionGroup $PermissionGroup
    }


    WriteLogToFile $global:LogFileName $global:Log
}

# Clear PS Modules:
#      Remove-Module *;

Main
