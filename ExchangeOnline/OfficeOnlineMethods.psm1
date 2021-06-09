function Get-O365Credentials {
    param (
        [string]$username,
        [string]$password
    )
    $secureStringPwd = ConvertTo-SecureString -String $password -AsPlainText -Force
    $creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $secureStringPwd
    return $creds
}

function Get-AllSharedMailboxes {
    $SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
    return $SharedMailboxes
}

function Connect-Office365 {
    param (
        [System.Management.Automation.PSCredential]$credentials
    )
    try {
        # EXO V1 - DO NOT USE
        # # Does not support 2FA!
        # # $global:ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirect
        # $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirect
        # # Import-PSSession $global:ExchangeOnlineSession -DisableNameChecking -AllowClobber
        # Import-PSSession $Session -DisableNameChecking -AllowClobber

        # EX Online V2 - USE THIS
        Connect-ExchangeOnline -Credential $credentials -ShowProgress $False
    }
    catch {
        return $false;
    }

    # return $Session
}

function Get-FullAccessPermissionForMailbox {
    param (
        [Object]$SharedMailbox
    )

    try {
        $permissions = Get-MailboxPermission -Identity $SharedMailbox.alias | Where { ($_.IsInherited -eq $False) -and -not ($_.User -like "NT AUTHORITY\SELF") }
    }
    catch {
        return $false;
    }

    return $permissions
}


function Add-FullAccessPermissionForMailbox {
    param (
        [Object]$SharedMailbox,
        [Microsoft.ActiveDirectory.Management.ADUser]$User
    )

    try {
        Add-MailboxPermission -Identity $SharedMailbox.alias -User $User.sAMAccountName -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$false
    }
    catch {
        return $false;
    }

    return $true
}

function Remove-FullAccessPermissionForMailbox {
    param (
        [Object]$SharedMailbox,
        [Microsoft.ActiveDirectory.Management.ADUser]$User
    )

    try {
        Remove-MailboxPermission -Identity $SharedMailbox.alias -User $User.sAMAccountName -AccessRights FullAccess -InheritanceType All -Confirm:$false
    }
    catch {
        return $false;
    }

    return $true
}

function Get-SendAsPermissionForMailbox {
    param (
        [Object]$SharedMailbox
    )

    try {
        $permissions = Get-RecipientPermission -Identity $SharedMailbox.alias | Where { ($_.IsInherited -eq $False) -and -not ($_.Trustee -like "NT AUTHORITY\SELF") }
        # $permissions = Get-Mailbox -Identity $SharedMailbox.alias -ResultSize Unlimited | Get-RecipientPermission | ? {$_.Trustee -ne "NT AUTHORITY\SELF"}
    }
    catch {
        return $false;
    }

    return $permissions
}


function Add-SendAsPermissionForMailbox {
    param (
        [Object]$SharedMailbox,
        [Microsoft.ActiveDirectory.Management.ADUser]$User
    )

    try {
        Add-RecipientPermission -Identity $SharedMailbox.alias -AccessRights SendAs -Trustee $User.sAMAccountName -Confirm:$false
    }
    catch {
        return $false
    }

    return $true
}

function Remove-SendAsPermissionForMailbox {
    param (
        [Object]$SharedMailbox,
        [Microsoft.ActiveDirectory.Management.ADUser]$User
    )

    try {
        Remove-RecipientPermission -Identity $SharedMailbox.alias -AccessRights SendAs -Trustee $User.sAMAccountName -Confirm:$false
    }
    catch {
        return $false
    }

    return $true
}
