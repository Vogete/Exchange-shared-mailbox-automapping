# Exchange shared mailbox automapping

Managing Exchange shared mailbox access can be tedious and annoying because from the UI. It's small, limited, and it's just "yet another place to check for permissions". While you can add Active Directory security groups as delegates to the mailbox (and it works just fine), you are losing one feature that might be very important for your users: Automapping.

Automapping is when your Outlook client automatically opens all shared mailboxes you have permissions to. This feature is very convenient for users who don't want to worry about following guides to open several mailboxes. Unfortunately this feature is only available if the user is directly added as a mailbox delegate in the Exchange Server or Exchange Online admin center. So as an admin, you have to choose easy administration (security groups in Active Directory), or easy of use for users. 

_Why didn't Microsoft made both of these features available at the same time? Well, there is a very good explanation to that. And that is.........that is.......that is a very good question._

These scripts are the answer to solve this problem. You can use Active Directory security groups to grant permissions to shared mailboxes (centralizing administration), and you also get the automapping benefit for your users.

I made 2 scripts (our company currently use both of them in production for multiple years now with no issues), one for Exchange Server (on-premises) and one for Exchange Online (Office 365). Both versions use Active Directory to fetch security group memberships, but they can be modified quite easily to use Azure AD security groups.

_Disclaimer: these were one of my very first PowerShell projects, so the code is far from PowerShell standard. They work fine (both of them are used in live environments right now), and it was not worth the time to rewrite it in more professional manner, because it wouldn't be able to justify the many hours of work for basically zero operational benefits._

## Requirements

Both scripts needs to be ran on a domain joined Windows machine (Server or Desktop), because of the Active Directory integration. It also needs to have the [Active Directory module](https://docs.microsoft.com/en-us/powershell/module/activedirectory) installed (on non-server Windows versions, this is achieved by the Remove Server Administration Tools (RSAT)).

The Exchange Server version has to be run on Exchange Management Shell (EMS) with PowerShell 4 or 5.
The Exchange Online version can run on Powershell 5 or 7, and you need to have [Exchange Online PowerShell V2](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2) installed.

They can be used together (which is what I'm doing), they won't interfere with each other.

## Usage

The 2 scripts has different config and operational needs, due to the nature of their behavior. The self-hosted Exchange Server version is simply applying the `msExchDelegateListLink` attribute to each shared mailbox in Active Directory (based on security group membership), which will activate Outlook automapping. The Exchange Online version is directly adding (and removing) users in Exchange Online to (and from) the shared mailbox delegates list. These two key behaviors are that make their setup and operation different.

Please note that if you start using these scripts as scheduled scripts, you will have to use Acitve Directory to manage shared mailboxes, because the scripts will just overwrite everything that is not "documented" in AD. Basically AD becomes the source of truth for shared mailbox access!

### Exchange Server

This script will require less configuration. You need to just simply add your Active Directory security group once to the shared mailbox with the following two EMS commands:

```powershell
# Add Full access permissions
Add-MailboxPermission -Identity 'sharedmailbox@yourdomain.com' -User 'YOURDOMAIN\Shared Mailbox Access Group' -AccessRights 'FullAccess'
# Add Send-as permission
Add-ADPermission -Identity 'Shared Mailbox' -User 'YOURDOMAIN\Shared Mailbox Access Group' -ExtendedRights 'Send-as'
```

_Alternatively, you can also do this from the Exchange Control Panel UI, there is really no difference._

After this, the script is basically on autopilot, it should work out of box. If you want to regularly run the script (to keep up the permission sync), you should create a Windows Scheduled Task that runs the script from Exchange Management Shell.

_Pro tip: If you run this script from your Exchange Server Windows server, it should have all requirements already, so it should work out of the box._

### Exchange Online

This is a little bit trickier to set up, but not by much. You don't need EMS anymore, only PowerShell 5 or 7, and the [Exchange Online PowerShell V2](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2) module. However, there is a configuration file that you need to set. 

First you'll need an Office 365 user that has at least `Recipient Management` role in Exchange Online, otherwise it can't modify the Exchange objects (add/remove members)! You'll need to enter its UserPrincipalName (`o365User`) and password (`o365UserPW`) to the `config.json` file.

Then you'll need to decide on an Active Directory naming convention that you should use for the security groups that will grant users access to the mailboxes. This is necessary to clearly and easily identify what security group has access to which mailbox (without having to store it in a key-value list or database). The template which the script is using is the following:

```
[Prefix][Mailbox Name][Suffix]
```

The Prefix (`mailboxGroupPrefix`) and Suffix (`mailboxGroupPostfix`) should be set from the config file, while the Mailbox Name is the shared mailbox's display name (not username!). For example, for a `Sales Inquires` shared mailbox using the example config below, you'll need to have a `MailAccess Sales Inquires Mailbox` security group created. 

`config.json` template:

```json
{
    "o365User": "user@domain.com",
    "o365UserPW": "password",
    "mailboxGroupPrefix": "MailAccess ",
    "mailboxGroupPostfix": " Mailbox"
}
```

_Yes, I know it's called suffix, not postfix, but I was lazy to fix it from the first time._

Once you configure everything, create the necessary Active Directory groups, and Exchange Online mailboxes (and of course add the users), the script should be doing its job without interferance.

I might remake this script to use Office 365 security groups in the future (should be fairly straightforward), but until then, it is what it is.