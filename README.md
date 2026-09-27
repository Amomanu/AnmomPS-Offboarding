# AnmomPS-Offboarding

PowerShell function library that offboards users across on-premises Active Directory, Entra ID (Azure AD), Exchange Online and Teams - one user at a time or in bulk from a CSV - with a CSV log of every step.

## What it does

For each user:

- **Active Directory** (synced users): disables the account, resets the password to a random value, removes AD group memberships, hides the user from the address book, clears description, title, department, company, ipPhone and manager, and moves the account to a Disabled Users OU.
- **Entra ID**: disables the account, revokes refresh tokens and disables registered devices (Cloud PCs are skipped).
- **Exchange Online**: optional auto-reply naming the manager, optional forwarding to the manager, 50 KB send limit, optional conversion to a shared mailbox.
- **Groups and Teams**: removes the user from Microsoft 365 groups and teams (owner and member), distribution lists and security groups, except one optional excluded group; optionally adds the user to a "disabled users" group.

## Usage

```powershell
. '.\Terminate_User_Script (1).ps1'
get-TerminationSessions
Terminate_User -userUPN user@contoso.com -exportLogPath C:\ExportLogs -ConvertToShared $true -setAutoReply $true -Forwarding $true
Terminate_UsersFromCSV -csvPath .\users.csv -exportLogPath C:\ExportLogs
```

CSV columns: UPN (required), DomainController, ManagerName, ManagerEmail, ConverToShared, SetAutoReply, ForwardEmail (Yes/No), ExcludedLicenseGroup, disabledGroupObjectID.

Full details and known issues: `Get-Help '.\Terminate_User_Script (1).ps1' -Full`

## Notes

- Requires the ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement and MicrosoftTeams modules.
- Before use, adapt the domain-to-OU mapping in `Add-ADUserToDisabledOU` and the organisation name in `SetAutoReply`.
- Organization names are placeholders.
