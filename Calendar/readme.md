# README – Department-Based Calendar Permission Assignment

This script assigns calendar permissions between two Microsoft 365 departments.  
Users in the **FROM Department** share their calendar folders (default name: **"Kalender"**) with all users in the **TO Department**, using a permission level selected at runtime (**LimitedDetails**, **Reviewer**, or **Editor**).

## What the script does
1. Prompts for the **FROM Department** (the department whose users will share their calendars).
2. Prompts for the **TO Department** (the department that will receive access).
3. Prompts for the calendar folder name (default: **"Kalender"**).
4. Prompts for the permission level to assign.
5. Connects to:
   - Microsoft Graph (to retrieve users based on Department)
   - Exchange Online (to set calendar permissions)
6. Retrieves all users in the FROM Department.
7. Retrieves all users in the TO Department.
8. Applies the selected calendar permission from each FROM user’s calendar to every user in the TO Department.
9. Skips users who already have the assigned permission.

## Requirements
- PowerShell modules:
  - `Microsoft.Graph.Users`
  - `ExchangeOnlineManagement`
- Ability to run:
  - `Connect-MgGraph -Scopes "User.Read.All"`
  - `Connect-ExchangeOnline`
- You must have sufficient permissions to modify mailbox folder permissions.
- Department values must match exactly what is stored in Microsoft Entra ID.
- All users must have a primary SMTP address.

## Calendar Permission Levels
| Level           | View details | Edit items | Notes |
|-----------------|--------------|------------|--------|
| LimitedDetails  | Partial      | No         | Time, subject, location only |
| Reviewer        | Full         | No         | Full read-only access |
| Editor          | Full         | Yes        | Can create and modify events |

## How permissions are applied
For every user mailbox in the FROM Department:
- The script constructs the target calendar folder path, e.g.:  
  `user@example.com:\Kalender`
- Each TO Department user is granted the selected access right unless already present.

## Test Mode
You can safely simulate the script's behavior without making changes by adding `-WhatIf` to the following cmdlets:
- `Add-MailboxFolderPermission`
- `Get-MailboxFolderPermission`

Example:
```powershell
Add-MailboxFolderPermission -Identity $calendarIdentity -User $target.Mail -AccessRights $permissionToGrant -WhatIf
