<#
.SYNOPSIS
Share calendars from one Department to another in Microsoft 365.

- All users in the FROM Department share their calendars
- All users in the TO Department receive a specified permission level
- Uses Microsoft Graph to find users and Exchange Online to set calendar permissions

REQUIREMENTS:
- Module: Microsoft.Graph.Users
- Module: ExchangeOnlineManagement
- Permissions: You must be allowed to read users in Graph and manage mailbox folder permissions in Exchange Online.
#>

# Load required modules
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module ExchangeOnlineManagement -ErrorAction Stop

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All"

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline

# Ask for Departments
$fromDept = Read-Host "Enter the EXACT Department name whose calendars will be shared (FROM)"
$toDept   = Read-Host "Enter the EXACT Department name that should receive calendar access (TO)"

# Ask for calendar folder name, default is 'Kalender'
$calendarFolderName = Read-Host "Enter the calendar folder name (press Enter for default 'Kalender')"
if ([string]::IsNullOrWhiteSpace($calendarFolderName)) {
    $calendarFolderName = "Kalender"
}

# Choose permission level
Write-Host ""
Write-Host "Select the calendar permission level to grant to the TO department:" -ForegroundColor Cyan
Write-Host "  1 = LimitedDetails (can see time/subject/location, cannot edit)"
Write-Host "  2 = Reviewer      (can see full details, cannot edit)"
Write-Host "  3 = Editor        (can see full details and can edit)"
$permChoice = Read-Host "Enter 1, 2, or 3 (default is 1)"

switch ($permChoice) {
    "2" { $permissionToGrant = "Reviewer" }
    "3" { $permissionToGrant = "Editor" }
    default { $permissionToGrant = "LimitedDetails" }
}

Write-Host ""
Write-Host "Fetching users from Microsoft Graph..." -ForegroundColor Cyan

# Get users in FROM Department (their calendars will be shared)
$fromUsers = Get-MgUser -Filter "department eq '$fromDept'" -All |
    Where-Object { $_.Mail } |
    Select-Object DisplayName, Mail, Department

# Get users in TO Department (they receive permissions)
$toUsers = Get-MgUser -Filter "department eq '$toDept'" -All |
    Where-Object { $_.Mail } |
    Select-Object DisplayName, Mail, Department

Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor DarkGray
Write-Host "FROM Department: $fromDept" -ForegroundColor Yellow
Write-Host "Number of users (calendars to share): $($fromUsers.Count)"
$fromUsers | Format-Table -AutoSize

Write-Host ""
Write-Host "TO Department:   $toDept" -ForegroundColor Yellow
Write-Host "Number of users (will receive access): $($toUsers.Count)"
$toUsers | Format-Table -AutoSize
Write-Host "----------------------------------------" -ForegroundColor DarkGray
Write-Host ""

if ($fromUsers.Count -eq 0 -or $toUsers.Count -eq 0) {
    Write-Warning "Either the FROM or TO Department returned 0 users. Check the department names and try again."
    return
}

Write-Host "Calendar folder to use: '$calendarFolderName'" -ForegroundColor Cyan
Write-Host "Permission level to grant: '$permissionToGrant'" -ForegroundColor Cyan
Write-Host ""

$confirm = Read-Host "Proceed with applying '$permissionToGrant' permissions on folder '$calendarFolderName' from all FROM users to all TO users? (y/n)"
if ($confirm -ne 'y' -and $confirm -ne 'Y') {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    return
}

Write-Host ""
Write-Host "Applying permissions. This may take some time if there are many users..." -ForegroundColor Cyan

foreach ($source in $fromUsers) {
    $calendarIdentity = "$($source.Mail):\$calendarFolderName"

    Write-Host ""
    Write-Host "Processing calendar: $calendarIdentity" -ForegroundColor Green

    foreach ($target in $toUsers) {
        try {
            # Check if a permission entry already exists for this user
            $existing = Get-MailboxFolderPermission -Identity $calendarIdentity -ErrorAction Stop |
                Where-Object { $_.User -eq $target.Mail }

            if (-not $existing) {
                # Add new permission
                Add-MailboxFolderPermission -Identity $calendarIdentity -User $target.Mail -AccessRights $permissionToGrant -ErrorAction Stop
                Write-Host "  Added $permissionToGrant for $($target.Mail)"
            }
            else {
                Write-Host "  Skipped $($target.Mail) (already has: $($existing.AccessRights))"
            }
        }
        catch {
            Write-Warning "  Error for $($target.Mail) on ${calendarIdentity}: $($_.Exception.Message)"
        }
    }
}


Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
Write-Host "All users in '$toDept' now have at least '$permissionToGrant' on the '$calendarFolderName' folder for all users in '$fromDept'."
