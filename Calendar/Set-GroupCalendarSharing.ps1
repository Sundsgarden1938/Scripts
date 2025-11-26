<#
.SYNOPSIS
    Set calendar sharing from one group (source) to another (target).

.DESCRIPTION
    For all user mailboxes that are members of the *source* group, this script
    sets calendar permissions on the folder "Kalender" so that the *target* group
    gets the specified access right (LimitedDetails by default).

    The script:
    - Prompts for source group SMTP address
    - Prompts for target group SMTP address
    - Resolves the source and target groups in Exchange Online
    - Resolves members of the source group (user mailboxes only)
    - Applies calendar folder permissions on "Kalender" for each member

.NOTES
    Run this in an Exchange Online PowerShell session.
#>

# ================================
# Interactive configuration
# ================================

$sourceGroupSmtp = Read-Host "Enter SOURCE group SMTP address (whose members' calendars will be shared)"
$targetGroupSmtp = Read-Host "Enter TARGET group SMTP address (group that should get calendar access)"

if ([string]::IsNullOrWhiteSpace($sourceGroupSmtp) -or [string]::IsNullOrWhiteSpace($targetGroupSmtp)) {
    Write-Error "Source and target group SMTP addresses are required. Exiting."
    return
}

# Optional: make folder name and access right configurable with defaults
$calendarFolderName = Read-Host "Enter calendar folder name (press Enter for default 'Kalender')"
if ([string]::IsNullOrWhiteSpace($calendarFolderName)) {
    $calendarFolderName = "Kalender"
}

$accessRight = Read-Host "Enter access right (press Enter for default 'LimitedDetails')"
if ([string]::IsNullOrWhiteSpace($accessRight)) {
    $accessRight = "LimitedDetails"
}

Write-Host ""
Write-Host "Source group: $sourceGroupSmtp"
Write-Host "Target group: $targetGroupSmtp"
Write-Host "Calendar folder: $calendarFolderName"
Write-Host "Access right: $accessRight"
Write-Host ""

# ================================
# Connect to Exchange Online
# ================================
try {
    # Comment this out if you already connect before running the script
    Connect-ExchangeOnline -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to Exchange Online. Error: $($_.Exception.Message)"
    return
}

# ================================
# Resolve source and target groups
# ================================
try {
    $sourceGroup = Get-EXORecipient -Identity $sourceGroupSmtp -PropertySets All -ErrorAction Stop
}
catch {
    Write-Error "Could not resolve source group '$sourceGroupSmtp' in Exchange Online. Error: $($_.Exception.Message)"
    return
}

try {
    $targetGroup = Get-EXORecipient -Identity $targetGroupSmtp -PropertySets All -ErrorAction Stop
}
catch {
    Write-Error "Could not resolve target group '$targetGroupSmtp' in Exchange Online. Error: $($_.Exception.Message)"
    return
}

Write-Host "Resolved source group: $($sourceGroup.DisplayName) [$($sourceGroup.RecipientTypeDetails)]"
Write-Host "Resolved target group: $($targetGroup.DisplayName) [$($targetGroup.RecipientTypeDetails)]"
Write-Host ""

# ================================
# Get members of the source group
# ================================
$sourceMembers = @()

switch ($sourceGroup.RecipientTypeDetails) {
    "MailUniversalDistributionGroup" {
        Write-Host "Source group is a Distribution Group. Using Get-DistributionGroupMember..."
        $sourceMembers = Get-DistributionGroupMember -Identity $sourceGroupSmtp -ResultSize Unlimited |
                         Where-Object { $_.RecipientType -eq "UserMailbox" }
    }
    "MailUniversalSecurityGroup" {
        Write-Host "Source group is a Mail-enabled Security Group. Using Get-DistributionGroupMember..."
        $sourceMembers = Get-DistributionGroupMember -Identity $sourceGroupSmtp -ResultSize Unlimited |
                         Where-Object { $_.RecipientType -eq "UserMailbox" }
    }
    "GroupMailbox" {
        Write-Host "Source group is a Microsoft 365 Group. Using Get-UnifiedGroupLinks..."
        $sourceMembers = Get-UnifiedGroupLinks -Identity $sourceGroupSmtp -LinkType Members -ResultSize Unlimited |
                         Where-Object { $_.RecipientType -eq "UserMailbox" }
    }
    default {
        Write-Error "Unsupported source group type: $($sourceGroup.RecipientTypeDetails). Cannot continue."
        return
    }
}

if (-not $sourceMembers -or $sourceMembers.Count -eq 0) {
    Write-Host "No user mailboxes were found in source group '$sourceGroupSmtp'. Nothing to do."
    return
}

Write-Host "Found $($sourceMembers.Count) user mailbox(es) in source group."
Write-Host ""

# ================================
# Apply calendar permissions
# ================================
foreach ($user in $sourceMembers) {
    $calendarIdentity = "$($user.PrimarySmtpAddress):\$calendarFolderName"
    Write-Host "Setting $accessRight on $calendarIdentity for $targetGroupSmtp..."

    try {
        # Get all permissions and check if the target group already has an entry
        $allPerms = Get-MailboxFolderPermission -Identity $calendarIdentity -ErrorAction SilentlyContinue

        $existing = $allPerms | Where-Object {
            $_.User.ToString() -eq $targetGroupSmtp -or
            $_.User.ToString() -eq $targetGroup.DisplayName
        }

        if ($existing) {
            Write-Host "  -> Existing permission entry found for target group. Updating to '$accessRight' with Set-MailboxFolderPermission."
            Set-MailboxFolderPermission -Identity $calendarIdentity `
                -User $targetGroupSmtp `
                -AccessRights $accessRight
        }
        else {
            Write-Host "  -> No existing entry for target group. Adding new permission with Add-MailboxFolderPermission."
            Add-MailboxFolderPermission -Identity $calendarIdentity `
                -User $targetGroupSmtp `
                -AccessRights $accessRight
        }
    }
    catch {
        Write-Warning "  -> Failed to set permissions on $calendarIdentity. Error: $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "Done. All user mailboxes in '$sourceGroupSmtp' now share their '$calendarFolderName' folder with '$targetGroupSmtp' using access right: $accessRight."
