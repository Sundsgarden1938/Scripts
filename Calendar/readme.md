# README – Group-Based Calendar Permission Assignment

This script assigns calendar permissions for all user mailboxes that are members of a specified *source* group.  
Each user's calendar folder (default name: **"Kalender"**) is shared with a specified *target* group using the access right **LimitedDetails** (or another access right if defined at runtime).

## What the script does
1. Prompts for the source group SMTP address.
2. Prompts for the target group SMTP address.
3. Resolves both groups in Exchange Online.
4. Retrieves all members of the source group who have user mailboxes.
5. Applies the chosen calendar permission (default: **LimitedDetails**) to the calendar folder of each user mailbox.

## Requirements
- Ability to run `Connect-ExchangeOnline`.
- Both source and target groups must be **mail-enabled** recipients in Exchange Online.
- PowerShell 7 is recommended but not strictly required.
- You must have sufficient permissions to modify mailbox folder permissions.

## Test Mode
You can safely test the script without applying changes by adding `-WhatIf` to the lines containing:

