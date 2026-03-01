# Set-CalendarSharing.ps1 — Usage Guide

## Purpose
Sets calendar sharing permissions across all M365 user mailboxes via Exchange Online,
while explicitly excluding the CEO and a configurable group of executives.

---

## Prerequisites

| Requirement | Details |
|---|---|
| PowerShell | 5.1+ or PowerShell 7+ |
| Module | `ExchangeOnlineManagement` (auto-installed if missing) |
| Admin Role | Exchange Administrator or Global Administrator |
| Auth | Modern Auth / MFA-compatible (interactive browser login) |

---

## Configuration (edit before first run)

Open `Set-CalendarSharing.ps1` and update these two sections:

### 1. Exclusion List
```powershell
$ExcludedUsers = @(
    "ceo@contoso.com",
    "cfo@contoso.com",
    "coo@contoso.com"
    # Add or remove executives here
)
```

### 2. Notification Settings
```powershell
$SmtpServer  = "smtp.office365.com"
$SenderEmail = "it-admin@contoso.com"
```

---

## How to Run

### Step 1 — Simulate first (always recommended)
```powershell
.\Set-CalendarSharing.ps1 -WhatIf
```
Shows exactly what would change. No mailboxes are modified.

### Step 2 — Live run with email notification
```powershell
.\Set-CalendarSharing.ps1 -NotifyEmail "admin@contoso.com"
```

### Step 3 — Change permission level
```powershell
.\Set-CalendarSharing.ps1 -SharingPermission "AvailabilityOnly" -NotifyEmail "admin@contoso.com"
```

### Available permission levels
| Value | What it shares |
|---|---|
| `AvailabilityOnly` | Free/Busy only |
| `LimitedDetails` | Free/Busy + subject & location (default) |
| `Reviewer` | Full read access to calendar items |
| `Editor` | Read and write access |
| `Author` | Read, write, create |
| `PublishingEditor` | Full control |

---

## What the Script Does — Step by Step

1. **Checks prerequisites** — Verifies/installs the Exchange Online module
2. **Connects** — Interactive modern auth login (MFA-supported)
3. **Retrieves mailboxes** — `Get-Mailbox -RecipientTypeDetails UserMailbox`
4. **Partitions** — Separates target users from excluded executives
5. **Validates exclusions** — Warns if any configured exclusion UPN is not found in Exchange
6. **Applies permissions** — `Add-MailboxFolderPermission` or `Set-MailboxFolderPermission`
7. **Verifies** — Re-reads permissions post-change to confirm they applied
8. **Exclusion audit** — Confirms exec calendars were not touched
9. **Reports** — Writes a `.log` file and a `.csv` report
10. **Notifies** — Emails the admin with a summary (if `-NotifyEmail` provided)
11. **Disconnects** — Cleanly closes the Exchange Online session

---

## Guardrails & Failure Handling

| Scenario | Behaviour |
|---|---|
| Module missing | Auto-installs; exits with error + email if it fails |
| Cannot connect to Exchange | Exits with error + email notification |
| `Get-Mailbox` fails | Exits with error + email notification |
| Per-mailbox permission failure | Logged as ERROR; script continues to next mailbox |
| Unexpected script crash | `trap {}` block catches it, logs, emails admin, disconnects |
| Exclusion UPN not found in Exchange | Logs a WARN (does not abort) |
| Permission mismatch after apply | Logs a WARN during verification pass |

---

## Output Files

All output is written to a `Logs\` subfolder next to the script.

| File | Contents |
|---|---|
| `CalendarSharing_YYYYMMDD_HHmmss.log` | Full timestamped run log |
| `CalendarSharing_Report_YYYYMMDD_HHmmss.csv` | Per-mailbox result table |

---

## Troubleshooting

**"Connect-ExchangeOnline is not recognized"**
Run: `Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force`

**Script blocked by execution policy**
Run: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`

**Email notification not sending**
Confirm `$SenderEmail` is a valid M365 mailbox and your admin account has Send-As rights.
