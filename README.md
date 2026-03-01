# PowerShell Training — Exchange Online Calendar Sharing Automation

## Project 1: Set-CalendarSharing.ps1

Automates calendar permission management across all Microsoft 365 mailboxes via Exchange Online PowerShell. Excludes a configurable group of executives. Designed for monthly scheduled production runs.

### Features

- Bulk calendar permission setting with a single command
- Executive exclusion list (CEO + named execs)
- `-WhatIf` simulation mode — preview all changes before applying
- Per-mailbox try/catch — one failure does not stop the run
- `trap {}` guardrail — catches unexpected crashes, emails admin, exits cleanly
- Post-run verification — re-reads permissions to confirm they applied
- Exclusion audit — confirms exec calendars were untouched
- Timestamped `.log` file and `.csv` report per run
- Email notification on success, partial failure, or crash
- Production-ready: supports Certificate-Based Authentication (CBA) for unattended scheduled runs

### Repository Structure

```
PowerShell_training/
├── README.md                          ← This file
├── scripts/
│   └── Set-CalendarSharing.ps1        ← Main production script
├── docs/
│   ├── Training_Guide_CalendarSharing.md   ← VAK training guide
│   └── Set-CalendarSharing_README.md       ← Quick-start usage guide
├── Logs/                              ← Runtime logs (git-ignored)
└── .gitignore
```

### Quick Start

```powershell
# 1. Simulate (no changes)
.\scripts\Set-CalendarSharing.ps1 -WhatIf

# 2. Live run with email notification
.\scripts\Set-CalendarSharing.ps1 -NotifyEmail "admin@contoso.com"

# 3. Custom permission level
.\scripts\Set-CalendarSharing.ps1 -SharingPermission "AvailabilityOnly"
```

### Prerequisites

| Requirement | Details |
|---|---|
| PowerShell | 5.1 or 7+ |
| Module | `ExchangeOnlineManagement` 3.x |
| Role | Exchange Administrator or Global Administrator |
| Auth (scheduled) | Certificate-Based Authentication — see training guide Section 8 |

### Documentation

- [Training Guide (VAK — Intermediate/Advanced)](docs/Training_Guide_CalendarSharing.md)
- [Quick-Start README](docs/Set-CalendarSharing_README.md)

### Official References

- [Exchange Online PowerShell — Microsoft Learn](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell?view=exchange-ps)
- [App-only authentication for unattended scripts](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps)
- [Add-MailboxFolderPermission](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/add-mailboxfolderpermission?view=exchange-ps)
- [Set-MailboxFolderPermission](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailboxfolderpermission?view=exchange-ps)
- [AZ-040T00: Automate Administration with PowerShell](https://learn.microsoft.com/en-us/training/courses/az-040t00)
