# Exchange Online Calendar Sharing Automation
## PowerShell Training Guide — Intermediate to Advanced

**Script:** `Set-CalendarSharing.ps1`
**Audience:** IT Administrators with working PowerShell knowledge
**Learning Style:** Visual · Kinesthetic
**Production Use:** Monthly scheduled run via Windows Task Scheduler or Azure Automation

---

## How to Use This Guide

This guide uses two learning styles. Both reinforce each other — work through them together for the best results.

| Symbol | Style | What it means for you |
|---|---|---|
| `[V]` | **Visual** | Stop and study the diagram or table before moving on. Let the picture do the explaining. |
| `[K]` | **Hands-On** | Close the guide, open your terminal, and type it yourself. Muscle memory matters more than reading about it. |

> **Human note:** There is no shortcut for the hands-on sections. The exercises are short by design — each one takes under five minutes. Do them as you go, not later.

---

## Table of Contents

1. [Background & Context](#1-background--context)
2. [Architecture Overview — Visual Map](#2-architecture-overview--visual-map)
3. [Prerequisites & Environment Setup](#3-prerequisites--environment-setup)
4. [Key Concepts Explained](#4-key-concepts-explained)
5. [Script Walkthrough — Section by Section](#5-script-walkthrough--section-by-section)
6. [Running the Script — Step-by-Step](#6-running-the-script--step-by-step)
7. [Simulation with -WhatIf](#7-simulation-with--whatif)
8. [Scheduling for Monthly Production Use](#8-scheduling-for-monthly-production-use)
9. [Guardrails & Failure Handling](#9-guardrails--failure-handling)
10. [Verification & Log Interpretation](#10-verification--log-interpretation)
11. [Knowledge Check](#11-knowledge-check)
12. [References & Official Resources](#12-references--official-resources)

---

## 1. Background & Context

### What problem does this script solve?

Your organisation needs all employee calendars visible to colleagues at a defined permission level — but the CEO and named executives must be excluded for privacy and governance reasons.

Doing this manually in the Microsoft 365 admin portal:
- Does not scale (hundreds or thousands of mailboxes)
- Leaves no audit trail
- Cannot easily be repeated monthly as new staff join

PowerShell with the **ExchangeOnlineManagement** module solves all three problems cleanly.

### The one-line summary of what this script does

> Connect to Exchange Online → get every mailbox → skip the executives → set the calendar permission → confirm it applied → log everything → report the result.

Every section in this guide is one of those steps in detail. Keep that sentence in your head as you work through the script.

---

## 2. Architecture Overview — Visual Map

### `[V]` What the script does from start to finish

```
┌─────────────────────────────────────────────────────────────┐
│                    ADMIN WORKSTATION                        │
│                                                             │
│  Set-CalendarSharing.ps1                                    │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ 1. Check module (ExchangeOnlineManagement)           │  │
│  │ 2. Connect-ExchangeOnline  (Modern Auth / CBA)       │  │
│  │ 3. Get-Mailbox  ──► All UserMailbox objects          │  │
│  │ 4. Filter ──► Remove $ExcludedUsers from list        │  │
│  │                                                      │  │
│  │   ┌────────────────┐   ┌──────────────────────────┐ │  │
│  │   │ TARGET USERS   │   │   EXCLUDED (Execs / CEO) │ │  │
│  │   │ Apply perm     │   │   Untouched              │ │  │
│  │   └──────┬─────────┘   └──────────────────────────┘ │  │
│  │          │                                           │  │
│  │ 5. Add/Set-MailboxFolderPermission  :\Calendar       │  │
│  │ 6. Verify Get-MailboxFolderPermission                │  │
│  │ 7. Write .log + .csv                                 │  │
│  │ 8. Send-MailMessage  ──► Admin email                 │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
          ┌───────────────────────────────┐
          │      EXCHANGE ONLINE (M365)   │
          │  Mailbox 1: user1@corp.com    │
          │  Mailbox 2: user2@corp.com    │
          │  ...                          │
          │  CEO mailbox  ─── SKIPPED     │
          └───────────────────────────────┘
```

### `[V]` Decision path for a single mailbox

```
Get-Mailbox ──► [Is UPN in $ExcludedSet?]
                       │
              YES ─────┴───── NO
               │               │
          Log EXCLUDED    Does "Default" entry
          Skip                 already exist?
                               │
                     YES ──────┴────── NO
                      │                │
              Set-MailboxFolder    Add-MailboxFolder
              Permission           Permission
                      │
              Verify with Get-MailboxFolderPermission
                      │
              Log SUCCESS or WARN
```

---

## 3. Prerequisites & Environment Setup

### `[K]` Type each of these commands yourself before going any further

#### Step 1 — Check your PowerShell version

```powershell
$PSVersionTable.PSVersion
# You need Major version 5 (Windows) or 7+ (cross-platform)
```

#### Step 2 — Install or update the Exchange Online module

```powershell
# Check if already installed
Get-Module -ListAvailable -Name ExchangeOnlineManagement

# Install if missing
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force

# Keep it current — run this periodically
Update-Module -Name ExchangeOnlineManagement
```

> **Current latest version:** 3.9.0 (as of early 2026)
> Source: [PowerShell Gallery — ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement)

#### Step 3 — Allow scripts to run in your session

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

#### Step 4 — Confirm you have the right admin role

You need one of these Microsoft Entra roles:
- **Exchange Administrator** — preferred, least privilege
- **Global Administrator** — for emergencies only, avoid for routine automation

Check your assigned roles in the M365 Admin Portal under **Roles → Role assignments**.

### `[V]` Everything you need at a glance

| Requirement | Minimum | How to check |
|---|---|---|
| PowerShell | 5.1 or 7+ | `$PSVersionTable.PSVersion` |
| ExchangeOnlineManagement | 3.x | `Get-Module ExchangeOnlineManagement -ListAvailable` |
| Admin Role | Exchange Administrator | M365 Admin Portal → Roles |
| Network | HTTPS port 443 outbound | `Test-NetConnection outlook.office365.com -Port 443` |
| Execution Policy | RemoteSigned | `Get-ExecutionPolicy -Scope CurrentUser` |

---

## 4. Key Concepts Explained

### 4.1 `[V]` Mailbox Folder Permissions vs Sharing Policies

There are three distinct Exchange sharing mechanisms. This script uses **Mailbox Folder Permissions** — the one that controls internal access to a specific folder inside a mailbox.

| Mechanism | Scope | Used for |
|---|---|---|
| `MailboxFolderPermission` | Per-folder inside a mailbox | Internal org calendar access — **this script** |
| `SharingPolicy` | Organisation-wide rule | External domain sharing |
| `OrganizationRelationship` | Federated tenants | Cross-company free/busy |

The script targets the `:\Calendar` folder using `Add-MailboxFolderPermission` and `Set-MailboxFolderPermission`.

---

### 4.2 `[V]` Permission levels — what each one reveals

Choose the level that matches your organisation's need. The default in this script is `LimitedDetails`.

| AccessRights value | What a colleague can see |
|---|---|
| `AvailabilityOnly` | Busy / Free blocks only — no titles, no locations |
| `LimitedDetails` | Busy / Free + meeting subject + location ← **default** |
| `Reviewer` | Full read access to all calendar items |
| `Author` | Read, plus create their own items in the calendar |
| `Editor` | Read, create, and modify all items |
| `PublishingEditor` | Full control including subfolders |

> `LimitedDetails` is the recommended starting point — it gives colleagues enough to schedule around someone without exposing private meeting details.

---

### 4.3 What `Default` means as an accessor identity

When the script sets `-User Default`, the word `Default` is not an actual user account. It is a built-in Exchange token that means **every authenticated person in the organisation**. It is the baseline permission row that already exists on every calendar — the script either updates it or adds it if it has been removed.

---

### 4.4 `[V]` Why two cmdlets exist: `Add` vs `Set`

A common source of confusion and errors. Here is the rule:

```
┌─────────────────────────────────────────────────────────┐
│  Permission row for "Default" already exists?           │
│                                                         │
│       YES                         NO                    │
│        │                           │                    │
│        ▼                           ▼                    │
│  Set-MailboxFolderPermission  Add-MailboxFolderPermission│
│  (modify the existing row)    (create a new row)        │
│                                                         │
│  Calling Add when row exists → ERROR                    │
│  Calling Set when row is missing → ERROR                │
└─────────────────────────────────────────────────────────┘
```

The script handles this automatically by checking with `Get-MailboxFolderPermission` first:

```powershell
$existing = Get-MailboxFolderPermission -Identity $CalendarPath `
                -User $AccessorIdentity -ErrorAction SilentlyContinue
if ($existing) { Set-MailboxFolderPermission ... }
else           { Add-MailboxFolderPermission ... }
```

---

### 4.5 Interactive login vs Certificate-Based Authentication (CBA)

Interactive login (`Connect-ExchangeOnline` with no extra parameters) opens a browser window for MFA. That works fine for manual runs.

For **monthly scheduled tasks**, there is no logged-in user and no browser session. The script will hang waiting for a prompt that never comes. The solution is **Certificate-Based Authentication (CBA)** — the certificate replaces the human in the login process.

```powershell
# Scheduled / unattended connection using CBA
Connect-ExchangeOnline `
    -CertificateThumbPrint "YOURCERTTHUMBPRINT" `
    -AppID                 "your-app-id-guid" `
    -Organization          "contoso.onmicrosoft.com"
```

The full setup for CBA is in [Section 8](#8-scheduling-for-monthly-production-use).

---

## 5. Script Walkthrough — Section by Section

### `[K]` Open the script in VS Code alongside this guide

```powershell
code scripts/Set-CalendarSharing.ps1
```

Work through each `#region` block in the script as you read each section below. The region names match the headings here.

---

### 5.1 `#region CONFIGURATION` — The only block you need to edit

```powershell
$ExcludedUsers = @(
    "ceo@contoso.com",
    "cfo@contoso.com"
)
$SharingPermission = "LimitedDetails"
$AccessorIdentity  = "Default"
```

All organisation-specific settings live here. The rest of the script reads from these variables. This separation means you can update the exclusion list or change the permission level without touching any logic.

**`[K]` Exercise:** Add three fictional exec UPNs to `$ExcludedUsers`. Save the file and run `-WhatIf`. Open the log and confirm each one appears as `EXCLUDED`.

---

### 5.2 `#region LOGGING SETUP` — Every action leaves a trail

```powershell
$LogFile = Join-Path $LogPath "CalendarSharing_$Timestamp.log"

function Write-Log {
    param ([string]$Message, [string]$Level = "INFO")
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $entry
}
```

`Add-Content` writes one line at a time without overwriting. It is safer than `Out-File` if the log is ever opened in another process simultaneously.

### `[V]` What each log level means

| Level | Console colour | When it appears |
|---|---|---|
| `INFO` | White | Normal progress — connections, counts, transitions |
| `SUCCESS` | Green | An action completed and was confirmed |
| `WARN` | Yellow | Something unexpected but non-fatal — needs a human to check |
| `ERROR` | Red | An action failed for a specific mailbox |
| `WHATIF` | Cyan | Simulation output — no real changes made |

---

### 5.3 `#region GUARDRAILS` — The script's safety net

```powershell
$ErrorActionPreference = "Stop"   # Converts non-terminating errors into terminating ones

trap {
    Write-Log "FATAL: $($_.Exception.Message)" "ERROR"
    Send-Notification -To $NotifyEmail -Subject "[ALERT] Script FAILED" -Body ...
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}
```

`trap` is the last line of defence. If anything unexpected crashes the script — a network drop mid-run, an API timeout, a null reference — `trap` catches it, writes the error to the log, emails the admin, disconnects from Exchange cleanly, and exits with code `1`.

**Why `exit 1`?** Windows Task Scheduler and Azure Automation both read the process exit code after a run. `0` means success. Anything else means failure. A non-zero exit code is what triggers your monitoring alerts.

---

### 5.4 `#region GET MAILBOXES` — Pulling and partitioning the list

```powershell
$AllMailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

$ExcludedSet = $ExcludedUsers | ForEach-Object { $_.ToLower().Trim() }

$TargetMailboxes = $AllMailboxes | Where-Object {
    ($_.PrimarySmtpAddress.ToLower() -notin $ExcludedSet) -and
    ($_.UserPrincipalName.ToLower()  -notin $ExcludedSet)
}
```

**Why `-ResultSize Unlimited`?** Without it, `Get-Mailbox` silently caps results at 1,000. In any organisation larger than that, users would be missed with no warning. Always use `Unlimited` in production.

**Why check both `PrimarySmtpAddress` and `UserPrincipalName`?** In some tenants these are different. For example, a UPN of `jsmith@corp.onmicrosoft.com` paired with an SMTP address of `j.smith@corp.com`. Checking both ensures no exec slips through an exclusion list match failure.

**`[K]` Exercise:** Export your mailbox list to CSV and inspect it:

```powershell
Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited |
    Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName |
    Export-Csv "$env:TEMP\mailboxes.csv" -NoTypeInformation
```

Open the file. Find any rows where `PrimarySmtpAddress` and `UserPrincipalName` differ — those are exactly the accounts where dual-checking matters.

---

### 5.5 `#region APPLY PERMISSIONS` — The core loop

```powershell
foreach ($Mailbox in $TargetMailboxes) {
    $CalendarPath = "$($Mailbox.PrimarySmtpAddress):\Calendar"
    $existing = Get-MailboxFolderPermission -Identity $CalendarPath `
                    -User $AccessorIdentity -ErrorAction SilentlyContinue

    if ($WhatIfMode) {
        Write-Log "[WhatIf] Would process $CalendarPath" "WHATIF"
    }
    elseif ($existing) {
        Set-MailboxFolderPermission -Identity $CalendarPath `
            -User $AccessorIdentity -AccessRights $SharingPermission
    }
    else {
        Add-MailboxFolderPermission -Identity $CalendarPath `
            -User $AccessorIdentity -AccessRights $SharingPermission
    }
}
```

### `[V]` The calendar path format

```
user@contoso.com:\Calendar
│                 │
└─ Mailbox ID     └─ Folder path  (backslash-separated, like a file system)

Subfolders:  user@contoso.com:\Calendar\Work
```

**Why does a per-mailbox failure not stop the whole script?** Each mailbox is inside a `try/catch`. If one mailbox errors, the error is logged and the loop moves on to the next. You get results for 999 mailboxes even if one fails.

---

### 5.6 `#region VERIFICATION` — Confirming what was set

```powershell
$perm = Get-MailboxFolderPermission -Identity $CalendarPath -User $AccessorIdentity
if ($perm.AccessRights -contains $SharingPermission) {
    Write-Log "VERIFIED: $DisplayName" "SUCCESS"
}
else {
    Write-Log "MISMATCH: $DisplayName — got '$($perm.AccessRights)'" "WARN"
}
```

Exchange Online is eventually consistent. In rare cases a write call returns success, but the value has not fully propagated when immediately re-read. This verification pass re-reads each permission after setting it and flags any mismatch as a `WARN`. It does not abort the script — it gives you a review item.

---

## 6. Running the Script — Step-by-Step

### `[K]` Always follow this sequence — never skip to step 4

```
Step 1  Edit $ExcludedUsers and confirm your permission level
Step 2  Run -WhatIf simulation
Step 3  Review the WhatIf log in the Logs\ folder
Step 4  Run live against one test mailbox
Step 5  Run live against all mailboxes
Step 6  Review the CSV report
```

#### Single-mailbox test run

Add this line temporarily after the partition section in the script:

```powershell
$TargetMailboxes = $TargetMailboxes | Where-Object {
    $_.PrimarySmtpAddress -eq "testuser@contoso.com"
}
```

Run the script live. Verify the log and CSV for that one mailbox. Remove the line before the full run.

#### Full production run with email notification

```powershell
.\scripts\Set-CalendarSharing.ps1 -NotifyEmail "admin@contoso.com"
```

#### Full run with a different permission level

```powershell
.\scripts\Set-CalendarSharing.ps1 -SharingPermission "AvailabilityOnly" -NotifyEmail "admin@contoso.com"
```

---

## 7. Simulation with -WhatIf

### What `-WhatIf` actually does

`-WhatIf` is a built-in PowerShell contract. When the script is declared with `[CmdletBinding(SupportsShouldProcess)]`, passing `-WhatIf` tells every operation inside: report what you would do, then stop. Nothing is written to Exchange Online.

### `[K]` Run the simulation now

```powershell
.\scripts\Set-CalendarSharing.ps1 -WhatIf
```

### `[V]` What a clean simulation output looks like

```
[2026-01-15 09:00:01] [WHATIF]  *** SIMULATION MODE (WhatIf) — No changes will be applied ***
[2026-01-15 09:00:03] [INFO]    Total mailboxes retrieved: 347
[2026-01-15 09:00:03] [INFO]    Mailboxes to be processed: 343
[2026-01-15 09:00:03] [INFO]    EXCLUDED: Jane Smith <ceo@contoso.com>
[2026-01-15 09:00:04] [WHATIF]  Would ADD permission 'LimitedDetails' for Alice Brown <abrown@contoso.com>
[2026-01-15 09:00:04] [WHATIF]  Would UPDATE permission 'LimitedDetails' for Bob Jones <bjones@contoso.com>
```

### `[V]` Pre-flight checklist — work through this before every live run

- [ ] Total mailbox count looks right
- [ ] Every exec name appears in an `EXCLUDED` line
- [ ] No exec name appears in a `WhatIf` action line
- [ ] The mix of `ADD` vs `UPDATE` actions makes sense for your org
- [ ] No unexpected `WARN` lines

---

## 8. Scheduling for Monthly Production Use

### `[V]` Two paths — pick the one that fits your environment

```
┌─────────────────────────────────────────────────────────────┐
│              SCHEDULING OPTION COMPARISON                   │
├───────────────────────┬─────────────────────────────────────┤
│  Windows Task         │  Azure Automation Runbook           │
│  Scheduler            │                                     │
├───────────────────────┼─────────────────────────────────────┤
│  Runs on a local      │  Cloud-native, no server required   │
│  server               │                                     │
│  Server must be on    │  Serverless — runs on demand        │
│  CBA cert in cert     │  Managed Identity (recommended)     │
│  store                │                                     │
│  Logs stored locally  │  Job history in Azure portal        │
│  No extra cost        │  Uses Azure Automation credits      │
└───────────────────────┴─────────────────────────────────────┘
```

---

### Option A — Windows Task Scheduler

#### Why interactive login does not work in Task Scheduler

When Task Scheduler runs a script with **"Run whether user is logged on or not"** enabled, there is no user session — no browser, no MFA prompt, no token cache. The script will hang or fail. Certificate-based authentication (CBA) solves this by replacing the human login with a certificate stored on the server.

#### Step 1 — Register an app in Microsoft Entra ID

1. Open **Microsoft Entra admin center** → App registrations → New registration
2. Name it `CalendarSharingAutomation`
3. Note the **Application (client) ID** and **Directory (tenant) ID**
4. Under **API permissions** → Add a permission → `Exchange.ManageAsApp` (application permission)
5. Click **Grant admin consent**
6. Go to **Roles and administrators** → assign the app the **Exchange Administrator** role

Official reference: [App-only authentication in Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps)

#### Step 2 — Create and install a certificate

```powershell
# Generate a self-signed certificate (valid 2 years)
$cert = New-SelfSignedCertificate `
    -Subject            "CN=ExchangeCalendarAutomation" `
    -CertStoreLocation  "Cert:\LocalMachine\My" `
    -KeyExportPolicy    Exportable `
    -KeySpec            Signature `
    -KeyLength          2048 `
    -HashAlgorithm      SHA256 `
    -NotAfter           (Get-Date).AddYears(2)

# Record the thumbprint — you will need it in the script
$cert.Thumbprint

# Export the public key (.cer) to upload to Entra
Export-Certificate -Cert $cert -FilePath "C:\Certs\CalendarAutomation.cer"
```

Upload the `.cer` file in your Entra app registration under **Certificates & secrets → Certificates**.

#### Step 3 — Update the script connection block for CBA

Replace the interactive `Connect-ExchangeOnline` call with:

```powershell
Connect-ExchangeOnline `
    -CertificateThumbPrint "YOUR_CERT_THUMBPRINT_HERE" `
    -AppID                 "YOUR_APP_ID_GUID" `
    -Organization          "contoso.onmicrosoft.com" `
    -ShowBanner:$false
```

#### Step 4 — Register the scheduled task

```powershell
$action = New-ScheduledTaskAction `
    -Execute  "pwsh.exe" `
    -Argument "-NonInteractive -NoProfile -ExecutionPolicy RemoteSigned -File `"C:\Scripts\Set-CalendarSharing.ps1`" -NotifyEmail admin@contoso.com"

# Runs on the 1st of every month at 02:00 AM
$trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At "02:00"

$principal = New-ScheduledTaskPrincipal `
    -UserId    "DOMAIN\SvcAccount" `
    -LogonType ServiceAccount `
    -RunLevel  Highest

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit  (New-TimeSpan -Hours 2) `
    -RestartCount        1 `
    -RestartInterval     (New-TimeSpan -Minutes 30)

Register-ScheduledTask `
    -TaskName    "Monthly-CalendarSharing" `
    -TaskPath    "\IT-Automation\" `
    -Action      $action `
    -Trigger     $trigger `
    -Principal   $principal `
    -Settings    $settings `
    -Description "Sets org-wide calendar sharing, excludes executives. Runs 1st of each month."
```

Official cmdlet references:
- [`New-ScheduledTask`](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/new-scheduledtask?view=windowsserver2025-ps)
- [`New-ScheduledTaskTrigger`](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/new-scheduledtasktrigger?view=windowsserver2025-ps)
- [`Register-ScheduledTask`](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/register-scheduledtask?view=windowsserver2025-ps)

#### `[K]` Confirm the task was registered correctly

```powershell
Get-ScheduledTask -TaskPath "\IT-Automation\" |
    Select-Object TaskName, State, LastRunTime, LastTaskResult
```

`LastTaskResult = 0` means the last run succeeded. Any other value means it failed — check the log file.

---

### Option B — Azure Automation (cloud-native)

```powershell
# In your Runbook, authenticate with Managed Identity — no certificate needed
Connect-ExchangeOnline -ManagedIdentity -Organization "contoso.onmicrosoft.com"
```

Official reference: [Connect using managed identity](https://learn.microsoft.com/en-us/powershell/exchange/connect-exo-powershell-managed-identity?view=exchange-ps)

---

## 9. Guardrails & Failure Handling

### `[V]` What happens when things go wrong

```
Script starts
      │
      ▼
Module installed? ──NO──► Auto-install ──FAIL──► Email admin + exit 1
      │ YES
      ▼
Connect-ExchangeOnline ──FAIL──► Email admin + exit 1
      │ OK
      ▼
Get-Mailbox ──FAIL──► Email admin + exit 1
      │ OK
      ▼
Per-mailbox loop
  └── Each mailbox has its own try/catch
         SUCCESS ──► Log SUCCESS
         FAIL    ──► Log ERROR, move to next mailbox (loop continues)
      │
      ▼
Unexpected crash ──► trap{} catches it
                         │
                         ▼
                   Log FATAL + Email admin + Disconnect + exit 1
      │
      ▼
Verification pass ──► MISMATCH logged as WARN (does not abort)
      │
      ▼
Email summary sent ──► Subject reflects SUCCESS / WARNING / SIMULATION
```

### The four guardrail principles built into this script

**1. Fail fast on infrastructure problems.**
If Exchange Online is unreachable, the script stops immediately. There is no value in attempting 1,000 mailbox updates against a dead connection.

**2. Fail slow on individual mailbox failures.**
If one mailbox errors, the script logs it and continues to the next. A bad mailbox should not prevent the other 999 from being processed.

**3. Always disconnect, even on crash.**
Both the `trap` block and the end of the script call `Disconnect-ExchangeOnline`. Open sessions count against your tenant's concurrent session limit.

**4. Exit codes are signals.**
`exit 1` tells the scheduler the job failed. Set up your monitoring to alert on non-zero exit codes from this task.

### `[K]` Test that the guardrail actually works

Temporarily add this line immediately after the logging setup region:

```powershell
throw "Deliberate test error"
```

Run the script. Then verify:
- [ ] The log contains a `FATAL` entry with your error message
- [ ] An email notification was sent (if `$NotifyEmail` is set)
- [ ] The script exited without processing any mailboxes

Remove the `throw` line once you have confirmed it.

---

## 10. Verification & Log Interpretation

### `[V]` A clean successful run looks like this

```
[2026-01-01 02:00:01] [INFO]    ===== Calendar Sharing Script Started =====
[2026-01-01 02:00:03] [SUCCESS] Connected to Exchange Online.
[2026-01-01 02:00:05] [INFO]    Total mailboxes retrieved: 412
[2026-01-01 02:00:05] [INFO]    Mailboxes to be processed: 407
[2026-01-01 02:00:05] [INFO]    Mailboxes excluded (execs): 5
[2026-01-01 02:00:05] [INFO]      EXCLUDED: Jane Smith <ceo@contoso.com>
[2026-01-01 02:01:44] [SUCCESS]   UPDATED: Alice Brown <abrown@contoso.com>
[2026-01-01 02:01:44] [SUCCESS]   ADDED:   Bob Jones <bjones@contoso.com>
[2026-01-01 02:03:10] [SUCCESS] Verification complete — all permissions confirmed.
[2026-01-01 02:03:11] [SUCCESS] EXCLUSION OK: Jane Smith — Not modified by this script.
[2026-01-01 02:03:12] [INFO]    CSV report saved: Logs\CalendarSharing_Report_20260101_020001.csv
[2026-01-01 02:03:12] [INFO]    ===== SUMMARY: Processed 407, Succeeded 407, Failed 0 =====
[2026-01-01 02:03:13] [INFO]    Disconnected from Exchange Online.
[2026-01-01 02:03:13] [INFO]    ===== Script Completed =====
```

### `[V]` Reading the CSV report in Excel

Open `Logs\CalendarSharing_Report_*.csv`. These are the columns that matter:

| Column | What a healthy run shows | What to investigate |
|---|---|---|
| `Status` | All rows say `Success` | Any row showing `Error` |
| `Action` | Mix of `Added` (new staff) and `Updated` (existing) | Unexpected `Failed` actions |
| `Error` | Empty for all rows | Any non-empty value — paste into Microsoft Learn search |

### `[K]` Monthly post-run checklist

After each scheduled run, work through this before closing your laptop:

- [ ] Log file has no `ERROR` or `FATAL` entries
- [ ] CSV `Status` column shows all `Success`
- [ ] Exec exclusion count matches your expected number
- [ ] `SuccessCount` + `FailCount` equals `TargetMailboxes.Count`
- [ ] Email notification arrived within the expected window

---

## 11. Knowledge Check

These questions test whether you understand the script well enough to support it in production. Work through them without looking at the code first.

### Questions

1. What switch do you pass to run the script without making any changes?
2. Why does `Get-Mailbox` need `-ResultSize Unlimited` in production?
3. What is the practical difference between `Add-MailboxFolderPermission` and `Set-MailboxFolderPermission`? When does the script use each one?
4. Why does interactive `Connect-ExchangeOnline` fail when run from Task Scheduler with "run when not logged on" enabled?
5. What does `exit 1` signal, and who receives that signal?
6. If `Get-Mailbox` fails entirely, should the script continue or abort? Why?
7. What does `-User Default` mean in the context of a calendar permission?
8. Name two things the `trap {}` block does when an unexpected error occurs.

### `[K]` Practical exercises

**Exercise 1 — Exclusion change**
Add `vp-sales@contoso.com` to `$ExcludedUsers`. Run `-WhatIf`. Confirm the name appears in the simulation log under `EXCLUDED`. Remove it and run `-WhatIf` again — confirm it is gone.

**Exercise 2 — Permission level comparison**
Run `-WhatIf` twice: once with `-SharingPermission LimitedDetails` and once with `-SharingPermission AvailabilityOnly`. Open both logs side by side. What is different?

**Exercise 3 — Break and recover**
Put a deliberate typo in the `$Organization` value in the CBA connection block. Run the script. Read the error in the log. Look up the correct error message in the [Connect-ExchangeOnline docs](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/connect-exchangeonline?view=exchange-ps). Fix the typo.

**Exercise 4 — Query the CSV report with PowerShell**
After a run (WhatIf or live), analyse the output:

```powershell
$report = Import-Csv (Get-ChildItem "Logs\CalendarSharing_Report_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$report | Group-Object Status  | Select-Object Name, Count
$report | Group-Object Action  | Select-Object Name, Count
$report | Where-Object { $_.Status -eq "Error" } | Select-Object DisplayName, Error
```

---

## 12. References & Official Resources

### Microsoft Learn — Core documentation

| Topic | Link |
|---|---|
| Exchange Online PowerShell overview | [exchange-online-powershell](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell?view=exchange-ps) |
| Connect-ExchangeOnline cmdlet | [connect-exchangeonline](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/connect-exchangeonline?view=exchange-ps) |
| Add-MailboxFolderPermission | [add-mailboxfolderpermission](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/add-mailboxfolderpermission?view=exchange-ps) |
| Set-MailboxFolderPermission | [set-mailboxfolderpermission](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailboxfolderpermission?view=exchange-ps) |
| Get-MailboxFolderPermission | [get-mailboxfolderpermission](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-mailboxfolderpermission?view=exchange-ps) |
| App-only (CBA) auth for unattended scripts | [app-only-auth-powershell-v2](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps) |
| Apply a sharing policy to mailboxes | [apply-a-sharing-policy](https://learn.microsoft.com/en-us/exchange/sharing/sharing-policies/apply-a-sharing-policy) |
| Create a sharing policy | [create-a-sharing-policy](https://learn.microsoft.com/en-us/exchange/sharing/sharing-policies/create-a-sharing-policy) |
| New-ScheduledTask | [new-scheduledtask](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/new-scheduledtask?view=windowsserver2025-ps) |
| New-ScheduledTaskTrigger | [new-scheduledtasktrigger](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/new-scheduledtasktrigger?view=windowsserver2025-ps) |
| Register-ScheduledTask | [register-scheduledtask](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/register-scheduledtask?view=windowsserver2025-ps) |
| Manage M365 services with PowerShell (Learning Path) | [manage-microsoft-365-services](https://learn.microsoft.com/en-us/training/paths/manage-microsoft-365-services-use-windows-powershell/) |
| AZ-040T00: Automate Administration with PowerShell | [az-040t00](https://learn.microsoft.com/en-us/training/courses/az-040t00) |
| Get started with PowerShell for Microsoft 365 | [getting-started-m365-powershell](https://learn.microsoft.com/en-us/microsoft-365/enterprise/getting-started-with-microsoft-365-powershell?view=o365-worldwide) |

### Microsoft Learn — Video series

| Video | Link |
|---|---|
| Getting Started with PowerShell 3.0 (Series) | [shows/getstartedpowershell3](https://learn.microsoft.com/en-us/shows/getstartedpowershell3/01) |
| PowerShell for Beginners (MVP Series) | [shows/powershell-beginners](https://learn.microsoft.com/en-us/shows/mvp-windows-and-devices-for-it/powershell-beginners) |

### PowerShell Gallery

| Package | Link |
|---|---|
| ExchangeOnlineManagement (latest) | [powershellgallery.com/ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) |

---

*Document version 2.0 — Visual and Kinesthetic edition*
*Covers `Set-CalendarSharing.ps1` v1.0*
*Next review: after any ExchangeOnlineManagement module major version update*
