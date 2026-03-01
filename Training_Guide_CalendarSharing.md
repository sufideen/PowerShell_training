# Exchange Online Calendar Sharing Automation
## PowerShell Training Guide — Intermediate to Advanced

**Script:** `Set-CalendarSharing.ps1`
**Audience:** IT Administrators with working PowerShell knowledge
**Learning Style Coverage:** Visual · Auditory · Kinesthetic (VAK)
**Production Use:** Monthly scheduled run via Windows Task Scheduler or Azure Automation

---

## How to Use This Guide (VAK Framework)

This guide is structured for three learning styles. Work through all three for maximum retention.

| Symbol | Style | How to engage |
|---|---|---|
| `[V]` | **Visual** | Read diagrams, tables, and flow charts. Study the code with syntax highlighting in VS Code. |
| `[A]` | **Auditory** | Read each section aloud. Watch the linked Microsoft videos. Narrate what each code block does. |
| `[K]` | **Kinesthetic** | Type every example yourself — do not copy-paste. Run `-WhatIf` first. Break things in a test tenant. |

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

Your organisation needs all employee calendars to be visible to colleagues at a defined permission level (e.g. `LimitedDetails`) — but the CEO and named executives must be explicitly excluded for privacy and governance reasons.

Doing this manually in the Microsoft 365 admin portal:
- Does not scale (hundreds or thousands of mailboxes)
- Has no audit trail
- Cannot easily be repeated monthly as new staff join

PowerShell with the **ExchangeOnlineManagement** module solves all three problems.

### `[A]` Auditory anchor

Say this aloud:
> *"The script connects to Exchange Online, gets every user mailbox, skips the executives, sets the calendar permission, checks it applied, logs everything, and tells me what happened."*

That one sentence is the entire script. Every section you read below is one of those steps in detail.

---

## 2. Architecture Overview — Visual Map

### `[V]` Flow Diagram

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

### `[V]` Data flow for a single mailbox

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

### `[K]` Do these steps yourself — type every command

#### Step 1 — Verify PowerShell version

```powershell
$PSVersionTable.PSVersion
# You need: Major version 5 (Windows) or 7+ (cross-platform)
```

#### Step 2 — Install or update the Exchange Online module

```powershell
# Check if already installed
Get-Module -ListAvailable -Name ExchangeOnlineManagement

# Install (if missing)
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force

# Update to latest (run periodically)
Update-Module -Name ExchangeOnlineManagement
```

> **Current latest version:** 3.9.0 (as of early 2026)
> Source: [PowerShell Gallery — ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement)

#### Step 3 — Set execution policy for your session

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

#### Step 4 — Confirm your admin role

You need one of these Microsoft Entra roles:
- **Exchange Administrator** (preferred — least privilege)
- **Global Administrator** (emergency only — avoid for routine scripts)

Check your role:
```powershell
Connect-ExchangeOnline
Get-ManagementRoleAssignment -RoleAssignee (Get-Mailbox -ResultSize 1 | Select -Exp UserPrincipalName)
```

### `[V]` Prerequisites summary table

| Requirement | Minimum | Check command |
|---|---|---|
| PowerShell | 5.1 or 7+ | `$PSVersionTable.PSVersion` |
| ExchangeOnlineManagement | 3.x | `Get-Module ExchangeOnlineManagement -ListAvailable` |
| Admin Role | Exchange Admin | M365 Admin Portal → Roles |
| Network | HTTPS 443 outbound | `Test-NetConnection outlook.office365.com -Port 443` |
| Execution Policy | RemoteSigned | `Get-ExecutionPolicy -Scope CurrentUser` |

---

## 4. Key Concepts Explained

### `[A]` Read each concept aloud and explain it to an imaginary colleague

#### 4.1 Mailbox Folder Permissions vs Sharing Policies

These are two separate Exchange mechanisms. The script uses **Mailbox Folder Permissions**.

| Mechanism | Scope | Used for |
|---|---|---|
| `MailboxFolderPermission` | Per-folder, per-user | Internal org calendar access |
| `SharingPolicy` | Organisation-wide rule | External domain sharing |
| `OrganizationRelationship` | Federated tenants | Cross-company free/busy |

This script targets `:\Calendar` using `Add/Set-MailboxFolderPermission`.

#### 4.2 Permission Levels — what each one exposes

| AccessRights value | What the accessor can see |
|---|---|
| `AvailabilityOnly` | Free/Busy time only (no details) |
| `LimitedDetails` | Free/Busy + subject + location |
| `Reviewer` | Full read of all calendar items |
| `Author` | Read + create own items |
| `Editor` | Read + create + modify all items |
| `PublishingEditor` | Full control including subfolders |

> Default in this script: **`LimitedDetails`** — the least-privilege option that still provides useful scheduling context.

#### 4.3 The `Default` accessor identity

In `Add-MailboxFolderPermission -User Default`, the word `Default` is a special Exchange token meaning *"all authenticated users in the organisation"*. It is not a user account. It is the baseline permission row that appears on every calendar.

#### 4.4 `Add` vs `Set` — why both cmdlets exist

```powershell
# Add-MailboxFolderPermission  ── creates a NEW permission row
# Set-MailboxFolderPermission  ── modifies an EXISTING permission row
# If you call Add on a row that already exists: ERROR
# If you call Set on a row that does not exist: ERROR
# Solution: check first with Get-, then branch
```

The script handles this automatically:
```powershell
$existing = Get-MailboxFolderPermission -Identity $CalendarPath -User $AccessorIdentity -ErrorAction SilentlyContinue
if ($existing) { Set-MailboxFolderPermission ... }
else           { Add-MailboxFolderPermission ... }
```

#### 4.5 Modern Auth and App-Only (CBA) for scheduled runs

Interactive login (`Connect-ExchangeOnline` with no parameters) works for manual runs. For **monthly scheduled tasks**, interactive login fails because there is no logged-in user.

The production-safe approach is **Certificate-Based Authentication (CBA)**:

```powershell
Connect-ExchangeOnline `
    -CertificateThumbPrint "YOURCERTTHUMBPRINT" `
    -AppID                 "your-app-id-guid" `
    -Organization          "contoso.onmicrosoft.com"
```

See [Section 8](#8-scheduling-for-monthly-production-use) for the complete setup.

---

## 5. Script Walkthrough — Section by Section

### `[K]` Open `Set-CalendarSharing.ps1` in VS Code alongside this guide

```
code Set-CalendarSharing.ps1
```

Work through each region tag in the script as you read below.

---

### 5.1 `#region CONFIGURATION` — The control panel

```powershell
$ExcludedUsers = @(
    "ceo@contoso.com",
    "cfo@contoso.com"
)
$SharingPermission    = "LimitedDetails"
$AccessorIdentity     = "Default"
```

**Why it matters:** This is the only block you need to edit for a new deployment. All business logic lives here — separation of configuration from code is a production best practice.

**`[K]` Exercise:** Add three more exec UPNs to `$ExcludedUsers`. Save and run `-WhatIf`. Confirm the new exclusions appear in the log.

---

### 5.2 `#region LOGGING SETUP` — Timestamped, structured logs

```powershell
$LogFile = Join-Path $LogPath "CalendarSharing_$Timestamp.log"

function Write-Log {
    param ([string]$Message, [string]$Level = "INFO")
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $entry
    # Also writes colour-coded output to the console
}
```

**Key technique:** `Add-Content` appends a single line atomically. It is safer than `Out-File` when multiple processes could write to the same file.

**`[V]` Log level colour map:**

| Level | Console colour | Meaning |
|---|---|---|
| `INFO` | White | Normal progress |
| `SUCCESS` | Green | Action confirmed |
| `WARN` | Yellow | Non-fatal issue, needs review |
| `ERROR` | Red | Action failed |
| `WHATIF` | Cyan | Simulation output |

---

### 5.3 `#region GUARDRAILS` — The safety net

```powershell
$ErrorActionPreference = "Stop"   # All errors become terminating

trap {
    Write-Log "FATAL: $($_.Exception.Message)" "ERROR"
    Send-Notification -To $NotifyEmail -Subject "[ALERT] Script FAILED" -Body ...
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}
```

**`[A]` Say aloud:** *"`trap` is the script's emergency brake. If anything unexpected happens — a network drop, an API timeout, a null reference — `trap` runs, writes the error, emails the admin, disconnects cleanly, and exits with code 1 so the scheduler knows it failed."*

**Why `exit 1`?** Windows Task Scheduler and Azure Automation both check the process exit code. Exit code `0` = success. Any non-zero = failure. This triggers alerts in your monitoring system.

---

### 5.4 `#region GET MAILBOXES` — Partition logic

```powershell
$AllMailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

# Normalize to lowercase for reliable string comparison
$ExcludedSet = $ExcludedUsers | ForEach-Object { $_.ToLower().Trim() }

$TargetMailboxes = $AllMailboxes | Where-Object {
    ($_.PrimarySmtpAddress.ToLower() -notin $ExcludedSet) -and
    ($_.UserPrincipalName.ToLower()  -notin $ExcludedSet)
}
```

**Why `-ResultSize Unlimited`?** The default `Get-Mailbox` returns only 1,000 results. In any organisation with more than 1,000 mailboxes this would silently miss users. Always use `Unlimited` in production.

**Why check both `PrimarySmtpAddress` AND `UserPrincipalName`?** In some tenants these differ (e.g. a UPN of `jsmith@corp.onmicrosoft.com` vs SMTP of `j.smith@corp.com`). Checking both prevents accidental exclusion failures.

**`[K]` Exercise:** Run this in a test tenant:
```powershell
Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited |
    Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName |
    Export-Csv "$env:TEMP\mailboxes.csv" -NoTypeInformation
```
Open the CSV. Find any rows where PrimarySmtpAddress ≠ UserPrincipalName.

---

### 5.5 `#region APPLY PERMISSIONS` — The core loop

```powershell
foreach ($Mailbox in $TargetMailboxes) {
    $CalendarPath = "$($Mailbox.PrimarySmtpAddress):\Calendar"
    $existing = Get-MailboxFolderPermission -Identity $CalendarPath -User $AccessorIdentity -ErrorAction SilentlyContinue

    if ($WhatIfMode) {
        Write-Log "[WhatIf] Would process $CalendarPath" "WHATIF"
    }
    elseif ($existing) {
        Set-MailboxFolderPermission -Identity $CalendarPath -User $AccessorIdentity -AccessRights $SharingPermission
    }
    else {
        Add-MailboxFolderPermission -Identity $CalendarPath -User $AccessorIdentity -AccessRights $SharingPermission
    }
}
```

**`[V]` The `:\Calendar` path format**

```
user@contoso.com:\Calendar
│                 │
└─ Mailbox ID     └─ Folder path (backslash-separated like a filesystem)
```

Subfolders use: `user@contoso.com:\Calendar\Work`

**`[A]` Explain this loop aloud** before running it live. Cover: what happens if `Get-MailboxFolderPermission` returns `$null`, what `-ErrorAction SilentlyContinue` does, and why the loop continues even if one mailbox fails (the `try/catch` inside the loop catches per-mailbox errors without halting the whole script).

---

### 5.6 `#region VERIFICATION` — Trust but verify

```powershell
$perm = Get-MailboxFolderPermission -Identity $CalendarPath -User $AccessorIdentity
if ($perm.AccessRights -contains $SharingPermission) {
    Write-Log "VERIFIED: $DisplayName" "SUCCESS"
}
else {
    Write-Log "MISMATCH: $DisplayName — got '$($perm.AccessRights)'" "WARN"
}
```

**Why verify?** Exchange Online is eventually consistent. In rare cases a `Set-` call returns success but the change propagates with a short delay. Re-reading confirms the state the system recorded.

---

## 6. Running the Script — Step-by-Step

### `[K]` Hands-on execution sequence

Always follow this order — never skip to step 3:

```
Step 1: Edit $ExcludedUsers in the configuration block
Step 2: Run -WhatIf simulation (see Section 7)
Step 3: Review the WhatIf log in Logs\
Step 4: Run live against a single test mailbox (see below)
Step 5: Run live against all mailboxes
Step 6: Review the CSV report
```

#### Single-mailbox test run

```powershell
# Override $TargetMailboxes to just one mailbox for your first live test
# In the script, after the partition section, temporarily add:
$TargetMailboxes = $TargetMailboxes | Where-Object { $_.PrimarySmtpAddress -eq "testuser@contoso.com" }
```

Remove that line after testing.

#### Full production run with notification

```powershell
.\Set-CalendarSharing.ps1 -NotifyEmail "admin@contoso.com"
```

#### Full run with custom permission level

```powershell
.\Set-CalendarSharing.ps1 -SharingPermission "AvailabilityOnly" -NotifyEmail "admin@contoso.com"
```

---

## 7. Simulation with -WhatIf

### `[A]` What `-WhatIf` actually does — say this aloud

> *"WhatIf is a contract. When a cmdlet supports `SupportsShouldProcess`, passing `-WhatIf` tells every supporting operation inside: report what you would do, then stop without doing it. Our script uses `[CmdletBinding(SupportsShouldProcess)]` so the entire script honours this contract."*

### `[K]` Run simulation now

```powershell
.\Set-CalendarSharing.ps1 -WhatIf
```

**Expect to see in the console and log:**
```
[2026-01-15 09:00:01] [WHATIF] *** SIMULATION MODE (WhatIf) — No changes will be applied ***
[2026-01-15 09:00:03] [INFO]   Total mailboxes retrieved: 347
[2026-01-15 09:00:03] [INFO]   Mailboxes to be processed: 343
[2026-01-15 09:00:03] [INFO]   EXCLUDED: Jane Smith <ceo@contoso.com>
[2026-01-15 09:00:04] [WHATIF] Would ADD calendar permission 'LimitedDetails' for Alice Brown <abrown@contoso.com>
[2026-01-15 09:00:04] [WHATIF] Would UPDATE calendar permission 'LimitedDetails' for Bob Jones <bjones@contoso.com>
```

**`[V]` WhatIf checklist — review the log before going live:**

- [ ] Total mailbox count looks correct
- [ ] All exec names appear in EXCLUDED lines
- [ ] No exec names appear in WhatIf action lines
- [ ] ADD vs UPDATE distribution is what you expected
- [ ] No unexpected `WARN` entries

---

## 8. Scheduling for Monthly Production Use

### `[V]` Two options — choose based on your environment

```
┌─────────────────────────────────────────────────────────────┐
│              SCHEDULING OPTION COMPARISON                   │
├───────────────────────┬─────────────────────────────────────┤
│  Windows Task         │  Azure Automation Runbook           │
│  Scheduler            │                                     │
├───────────────────────┼─────────────────────────────────────┤
│  On-premise server    │  Cloud-native, no server needed     │
│  Must be always on    │  Serverless, runs on demand         │
│  CBA cert in store    │  Managed Identity (best practice)   │
│  Log files local      │  Job history in Azure portal        │
│  Free                 │  Costs Azure Automation credits     │
└───────────────────────┴─────────────────────────────────────┘
```

---

### Option A — Windows Task Scheduler (on-premise server)

#### `[A]` Why interactive login breaks in Task Scheduler

> *"When Task Scheduler runs a script with 'Run whether user is logged on or not', there is no user session. No session means no browser for MFA, no token cache, no interactive prompt. The script hangs or fails. Certificate-based authentication removes the human from the login — the certificate IS the credential."*

#### Step 1 — Register an app in Microsoft Entra ID

1. Go to **Microsoft Entra admin center** → App registrations → New registration
2. Name: `CalendarSharingAutomation`
3. Note the **Application (client) ID** and **Directory (tenant) ID**
4. Under **API permissions** → Add → `Exchange.ManageAsApp` (application permission)
5. Grant admin consent
6. Assign the app the **Exchange Administrator** role in Entra → Roles and administrators

Official reference: [App-only authentication in Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps)

#### Step 2 — Create and install a certificate

```powershell
# Generate a self-signed certificate (valid 2 years)
$cert = New-SelfSignedCertificate `
    -Subject       "CN=ExchangeCalendarAutomation" `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Note the thumbprint
$cert.Thumbprint

# Export the public key (.cer) to upload to Entra
Export-Certificate -Cert $cert -FilePath "C:\Certs\CalendarAutomation.cer"
```

Upload the `.cer` file to your Entra app registration under **Certificates & secrets → Certificates**.

#### Step 3 — Update the script connection block for CBA

Replace the interactive `Connect-ExchangeOnline` call with:

```powershell
Connect-ExchangeOnline `
    -CertificateThumbPrint "YOUR_CERT_THUMBPRINT_HERE" `
    -AppID                 "YOUR_APP_ID_GUID" `
    -Organization          "contoso.onmicrosoft.com" `
    -ShowBanner:$false
```

#### Step 4 — Create the scheduled task via PowerShell

```powershell
$action = New-ScheduledTaskAction `
    -Execute  "pwsh.exe" `
    -Argument "-NonInteractive -NoProfile -ExecutionPolicy RemoteSigned -File `"C:\Scripts\Set-CalendarSharing.ps1`" -NotifyEmail admin@contoso.com"

# Run on the 1st of every month at 02:00 AM
$trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At "02:00"

$principal = New-ScheduledTaskPrincipal `
    -UserId    "DOMAIN\SvcAccount" `
    -LogonType ServiceAccount `
    -RunLevel  Highest

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
    -RestartCount 1 `
    -RestartInterval (New-TimeSpan -Minutes 30)

Register-ScheduledTask `
    -TaskName  "Monthly-CalendarSharing" `
    -TaskPath  "\IT-Automation\" `
    -Action    $action `
    -Trigger   $trigger `
    -Principal $principal `
    -Settings  $settings `
    -Description "Sets org-wide calendar sharing permissions, excludes executives. Runs 1st of month."
```

Official cmdlet reference:
- [`New-ScheduledTask`](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/new-scheduledtask?view=windowsserver2025-ps)
- [`New-ScheduledTaskTrigger`](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/new-scheduledtasktrigger?view=windowsserver2025-ps)
- [`Register-ScheduledTask`](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/register-scheduledtask?view=windowsserver2025-ps)

#### `[K]` Verify the task registered correctly

```powershell
Get-ScheduledTask -TaskPath "\IT-Automation\" | Select-Object TaskName, State, LastRunTime, LastTaskResult
```

`LastTaskResult = 0` means the last run succeeded.

---

### Option B — Azure Automation (cloud-native)

```powershell
# In your Runbook, use Managed Identity instead of CBA
Connect-ExchangeOnline -ManagedIdentity -Organization "contoso.onmicrosoft.com"
```

Official reference: [Connect using managed identity](https://learn.microsoft.com/en-us/powershell/exchange/connect-exo-powershell-managed-identity?view=exchange-ps)

---

## 9. Guardrails & Failure Handling

### `[V]` Failure decision tree

```
Script starts
      │
      ▼
Module installed? ──NO──► Auto-install ──FAIL──► Email + exit 1
      │ YES
      ▼
Connect-ExchangeOnline ──FAIL──► Email + exit 1
      │ OK
      ▼
Get-Mailbox ──FAIL──► Email + exit 1
      │ OK
      ▼
Per-mailbox loop
  └── Each mailbox: try/catch
         SUCCESS ──► Log SUCCESS, add to results
         FAIL    ──► Log ERROR, continue to next mailbox
      │
      ▼
Unexpected crash ──► trap{} ──► Log FATAL, email, Disconnect, exit 1
      │
      ▼
Verification pass ──► MISMATCH? Log WARN (does not abort)
      │
      ▼
Email summary ──► Subject: SUCCESS / WARNING / SIMULATION
```

### `[A]` Guardrail principles — say each one aloud

1. **Fail fast on infrastructure failures** — if Exchange Online is unreachable, stop immediately. There is no point looping 1,000 mailboxes against a dead connection.

2. **Fail slow on per-item failures** — if one mailbox errors, log it and move on. You want the other 999 to succeed.

3. **Always disconnect** — the `trap` block and the final disconnect both call `Disconnect-ExchangeOnline`. Leaving sessions open consumes your tenant's concurrent session quota.

4. **Exit codes are signals** — `exit 1` tells the scheduler the job failed. Configure your scheduler to email or alert on non-zero exit codes.

### `[K]` Test the guardrail

Temporarily break the script to confirm `trap` fires:

```powershell
# Add this line temporarily after the logging setup region:
throw "Deliberate test error"
```

Run the script. Confirm:
- [ ] `FATAL` entry appears in the log
- [ ] Email sent to `$NotifyEmail` (if configured)
- [ ] Script exits and does not process any mailboxes

Remove the throw line afterwards.

---

## 10. Verification & Log Interpretation

### `[V]` Reading a successful run log

```
[2026-01-01 02:00:01] [INFO]    ===== Calendar Sharing Script Started =====
[2026-01-01 02:00:03] [SUCCESS] Connected to Exchange Online.
[2026-01-01 02:00:05] [INFO]    Total mailboxes retrieved: 412
[2026-01-01 02:00:05] [INFO]    Mailboxes to be processed: 407
[2026-01-01 02:00:05] [INFO]    Mailboxes excluded (execs): 5
[2026-01-01 02:00:05] [INFO]      EXCLUDED: Jane Smith <ceo@contoso.com>
[2026-01-01 02:01:44] [SUCCESS]   UPDATED: Alice Brown <abrown@contoso.com>
[2026-01-01 02:01:44] [SUCCESS]   ADDED:   Bob Jones <bjones@contoso.com>
[2026-01-01 02:03:10] [SUCCESS] Verification complete — all checked permissions confirmed.
[2026-01-01 02:03:11] [SUCCESS] EXCLUSION OK: Jane Smith — Not modified by this script.
[2026-01-01 02:03:12] [INFO]    CSV report saved: Logs\CalendarSharing_Report_20260101_020001.csv
[2026-01-01 02:03:12] [INFO]    ===== SUMMARY: Processed 407, Succeeded 407, Failed 0 =====
[2026-01-01 02:03:13] [INFO]    Disconnected from Exchange Online.
[2026-01-01 02:03:13] [INFO]    ===== Script Completed =====
```

### `[V]` Reading the CSV report

Open `CalendarSharing_Report_*.csv` in Excel. Key columns:

| Column | What to look for |
|---|---|
| `Status` | All rows should be `Success`. Filter for `Error`. |
| `Action` | Mix of `Added` (new staff) and `Updated` (returning) is normal |
| `Error` | For failed rows — paste error text into Microsoft Learn search |

### `[K]` Monthly review checklist

After each scheduled run:
- [ ] Open log — no `ERROR` or `FATAL` lines
- [ ] Open CSV — `Status` column: all `Success`
- [ ] Exec count matches expectation (no new execs unaccounted for)
- [ ] `SuccessCount` + `FailCount` = `TargetMailboxes.Count`
- [ ] Email notification received within expected time window

---

## 11. Knowledge Check

### `[A]` Answer these aloud without looking at the script

1. What cmdlet do you run first to test a script without making changes?
2. Why does `-ResultSize Unlimited` matter in `Get-Mailbox`?
3. What is the difference between `Add-MailboxFolderPermission` and `Set-MailboxFolderPermission`?
4. Why does interactive `Connect-ExchangeOnline` fail in Task Scheduler?
5. What exit code signals failure to a scheduler, and why?
6. If the `Get-Mailbox` call fails, should the script continue or abort? Why?
7. What does the `Default` user identity mean in calendar permissions?
8. Name two things the `trap {}` block does.

### `[K]` Practical exercises

**Exercise 1 — Exclusion change**
Add a new exec `vp-sales@contoso.com` to `$ExcludedUsers`. Run `-WhatIf`. Confirm they appear in the exclusion log. Remove them and repeat.

**Exercise 2 — Permission level comparison**
Run the script twice with `-WhatIf`: once with `LimitedDetails`, once with `AvailabilityOnly`. Compare the WhatIf log entries — what changes?

**Exercise 3 — Break and recover**
Intentionally use a wrong `$Organization` value in the CBA connection block. Run the script. Read the error in the log. Find the correct error message in the [Microsoft Learn Connect-ExchangeOnline docs](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/connect-exchangeonline?view=exchange-ps).

**Exercise 4 — Log analysis**
Open an old CSV report. Use PowerShell to query it:
```powershell
$report = Import-Csv "Logs\CalendarSharing_Report_*.csv" | Select-Object -Last 1
$report | Group-Object Status | Select-Object Name, Count
$report | Where-Object { $_.Status -eq "Error" } | Select-Object DisplayName, Error
```

---

## 12. References & Official Resources

### Microsoft Learn — Core documentation

| Topic | URL |
|---|---|
| Exchange Online PowerShell overview | [learn.microsoft.com/exchange-online-powershell](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell?view=exchange-ps) |
| Connect-ExchangeOnline cmdlet | [learn.microsoft.com/connect-exchangeonline](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/connect-exchangeonline?view=exchange-ps) |
| Add-MailboxFolderPermission | [learn.microsoft.com/add-mailboxfolderpermission](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/add-mailboxfolderpermission?view=exchange-ps) |
| Set-MailboxFolderPermission | [learn.microsoft.com/set-mailboxfolderpermission](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailboxfolderpermission?view=exchange-ps) |
| Get-MailboxFolderPermission | [learn.microsoft.com/get-mailboxfolderpermission](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-mailboxfolderpermission?view=exchange-ps) |
| App-only (CBA) auth for unattended scripts | [learn.microsoft.com/app-only-auth-powershell-v2](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps) |
| Apply a sharing policy to mailboxes | [learn.microsoft.com/apply-a-sharing-policy](https://learn.microsoft.com/en-us/exchange/sharing/sharing-policies/apply-a-sharing-policy) |
| Create a sharing policy | [learn.microsoft.com/create-a-sharing-policy](https://learn.microsoft.com/en-us/exchange/sharing/sharing-policies/create-a-sharing-policy) |
| New-ScheduledTask | [learn.microsoft.com/new-scheduledtask](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/new-scheduledtask?view=windowsserver2025-ps) |
| New-ScheduledTaskTrigger | [learn.microsoft.com/new-scheduledtasktrigger](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/new-scheduledtasktrigger?view=windowsserver2025-ps) |
| Register-ScheduledTask | [learn.microsoft.com/register-scheduledtask](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/register-scheduledtask?view=windowsserver2025-ps) |
| Manage M365 services with PowerShell (Learning Path) | [learn.microsoft.com/manage-microsoft-365-services](https://learn.microsoft.com/en-us/training/paths/manage-microsoft-365-services-use-windows-powershell/) |
| AZ-040T00: Automate Administration with PowerShell | [learn.microsoft.com/az-040t00](https://learn.microsoft.com/en-us/training/courses/az-040t00) |
| Get started with PowerShell for Microsoft 365 | [learn.microsoft.com/getting-started-m365-powershell](https://learn.microsoft.com/en-us/microsoft-365/enterprise/getting-started-with-microsoft-365-powershell?view=o365-worldwide) |

### Microsoft Learn — Video series

| Video | URL |
|---|---|
| Getting Started with PowerShell 3.0 (Series, Microsoft Learn Shows) | [learn.microsoft.com/shows/getstartedpowershell3](https://learn.microsoft.com/en-us/shows/getstartedpowershell3/01) |
| PowerShell for Beginners (MVP Series, Microsoft Learn Shows) | [learn.microsoft.com/shows/powershell-beginners](https://learn.microsoft.com/en-us/shows/mvp-windows-and-devices-for-it/powershell-beginners) |

### PowerShell Gallery

| Package | URL |
|---|---|
| ExchangeOnlineManagement (latest) | [powershellgallery.com/ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) |

---

*Document version 1.0 — Covers `Set-CalendarSharing.ps1` v1.0*
*Next review: after any ExchangeOnlineManagement module major version update*
