<#
.SYNOPSIS
    Sets Calendar Sharing permissions for all M365 mailboxes, excluding specified executives.

.DESCRIPTION
    Connects to Exchange Online and configures calendar sharing permissions for all users.
    Excludes a defined list of executives (CEO and exec group) from having their calendars shared.
    Supports -WhatIf simulation, logging, email notifications, and post-run verification.

.PARAMETER WhatIf
    Simulates all changes without applying them. No mailboxes are modified.

.PARAMETER SharingPermission
    The calendar permission level to apply. Default: "LimitedDetails"
    Valid: Reviewer, LimitedDetails, AvailabilityOnly, Editor, Author, PublishingEditor

.PARAMETER LogPath
    Path to write the log file. Defaults to script directory.

.PARAMETER NotifyEmail
    Admin email address to receive success/failure notifications.

.EXAMPLE
    .\Set-CalendarSharing.ps1 -WhatIf
    Runs a simulation — shows what would happen without making changes.

.EXAMPLE
    .\Set-CalendarSharing.ps1 -NotifyEmail "admin@contoso.com"
    Runs the script and emails results to the admin.

.NOTES
    Requires: ExchangeOnlineManagement module
    Run As:   Authenticated admin with Exchange Admin or Global Admin role
    Author:   IT Administration
    Version:  1.0
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
    [ValidateSet("Reviewer", "LimitedDetails", "AvailabilityOnly", "Editor", "Author", "PublishingEditor")]
    [string]$SharingPermission = "LimitedDetails",

    [string]$LogPath = "",

    [string]$NotifyEmail = "",

    [switch]$SkipConnectionCheck
)

#region ── CONFIGURATION ────────────────────────────────────────────────────────

# ── Exclusion list: CEO and executive group ──────────────────────────────────
# Add UPN (UserPrincipalName) or primary SMTP addresses here.
$ExcludedUsers = @(
    "ceo@contoso.com",
    "cfo@contoso.com",
    "coo@contoso.com",
    "cto@contoso.com",
    "chro@contoso.com"
    # Add more executives as needed
)

# ── Notification settings ────────────────────────────────────────────────────
$SmtpServer     = "smtp.office365.com"
$SmtpPort       = 587
$SenderEmail    = "it-admin@contoso.com"   # Must be a valid M365 mailbox

# ── Sharing target: who can SEE the calendar (the external accessor) ─────────
# "Default" = all authenticated users in the org; change to a specific user if needed
$AccessorIdentity = "Default"

#endregion ──────────────────────────────────────────────────────────────────────

#region ── LOGGING SETUP ────────────────────────────────────────────────────────

if (-not $LogPath) {
    $LogPath = Join-Path $PSScriptRoot "Logs"
}

if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile    = Join-Path $LogPath "CalendarSharing_$Timestamp.log"
$WhatIfMode = $WhatIfPreference.IsPresent -or ($PSBoundParameters.ContainsKey('WhatIf'))

function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS","WHATIF")]
        [string]$Level = "INFO"
    )
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $entry
    switch ($Level) {
        "ERROR"   { Write-Host $entry -ForegroundColor Red }
        "WARN"    { Write-Host $entry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $entry -ForegroundColor Green }
        "WHATIF"  { Write-Host $entry -ForegroundColor Cyan }
        default   { Write-Host $entry }
    }
}

#endregion ──────────────────────────────────────────────────────────────────────

#region ── NOTIFICATION HELPER ──────────────────────────────────────────────────

function Send-Notification {
    param (
        [string]$To,
        [string]$Subject,
        [string]$Body
    )
    if (-not $To) { return }
    try {
        $credential = Get-Credential -Message "Enter credentials for notification email sender ($SenderEmail)" -ErrorAction Stop
        Send-MailMessage `
            -From       $SenderEmail `
            -To         $To `
            -Subject    $Subject `
            -Body       $Body `
            -SmtpServer $SmtpServer `
            -Port       $SmtpPort `
            -UseSsl `
            -Credential $credential `
            -ErrorAction Stop
        Write-Log "Notification email sent to $To." "INFO"
    }
    catch {
        Write-Log "Failed to send notification email: $_" "WARN"
    }
}

#endregion ──────────────────────────────────────────────────────────────────────

#region ── GUARDRAILS: TRAP & GLOBAL ERROR HANDLER ──────────────────────────────

$ErrorActionPreference = "Stop"   # Convert all non-terminating errors to terminating

trap {
    $msg = "FATAL: Script terminated unexpectedly. Error: $($_.Exception.Message)"
    Write-Log $msg "ERROR"
    Send-Notification -To $NotifyEmail `
        -Subject "[ALERT] CalendarSharing script FAILED" `
        -Body "$msg`n`nCheck log: $LogFile"

    # Attempt to disconnect cleanly even on failure
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    exit 1
}

#endregion ──────────────────────────────────────────────────────────────────────

#region ── PREREQUISITE CHECK ───────────────────────────────────────────────────

Write-Log "========== Calendar Sharing Script Started ==========" "INFO"
if ($WhatIfMode) { Write-Log "*** SIMULATION MODE (WhatIf) — No changes will be applied ***" "WHATIF" }

# Verify ExchangeOnlineManagement module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Log "ExchangeOnlineManagement module not found. Attempting install..." "WARN"
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -ErrorAction Stop
        Write-Log "Module installed successfully." "SUCCESS"
    }
    catch {
        Write-Log "Cannot install ExchangeOnlineManagement module: $_" "ERROR"
        Send-Notification -To $NotifyEmail `
            -Subject "[ALERT] CalendarSharing script FAILED — Missing Module" `
            -Body "ExchangeOnlineManagement module could not be installed.`nError: $_`nLog: $LogFile"
        exit 1
    }
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

#endregion ──────────────────────────────────────────────────────────────────────

#region ── CONNECT TO EXCHANGE ONLINE ───────────────────────────────────────────

if (-not $SkipConnectionCheck) {
    try {
        Write-Log "Connecting to Exchange Online..." "INFO"
        # Uses modern auth / MFA-compatible interactive login
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Connected to Exchange Online." "SUCCESS"
    }
    catch {
        Write-Log "Connection to Exchange Online failed: $_" "ERROR"
        Send-Notification -To $NotifyEmail `
            -Subject "[ALERT] CalendarSharing script FAILED — Connection Error" `
            -Body "Could not connect to Exchange Online.`nError: $_`nLog: $LogFile"
        exit 1
    }
}

#endregion ──────────────────────────────────────────────────────────────────────

#region ── GET MAILBOXES & BUILD EXCLUSION SET ───────────────────────────────────

Write-Log "Retrieving all user mailboxes from Exchange Online..." "INFO"

try {
    $AllMailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -ErrorAction Stop
    Write-Log "Total mailboxes retrieved: $($AllMailboxes.Count)" "INFO"
}
catch {
    Write-Log "Failed to retrieve mailboxes: $_" "ERROR"
    Send-Notification -To $NotifyEmail `
        -Subject "[ALERT] CalendarSharing script FAILED — Get-Mailbox Error" `
        -Body "Could not retrieve mailboxes.`nError: $_`nLog: $LogFile"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

# Normalize exclusion list to lowercase for reliable comparison
$ExcludedSet = $ExcludedUsers | ForEach-Object { $_.ToLower().Trim() }

# Partition mailboxes
$TargetMailboxes   = $AllMailboxes | Where-Object {
    ($_.PrimarySmtpAddress.ToLower() -notin $ExcludedSet) -and
    ($_.UserPrincipalName.ToLower()  -notin $ExcludedSet)
}
$ExcludedMailboxes = $AllMailboxes | Where-Object {
    ($_.PrimarySmtpAddress.ToLower() -in $ExcludedSet) -or
    ($_.UserPrincipalName.ToLower()  -in $ExcludedSet)
}

Write-Log "Mailboxes to be processed : $($TargetMailboxes.Count)" "INFO"
Write-Log "Mailboxes excluded (execs): $($ExcludedMailboxes.Count)" "INFO"

# Log excluded names for audit trail
foreach ($ex in $ExcludedMailboxes) {
    Write-Log "  EXCLUDED: $($ex.DisplayName) <$($ex.PrimarySmtpAddress)>" "INFO"
}

# Validate exclusions — warn if a configured UPN was not found in Exchange
foreach ($upn in $ExcludedSet) {
    $found = $AllMailboxes | Where-Object {
        $_.PrimarySmtpAddress.ToLower() -eq $upn -or $_.UserPrincipalName.ToLower() -eq $upn
    }
    if (-not $found) {
        Write-Log "WARNING: Configured exclusion '$upn' was NOT found in Exchange Online mailboxes." "WARN"
    }
}

#endregion ──────────────────────────────────────────────────────────────────────

#region ── APPLY CALENDAR SHARING PERMISSIONS ───────────────────────────────────

$Results   = [System.Collections.Generic.List[PSObject]]::new()
$SuccessCount = 0
$FailCount    = 0

Write-Log "Applying '$SharingPermission' calendar sharing to $($TargetMailboxes.Count) mailboxes..." "INFO"

foreach ($Mailbox in $TargetMailboxes) {

    $CalendarPath = "$($Mailbox.PrimarySmtpAddress):\Calendar"
    $DisplayInfo  = "$($Mailbox.DisplayName) <$($Mailbox.PrimarySmtpAddress)>"

    try {
        # Check if a permission entry already exists for this accessor
        $existing = Get-MailboxFolderPermission -Identity $CalendarPath `
                        -User $AccessorIdentity -ErrorAction SilentlyContinue

        if ($WhatIfMode) {
            $action = if ($existing) { "UPDATE" } else { "ADD" }
            Write-Log "  [WhatIf] Would $action calendar permission '$SharingPermission' for $DisplayInfo" "WHATIF"
            $Results.Add([PSCustomObject]@{
                DisplayName  = $Mailbox.DisplayName
                Email        = $Mailbox.PrimarySmtpAddress
                Action       = "WhatIf-$action"
                Permission   = $SharingPermission
                Status       = "Simulated"
                Error        = ""
            })
        }
        else {
            if ($existing) {
                Set-MailboxFolderPermission -Identity $CalendarPath `
                    -User $AccessorIdentity -AccessRights $SharingPermission -ErrorAction Stop
                Write-Log "  UPDATED: $DisplayInfo" "SUCCESS"
                $action = "Updated"
            }
            else {
                Add-MailboxFolderPermission -Identity $CalendarPath `
                    -User $AccessorIdentity -AccessRights $SharingPermission -ErrorAction Stop
                Write-Log "  ADDED: $DisplayInfo" "SUCCESS"
                $action = "Added"
            }
            $SuccessCount++
            $Results.Add([PSCustomObject]@{
                DisplayName  = $Mailbox.DisplayName
                Email        = $Mailbox.PrimarySmtpAddress
                Action       = $action
                Permission   = $SharingPermission
                Status       = "Success"
                Error        = ""
            })
        }
    }
    catch {
        $FailCount++
        Write-Log "  FAILED: $DisplayInfo — $_" "ERROR"
        $Results.Add([PSCustomObject]@{
            DisplayName  = $Mailbox.DisplayName
            Email        = $Mailbox.PrimarySmtpAddress
            Action       = "Failed"
            Permission   = $SharingPermission
            Status       = "Error"
            Error        = $_.Exception.Message
        })
    }
}

#endregion ──────────────────────────────────────────────────────────────────────

#region ── POST-RUN VERIFICATION ────────────────────────────────────────────────

if (-not $WhatIfMode) {
    Write-Log "---- Verification Pass: Spot-checking applied permissions ----" "INFO"
    $VerifyErrors = 0

    foreach ($r in ($Results | Where-Object { $_.Status -eq "Success" })) {
        $CalendarPath = "$($r.Email):\Calendar"
        try {
            $perm = Get-MailboxFolderPermission -Identity $CalendarPath `
                        -User $AccessorIdentity -ErrorAction Stop
            if ($perm.AccessRights -contains $SharingPermission) {
                Write-Log "  VERIFIED: $($r.DisplayName) — Permission '$SharingPermission' confirmed." "SUCCESS"
            }
            else {
                Write-Log "  MISMATCH: $($r.DisplayName) — Expected '$SharingPermission', got '$($perm.AccessRights)'." "WARN"
                $VerifyErrors++
            }
        }
        catch {
            Write-Log "  VERIFY FAILED: $($r.DisplayName) — $_" "WARN"
            $VerifyErrors++
        }
    }

    if ($VerifyErrors -eq 0) {
        Write-Log "Verification complete — all checked permissions confirmed." "SUCCESS"
    }
    else {
        Write-Log "Verification complete — $VerifyErrors permission(s) could not be confirmed. Review log." "WARN"
    }
}

#endregion ──────────────────────────────────────────────────────────────────────

#region ── EXCLUSION VERIFICATION ───────────────────────────────────────────────

if (-not $WhatIfMode) {
    Write-Log "---- Exclusion Check: Confirming exec calendars were NOT modified ----" "INFO"
    foreach ($ex in $ExcludedMailboxes) {
        $CalendarPath = "$($ex.PrimarySmtpAddress):\Calendar"
        try {
            $perm = Get-MailboxFolderPermission -Identity $CalendarPath `
                        -User $AccessorIdentity -ErrorAction SilentlyContinue
            if ($perm -and ($perm.AccessRights -contains $SharingPermission)) {
                Write-Log "  EXCLUSION BREACH: $($ex.DisplayName) has '$SharingPermission' — was this pre-existing?" "WARN"
            }
            else {
                Write-Log "  EXCLUSION OK: $($ex.DisplayName) — Not modified by this script." "SUCCESS"
            }
        }
        catch {
            Write-Log "  Could not check exclusion for $($ex.DisplayName): $_" "WARN"
        }
    }
}

#endregion ──────────────────────────────────────────────────────────────────────

#region ── EXPORT CSV REPORT & SUMMARY ──────────────────────────────────────────

$CsvFile = Join-Path $LogPath "CalendarSharing_Report_$Timestamp.csv"
$Results | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8
Write-Log "CSV report saved: $CsvFile" "INFO"

$Summary = @"
========== CALENDAR SHARING SCRIPT SUMMARY ==========
Run Mode         : $(if ($WhatIfMode) { 'SIMULATION (WhatIf)' } else { 'LIVE' })
Permission Set   : $SharingPermission
Accessor         : $AccessorIdentity
Total Mailboxes  : $($AllMailboxes.Count)
Excluded (Execs) : $($ExcludedMailboxes.Count)
Processed        : $($TargetMailboxes.Count)
Succeeded        : $SuccessCount
Failed           : $FailCount
Log File         : $LogFile
CSV Report       : $CsvFile
======================================================
"@

Write-Log $Summary "INFO"

#endregion ──────────────────────────────────────────────────────────────────────

#region ── EMAIL NOTIFICATION ────────────────────────────────────────────────────

if ($NotifyEmail) {
    $subject = if ($FailCount -gt 0) {
        "[WARNING] Calendar Sharing completed with $FailCount error(s)"
    } elseif ($WhatIfMode) {
        "[INFO] Calendar Sharing Simulation completed"
    } else {
        "[SUCCESS] Calendar Sharing completed for $SuccessCount mailbox(es)"
    }

    Send-Notification -To $NotifyEmail -Subject $subject -Body $Summary
}

#endregion ──────────────────────────────────────────────────────────────────────

#region ── DISCONNECT ────────────────────────────────────────────────────────────

try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Log "Disconnected from Exchange Online." "INFO"
}
catch {
    Write-Log "Disconnect warning (non-fatal): $_" "WARN"
}

Write-Log "========== Script Completed ==========" "INFO"

#endregion ──────────────────────────────────────────────────────────────────────
