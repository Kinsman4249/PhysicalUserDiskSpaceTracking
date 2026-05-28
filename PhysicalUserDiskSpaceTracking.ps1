#requires -version 5.1
<#
.SYNOPSIS
    PhysicalUserDiskSpaceTracking.ps1
    Measures the size-on-disk of the current user's profile folder and emails
    one of two recipients based on whether usage crosses a configurable threshold.

.DESCRIPTION
    - Runs in the user's logon context (no admin required to measure %USERPROFILE%).
    - Uses the Win32 API GetCompressedFileSizeW so that OneDrive "Files On-Demand"
      placeholders are counted at their on-disk (allocated) size, not their
      logical (cloud) size. This is the same fix that "Size on disk" in
      Windows Explorer applies.
    - Sends a single email per run via SMTP (with TLS) and exits silently.
    - Designed to be launched by RunHiddenPS.vbs from a Scheduled Task so no
      console window flashes on user logon.

.NOTES
    Author : (preserved from original FxLogixSizeAlert.ps1)
    History: Renamed from FxLogixSizeAlert.ps1 to PhysicalUserDiskSpaceTracking.ps1;
             logic preserved verbatim, comments added throughout.

.SECURITY
    !! WARNING !!
    The SMTP settings below are PLACEHOLDERS. Fill in your real server,
    username, and password locally before deploying. Do NOT commit real
    credentials to a public repository. For production, prefer reading
    credentials from environment variables, a credential vault, or a
    DPAPI-protected file rather than hard-coding them in this script.
#>

# Stop on any uncaught error. The outer try/catch around the mail send
# converts a thrown exception into a silent `exit 1` so Task Scheduler
# does not pop a UI for the end user.
$ErrorActionPreference = 'Stop'

# Hide the cmdlet progress bars (Write-Progress) that some pipelines emit.
# Task Scheduler is non-interactive, so the bar is just noise.
$ProgressPreference    = 'SilentlyContinue'

# =====================================================================
# HARD-CODED CONFIGURATION
# Edit these values for your environment. Everything below this block
# is logic; you should not normally need to touch it.
# =====================================================================

# ---- Folder to measure ----------------------------------------------
# %USERPROFILE% resolves to C:\Users\<currentUser> for whoever the task
# is running as. Because the Scheduled Task uses BUILTIN\Users as its
# principal, each logged-on user measures their OWN profile.
$Folder = $env:USERPROFILE

# ---- Default threshold (bytes) --------------------------------------
# 21,307,064,320 bytes is the original constant from FxLogixSizeAlert.ps1
# (~19.84 GiB). Use 21474836480 for a clean 20 GiB if you prefer.
$DefaultMinBytes = 21307064320

# ---- Per-user threshold multipliers ---------------------------------
# Some users legitimately need more profile space (developers, designers,
# execs with large mailboxes cached locally, etc.). List exceptions here
# as DOMAIN\samaccountname => multiplier. The default threshold is then
# multiplied by this number for that user only.
#
# Example: 'CONTOSO\jdoe' = 2  means jdoe is alerted at 2x the limit.
$UserLimitMultiplier = @{
    'CONTOSO\jdoe' = 2
}

# ---- Email recipients -----------------------------------------------
# $ToIfOver  receives the "Over threshold" alert (true positive — act on it).
# $ToIfUnder receives the "Under threshold" notification (used as a
#            heartbeat so you know the task ran for that user today).
$ToIfOver  = 'alerts@example.com'
$ToIfUnder = 'registrations@example.com'

# ---- SMTP server settings -------------------------------------------
# Using SMTP2GO on submission port 2525 with TLS. If you switch ESPs,
# update server, port, and credentials together. Port 587 is also common.
$SmtpServer   = 'mail.smtp2go.com'
$SmtpPort     = 2525
$SmtpUser     = 'REPLACE-WITH-SMTP-USERNAME'
$SmtpPassword = 'REPLACE-WITH-SMTP-PASSWORD'   # fill in locally; never commit a real value

# ---- From-address domain --------------------------------------------
# The script constructs the From header as <username>@<FromDomain> so
# replies route to the actual end-user mailbox.
$FromDomain = 'example.com'

# =====================================================================
# LOGIC (do not normally edit below this line)
# =====================================================================

# Defensive: if the user profile folder is missing for any reason
# (broken profile, race during logoff, etc.) just exit cleanly so we
# do not spam alerts.
if (-not (Test-Path -LiteralPath $Folder)) { exit 0 }

# Build the identity key used to look up per-user multipliers.
# Format: DOMAIN\username — matches how Windows refers to principals.
$principal = '{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME

# Apply the per-user multiplier if this principal is listed; otherwise
# fall back to the default. [int64] cast keeps the math 64-bit so it
# handles >2 GB values without overflow.
$minBytes = if ($UserLimitMultiplier.ContainsKey($principal)) {
    $DefaultMinBytes * [int64]$UserLimitMultiplier[$principal]
} else {
    $DefaultMinBytes
}

# ---------------------------------------------------------------------
# NATIVE "Size on disk" helper
# ---------------------------------------------------------------------
# Get-ChildItem | Measure-Object Length returns the LOGICAL size of each
# file. For OneDrive Files On-Demand placeholders, that includes cloud-only
# bytes that are NOT actually consuming disk. We want allocated bytes only.
#
# GetCompressedFileSizeW is the Win32 API that Windows Explorer itself
# calls to populate the "Size on disk" column. It returns the on-disk
# (allocated) size for sparse files, NTFS-compressed files, and OneDrive
# cloud-only placeholders.
#
# Add-Type compiles this small C# shim once per process and exposes it
# as [NativeSizeOnDisk]::SizeOnDisk("path").
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class NativeSizeOnDisk {
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    static extern uint GetCompressedFileSizeW(string lpFileName, out uint lpFileSizeHigh);

    public static ulong SizeOnDisk(string path) {
        uint high;
        uint low = GetCompressedFileSizeW(path, out high);
        if (low == 0xFFFFFFFF) {
            int err = Marshal.GetLastWin32Error();
            if (err != 0) throw new System.ComponentModel.Win32Exception(err);
        }
        return ((ulong)high << 32) | low;
    }
}
"@ -ErrorAction SilentlyContinue

# Accumulator for total bytes used by the profile.
$folderSize = 0L

# Walk every file under %USERPROFILE% (including hidden + system files)
# and add the on-disk size. -ErrorAction SilentlyContinue skips files
# we cannot stat (locked by another process, ACL-denied, etc.) — these
# are noise we do not want in the alert path.
Get-ChildItem -LiteralPath $Folder -Recurse -Force -File -ErrorAction SilentlyContinue |
ForEach-Object {
    try {
        # Cast result to [int64] (signed) and accumulate. SizeOnDisk
        # returns ulong; cast is safe for profiles under 8 EB.
        $folderSize += [int64][NativeSizeOnDisk]::SizeOnDisk($_.FullName)
    } catch {
        # Swallow individual file errors so a single locked file does not
        # abort the whole scan. This matches the original behaviour.
    }
}

# Belt-and-suspenders: if every file errored, $folderSize stays 0.
# That is fine, but normalise so downstream comparisons are unambiguous.
if (-not $folderSize) { $folderSize = 0 }

# ---------------------------------------------------------------------
# Compose mail
# ---------------------------------------------------------------------
# From: <username>@<domain> so the email looks like it came from the user.
# Body: just the From address — gives the recipient a quick visual of
#       who tripped the alert without needing to parse headers.
$from = '{0}@{1}' -f $env:USERNAME, $FromDomain
$body = $from

# Branch on threshold. Add an org tag to the subjects (e.g. '[ACME] ')
# if you rely on inbox rules / triage workflows to match on it.
if ($folderSize -ge $minBytes) {
    $to      = $ToIfOver
    $subject = 'User Profile Disk is Over 20 GB'
} else {
    $to      = $ToIfUnder
    $subject = 'User Profile Disk is NOT Over 20 GB'
}

# ---------------------------------------------------------------------
# Send mail
# ---------------------------------------------------------------------
# Wrapped in try/catch so a transient SMTP outage cannot surface as a
# scripting host error dialog on the user's desktop. Failures exit with
# code 1; successes exit with code 0. Task Scheduler logs the exit code
# under the Last Run Result column for debugging.
try {
    # Build a PSCredential. ConvertTo-SecureString -AsPlainText is OK here
    # only because the credential is already in-memory in plaintext from
    # the config block above.
    $securePwd = ConvertTo-SecureString $SmtpPassword -AsPlainText -Force
    $cred      = New-Object System.Management.Automation.PSCredential ($SmtpUser, $securePwd)

    # Construct the message via System.Net.Mail (works on PS 5.1).
    # Send-MailMessage is deprecated by Microsoft so we use the .NET
    # types directly.
    $msg = New-Object System.Net.Mail.MailMessage
    $msg.From = $from
    [void]$msg.To.Add($to)
    $msg.Subject = $subject
    $msg.Body    = $body

    # SmtpClient performs the actual submission. EnableSsl = STARTTLS
    # upgrade on the submission port.
    $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
    $smtp.EnableSsl   = $true
    $smtp.Credentials = $cred
    $smtp.Send($msg)

    exit 0
}
catch {
    # Silent failure: matches the original ">nul 2>&1" behaviour so the
    # user never sees a popup. If you need to debug, comment the catch
    # out temporarily and run the script from an interactive PowerShell.
    exit 1
}
