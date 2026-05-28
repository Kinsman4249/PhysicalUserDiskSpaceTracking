# PhysicalUserDiskSpaceTracking

> Per-user profile disk-usage monitoring for Windows 10/11, AVD, RDS and any
> multi-session host. Measures `%USERPROFILE%` at logon and emails one of two
> recipients depending on whether the user has crossed a configurable
> threshold.

[![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011%20%7C%20AVD%20%7C%20RDS-0078D6?logo=windows)](https://learn.microsoft.com/en-us/windows-server/remote/remote-desktop-services/)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell)](https://learn.microsoft.com/en-us/powershell/)
[![License](https://img.shields.io/badge/license-All%20Rights%20Reserved-red)](#license)

---

## Contents

1. [What this is](#what-this-is)
2. [How it works](#how-it-works)
3. [Files in this repo](#files-in-this-repo)
4. [Prerequisites](#prerequisites)
5. [Install](#install)
6. [Configuration](#configuration)
7. [Per-user threshold overrides](#per-user-threshold-overrides)
8. [Manual test run](#manual-test-run)
9. [Uninstall](#uninstall)
10. [Reusing the installer for other scripts](#reusing-the-installer-for-other-scripts)
11. [Security notes](#security-notes)
12. [Contributing](#contributing)
13. [License](#license)

---

## What this is

A three-file PowerShell + VBScript package that, on every interactive logon,
sums the on-disk size of the signed-in user's profile folder and emails one
of two distribution addresses depending on whether that user has crossed a
configurable threshold.

It is intended for use on multi-session Windows hosts (Azure Virtual Desktop,
Remote Desktop Services, shared workstations) where profile bloat is a
recurring operational concern.

---

## How it works

1. A Scheduled Task triggers **At logon**, for any user in `BUILTIN\Users`.
2. The task launches `wscript.exe` against `RunHiddenPS.vbs`. Using `wscript`
   instead of `powershell.exe` directly guarantees no console window flashes
   on the user's screen.
3. `RunHiddenPS.vbs` spawns `powershell.exe -WindowStyle Hidden -File
   PhysicalUserDiskSpaceTracking.ps1`.
4. The PowerShell script walks `%USERPROFILE%` and sums each file's
   **size on disk** (allocated bytes) by calling the Win32
   `GetCompressedFileSizeW` API. This is the same calculation Explorer
   uses, and correctly excludes cloud-only OneDrive Files On-Demand
   placeholders.
5. If the total is at or above `$DefaultMinBytes` (with optional per-user
   multiplier), the script emails `$ToIfOver`. Otherwise it emails
   `$ToIfUnder` as a heartbeat.

---

## Files in this repo

| File | Purpose |
| --- | --- |
| `PhysicalUserDiskSpaceTracking.ps1` | The measurement + alert script. Runs as the signed-in user. |
| `RunHiddenPS.vbs` | Generic VBS wrapper that launches any PowerShell script with no visible window. |
| `Install-ScheduledTask.ps1` | Registers the Scheduled Task that wires everything together. |

---

## Prerequisites

- **Windows 10 / 11 / Server 2016+** with Windows PowerShell **5.1** (PowerShell 7 also works).
- An SMTP relay reachable from each host. The script ships with **placeholder** credentials — fill in your own before deploying.
- **Local administrator rights** to register the Scheduled Task (one-time, on each host).
- A consistent install path. The defaults assume:

  ```
  C:\ScripsDoNotDelete\RunHiddenPS.vbs
  C:\ScripsDoNotDelete\PhysicalUserDiskSpaceTracking.ps1
  C:\ScripsDoNotDelete\Install-ScheduledTask.ps1
  ```

  > **Note:** `ScripsDoNotDelete` is intentionally spelled without the "t" in
  > "Scrips" — that matches the existing convention. Either rename the folder
  > AND every path constant in the scripts, or leave it as-is.

---

## Install

1. Create the folder and copy the three files onto the host:

   ```powershell
   New-Item -ItemType Directory -Path 'C:\ScripsDoNotDelete' -Force | Out-Null
   Copy-Item .\PhysicalUserDiskSpaceTracking.ps1 'C:\ScripsDoNotDelete\'
   Copy-Item .\RunHiddenPS.vbs                  'C:\ScripsDoNotDelete\'
   Copy-Item .\Install-ScheduledTask.ps1        'C:\ScripsDoNotDelete\'
   ```

2. Edit `PhysicalUserDiskSpaceTracking.ps1` and fill in your SMTP credentials,
   from-domain, and recipient addresses. See [Configuration](#configuration).

3. Open **PowerShell as Administrator** and register the task:

   ```powershell
   Set-Location C:\ScripsDoNotDelete
   .\Install-ScheduledTask.ps1
   ```

4. Confirm in **Task Scheduler → Task Scheduler Library →
   UserProfileCleanup → PhysicalUserDiskSpaceTracking** that the task exists
   and is enabled.

The next time any user signs in to this host, the task will run for them.

---

## Configuration

All knobs live in the top of `PhysicalUserDiskSpaceTracking.ps1`:

| Variable | Purpose | Default |
| --- | --- | --- |
| `$Folder` | Folder to measure. Leave as `$env:USERPROFILE` for per-user profiles. | `$env:USERPROFILE` |
| `$DefaultMinBytes` | Threshold in bytes. The default is ~19.84 GiB; use `21474836480` for a clean 20 GiB. | `21307064320` |
| `$UserLimitMultiplier` | Per-user threshold overrides (see below). | `@{ 'CONTOSO\jdoe' = 2 }` |
| `$ToIfOver` | Recipient when the user is **over** the threshold (the real alert). | `alerts@example.com` |
| `$ToIfUnder` | Recipient when the user is **under** the threshold (heartbeat). | `registrations@example.com` |
| `$SmtpServer` / `$SmtpPort` | SMTP relay + submission port. TLS is enabled. | `mail.smtp2go.com` / `2525` |
| `$SmtpUser` / `$SmtpPassword` | SMTP credential. **Fill in your own before deploying.** | `REPLACE-WITH-SMTP-USERNAME` / `REPLACE-WITH-SMTP-PASSWORD` |
| `$FromDomain` | Domain appended to `$env:USERNAME` to form the `From:` header. | `example.com` |

---

## Per-user threshold overrides

Some users legitimately need more profile space. Edit `$UserLimitMultiplier`
to multiply the default threshold for specific principals:

```powershell
$UserLimitMultiplier = @{
    'CONTOSO\jdoe'   = 2   # alert jdoe at 2x the default
    'CONTOSO\asmith' = 3   # alert asmith at 3x the default
}
```

The lookup key is built at runtime as:

```powershell
$principal = "$env:USERDOMAIN\$env:USERNAME"
```

If `$principal` is present in the hashtable, the effective threshold becomes
`$DefaultMinBytes * multiplier`. Otherwise the default applies.

---

## Manual test run

You can fire the same code path that the Scheduled Task uses, without
waiting for a logon:

```cmd
wscript.exe //nologo "C:\ScripsDoNotDelete\RunHiddenPS.vbs" "C:\ScripsDoNotDelete\PhysicalUserDiskSpaceTracking.ps1"
```

For an interactive debug run (with console output and stack traces),
bypass the VBS wrapper and run the .ps1 directly:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\ScripsDoNotDelete\PhysicalUserDiskSpaceTracking.ps1
```

---

## Uninstall

```powershell
Unregister-ScheduledTask `
    -TaskName 'PhysicalUserDiskSpaceTracking' `
    -TaskPath '\\UserProfileCleanup\\' `
    -Confirm:$false
```

If you changed `$TaskPath` in `Install-ScheduledTask.ps1`, pass the same
value here.

---

## Reusing the installer for other scripts

`Install-ScheduledTask.ps1` is intentionally written as a thin shell over
`New-ScheduledTask*` cmdlets so you can clone it for any other PowerShell
script you want to fire at logon. To repurpose it:

### Step 1 — Copy the installer and rename it

Pick a descriptive name and keep the `Install-` prefix for clarity:

```
Install-MyOtherScriptTask.ps1
```

### Step 2 — Edit the five variables at the top of the CONFIGURATION block

These are the only values you normally need to change.

```powershell
# Name shown in Task Scheduler. Must be unique within $TaskPath.
$TaskName = 'MyOtherScriptTask'

# Task Scheduler folder. Use '\\' for the root library, or any sub-folder.
# Group related tasks under the same path so they are easy to find.
$TaskPath = '\\MyCompanyTasks\\'

# Absolute paths to the VBS wrapper and the target .ps1.
$VbsPath  = 'C:\\ScripsDoNotDelete\\RunHiddenPS.vbs'
$Ps1Path  = 'C:\\ScripsDoNotDelete\\MyOtherScript.ps1'

# Working directory the task starts in.
$WorkDir  = 'C:\\ScripsDoNotDelete'
```

> **Tip:** `RunHiddenPS.vbs` is fully generic — it takes the .ps1 path as
> its first argument. You do **not** need a separate copy of the VBS for
> each script. Point every installer at the same `RunHiddenPS.vbs`.

### Step 3 — Decide who the task should run as

By default the installer uses `BUILTIN\Users`, which means *the currently
signed-in user* runs the task in their own context. That is correct for
anything that needs to touch `%USERPROFILE%` or `HKCU`.

If your script needs elevated rights, swap the principal block:

```powershell
# Run as SYSTEM (no UI, full machine rights)
$Principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest

# Or as a specific service account
$Principal = New-ScheduledTaskPrincipal -UserId 'CONTOSO\svc-tasks' -LogonType Password -RunLevel Highest
```

### Step 4 — Decide when the task should run

Replace the `New-ScheduledTaskTrigger` line. Common patterns:

```powershell
# At every user logon (default)
$Trigger = New-ScheduledTaskTrigger -AtLogOn

# Only when a specific user signs in
$Trigger = New-ScheduledTaskTrigger -AtLogOn -User 'CONTOSO\jdoe'

# At system startup (combine with SYSTEM principal)
$Trigger = New-ScheduledTaskTrigger -AtStartup

# Daily at 03:00
$Trigger = New-ScheduledTaskTrigger -Daily -At '3:00am'

# Every 15 minutes, forever, starting now
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) `
            -RepetitionInterval (New-TimeSpan -Minutes 15) `
            -RepetitionDuration ([TimeSpan]::MaxValue)
```

### Step 5 — Decide how concurrent runs should behave

```powershell
# Parallel (default) — fire a new instance even if one is already running.
# Use this for short, per-user tasks on multi-session hosts.
$Settings = New-ScheduledTaskSettingsSet -MultipleInstances Parallel

# Queue — wait for the running instance to finish, then run the next one.
$Settings = New-ScheduledTaskSettingsSet -MultipleInstances Queue

# IgnoreNew — drop overlapping fires. Good for "hourly housekeeping" tasks.
$Settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew
```

### Step 6 — Install

Run the new installer **elevated**:

```powershell
.\Install-MyOtherScriptTask.ps1
```

Re-running it overwrites the existing task in place (because the registration
uses `-Force`), so you can iterate on the installer without manually
unregistering each time.

### Step 7 — Verify

```powershell
Get-ScheduledTask -TaskPath '\MyCompanyTasks\' | Format-Table TaskName, State
Get-ScheduledTaskInfo -TaskName 'MyOtherScriptTask' -TaskPath '\MyCompanyTasks\'
```

`LastRunTime` and `LastTaskResult` will tell you whether the task fired and
what exit code the script returned. `0` = success, anything else = failure.

---

## Security notes

> **SMTP credentials are stored in the script.** `PhysicalUserDiskSpaceTracking.ps1`
> ships with **placeholder** values (`REPLACE-WITH-SMTP-USERNAME` /
> `REPLACE-WITH-SMTP-PASSWORD`) that you fill in locally before deploying.
> Once you do, the username and password live in the .ps1 as plain string
> literals — anyone who can read the file can read those credentials. To keep
> them safe, do at least one of:
>
> 1. **Never commit real credentials.** Keep the placeholders in source control
>    and inject the real values only on the host (or via a deployment step).
> 2. **Refactor** to read from environment variables, a DPAPI-encrypted
>    file, or a managed secret store (Azure Key Vault, Windows Credential
>    Manager).
> 3. **Restrict NTFS permissions** on `C:\ScripsDoNotDelete\` to admins +
>    SYSTEM only, so non-admin users on the box cannot read the file.
> 4. **Rotate** the SMTP password immediately if a real value was ever
>    committed or shared.

The script intentionally exits silently on error so end users never see a
popup. This means failures are only visible via Task Scheduler's
`LastTaskResult` column or by un-commenting the `catch` block during
debugging.

---

## Contributing

Repo-level contribution conventions live in the organisation-wide
[`Kinsman4249/.github-private`](https://github.com/Kinsman4249/.github-private)
repository. That includes:

- `CONTRIBUTING.md` — branching, PR, and review conventions
- `CODE_OF_CONDUCT.md`
- `SECURITY.md` — how to report a vulnerability
- `SUPPORT.md`
- Default issue and pull-request templates

Those files are inherited automatically by every repository in the org;
you do not need to duplicate them here.

---

## License

**All Rights Reserved.** Copyright (c) 2026 Ethan Antonio.

This is proprietary software — see [`LICENSE`](LICENSE) for the full terms. No
use, copying, modification, or distribution is permitted without the copyright
holder's prior written consent.
