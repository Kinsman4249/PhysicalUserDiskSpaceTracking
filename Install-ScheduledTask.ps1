#requires -version 5.1

<
.SYNOPSIS
  Installs a Scheduled Task that runs a PowerShell script at user logon for all members of BUILTIN\Users,
  using wscript.exe + RunHiddenPS.vbs to avoid a visible PowerShell window.

.DESCRIPTION
  - Trigger: At logon (any user)
  - Principal: BUILTIN\Users (group)
  - Multiple instances: Parallel
  - Action: wscript.exe //nologo "<VBS>" "<PS1>"

.NOTES
  Run elevated to register a task under a shared task path.

  To repurpose this installer for another script, see the
  "Reusing the installer for other scripts" section of README.md.
>

# --- Config ---
# Five values you normally edit when repurposing this installer for another script.
# TaskName must be unique within $TaskPath; $TaskPath is the folder inside Task Scheduler's library.
$TaskName = 'PhysicalUserDiskSpaceTracking'
$TaskPath = '\\UserProfileCleanup\\'   # change to '\\' if you want it in the root library

# Absolute paths to the VBS wrapper (which hides the PowerShell window) and the .ps1 it should launch.
# RunHiddenPS.vbs is fully generic; the .ps1 path is the only thing that changes when you reuse it.
$VbsPath  = 'C:\\ScripsDoNotDelete\\RunHiddenPS.vbs'
$Ps1Path  = 'C:\\ScripsDoNotDelete\\PhysicalUserDiskSpaceTracking.ps1'
$WorkDir  = 'C:\\ScripsDoNotDelete'

# --- Action: run VBS wrapper (GUI host) which runs the PS1 hidden ---
# wscript.exe is a GUI-subsystem host, so the OS never allocates a console for it -> no flash.
# //nologo suppresses the Windows Script Host banner.
# The two quoted format arguments become the VBS path and the PS1 path passed to it.
$Action = New-ScheduledTaskAction \
  -Execute 'wscript.exe' \
  -Argument ("//nologo `"{0}`" `"{1}`"" -f $VbsPath, $Ps1Path) \
  -WorkingDirectory $WorkDir

# --- Trigger: at logon (no specific user set => any user logon) ---
# Combined with the BUILTIN\Users principal below, every interactive logon fires its own run
# against that user's own profile.
$Trigger = New-ScheduledTaskTrigger -AtLogOn

# --- Principal: BUILTIN\Users group ---
# Running as BUILTIN\Users means the task runs in the logged-on user's context with standard
# rights — required so %USERPROFILE% resolves to the right folder. Swap for 'SYSTEM' (with
# -LogonType ServiceAccount -RunLevel Highest) if the script needs machine-wide privileges.
$Principal = New-ScheduledTaskPrincipal -GroupId 'BUILTIN\Users'

# --- Settings: allow parallel instances ---
# Parallel: fire a new instance even if a prior user's run is still going. Critical on
# multi-session hosts (AVD/RDS) where multiple users can sign in within seconds of each other.
# Alternatives: Queue (wait for the running one) or IgnoreNew (drop overlapping fires).
$Settings = New-ScheduledTaskSettingsSet -MultipleInstances Parallel

# --- Register ---
# -Force overwrites any existing task with the same name in $TaskPath, so re-running this
# installer upgrades the task in place instead of erroring.
Register-ScheduledTask \
  -TaskName $TaskName \
  -TaskPath $TaskPath \
  -Action $Action \
  -Trigger $Trigger \
  -Principal $Principal \
  -Settings $Settings \
  -Force
