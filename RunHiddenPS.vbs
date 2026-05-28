Option Explicit

' =====================================================================
' RunHiddenPS.vbs
' ---------------------------------------------------------------------
' Purpose
'   Launch a PowerShell script with absolutely no visible window.
'   Running powershell.exe directly from Task Scheduler in user context
'   can briefly flash a black console even with -WindowStyle Hidden.
'   wscript.exe is a GUI-subsystem host, so when it spawns powershell.exe
'   the OS never allocates a console for it -> no flash, ever.
'
' Usage
'   wscript.exe //nologo "C:\ScripsDoNotDelete\RunHiddenPS.vbs" ^
'       "C:\ScripsDoNotDelete\PhysicalUserDiskSpaceTracking.ps1" [optional args...]
'
'   Argument 1 (required) : Full path to the .ps1 file to run.
'   Arguments 2+ (optional): Forwarded to the .ps1 unchanged.
'
' Notes
'   - WindowStyle 0 = SW_HIDE (no window at all).
'   - WaitOnReturn = True so Task Scheduler can track exit codes accurately
'     and so its "running" state matches the real lifetime of the script.
' =====================================================================

Dim shell, i, cmd

' WScript.Shell is what actually spawns child processes from VBScript.
Set shell = CreateObject("WScript.Shell")

' Refuse to run if no script path was passed. Exit code 87 = "The parameter
' is incorrect" (ERROR_INVALID_PARAMETER), which is what Task Scheduler
' shows as "Last Run Result" so the misconfiguration is visible.
If WScript.Arguments.Count < 1 Then
  WScript.Quit 87
End If

' Build the powershell.exe command line.
'   -NoProfile        : do not load $PROFILE (faster startup, predictable env)
'   -NonInteractive   : refuse to prompt; failures must be silent
'   -ExecutionPolicy Bypass : run unsigned scripts (this process only)
'   -WindowStyle Hidden     : belt-and-suspenders alongside the wscript hide
'   -File <ps1>             : the script to execute
cmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Q(WScript.Arguments(0))

' Forward any additional arguments (after the .ps1 path) to the script.
For i = 1 To WScript.Arguments.Count - 1
  cmd = cmd & " " & Q(WScript.Arguments(i))
Next

' shell.Run(command, windowStyle, waitOnReturn)
'   windowStyle  = 0    -> SW_HIDE
'   waitOnReturn = True -> block until powershell.exe exits
shell.Run cmd, 0, True

' ---------------------------------------------------------------------
' Q(s)
'   Quote a string for safe inclusion in a Windows command line.
'   Adds surrounding double-quotes if the value contains spaces or tabs,
'   and doubles any embedded double-quotes per the Win32 quoting rules.
' ---------------------------------------------------------------------
Function Q(s)
  If InStr(s, """") > 0 Then s = Replace(s, """", """""")
  If (InStr(s, " ") > 0) Or (InStr(s, vbTab) > 0) Then
    Q = """" & s & """"
  Else
    Q = s
  End If
End Function
