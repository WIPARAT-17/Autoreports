Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "C:\path\to\your\target\folder\run_Autoreport_Email.bat" & chr(34), 0

Set WshShell = Nothing
