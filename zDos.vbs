Dim oShell
Set oShell = WScript.CreateObject ("WScript.shell")
Set WshShell = WScript.CreateObject("WScript.Shell")
oShell.run "cmd /K CD meuappxx & color 03 & cls"
'oShell.run "cmd /K CD F:\SYS\S0FTW4R35\000 My Soft\Windows Sysinternals\PsTools\ & color 0a & Dir"
Set oShell = Nothing