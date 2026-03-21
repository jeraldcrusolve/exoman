Set shell = CreateObject("WScript.Shell")
Dim scriptDir
scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
Dim psFile
psFile = scriptDir & "Migraze.ps1"
shell.Run "powershell.exe -STA -ExecutionPolicy Bypass -NoProfile -File """ & psFile & """", 0, False
