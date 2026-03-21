' Launch-ExoMan.vbs
' Launches ExoMan.ps1 in STA PowerShell with no console window flash.

Dim fso, scriptDir, psScript, shell

Set fso       = CreateObject("Scripting.FileSystemObject")
scriptDir     = fso.GetParentFolderName(WScript.ScriptFullName)
psScript      = scriptDir & "\ExoMan.ps1"

If Not fso.FileExists(psScript) Then
    MsgBox "ExoMan.ps1 not found in:" & vbCrLf & scriptDir, 16, "ExoMan - Launch Error"
    WScript.Quit 1
End If

Set shell = CreateObject("WScript.Shell")
' 0 = hide window, False = don't wait (WPF app runs its own message loop)
shell.Run "powershell.exe -STA -ExecutionPolicy Bypass -NoProfile -File """ & psScript & """", 0, False

Set shell = Nothing
Set fso   = Nothing
