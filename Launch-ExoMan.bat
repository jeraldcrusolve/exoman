@ECHO OFF
:: ExoMan v1.0 - Exchange Online Management Tool
:: Uses VBScript launcher for clean no-flash start. Falls back to direct launch.
IF EXIST "%~dp0Launch-ExoMan.vbs" (
    wscript.exe "%~dp0Launch-ExoMan.vbs"
) ELSE (
    powershell.exe -STA -ExecutionPolicy Bypass -NoProfile -File "%~dp0ExoMan.ps1"
    IF ERRORLEVEL 1 PAUSE
)
