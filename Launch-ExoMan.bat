# ExoMan – Exchange Online Management Tool
# Launch with: .\Launch-ExoMan.bat  OR  powershell -STA -ExecutionPolicy Bypass -File ExoMan.ps1

@ECHO OFF
TITLE ExoMan v1.0 - Launching...
powershell.exe -STA -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File "%~dp0ExoMan.ps1"
