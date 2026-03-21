@ECHO OFF
SET "SCRIPT=%~dp0Migraze.ps1"
IF EXIST "%~dp0Launch-Migraze.vbs" (
    cscript //nologo "%~dp0Launch-Migraze.vbs"
) ELSE (
    powershell.exe -STA -ExecutionPolicy Bypass -NoProfile -File "%SCRIPT%"
    IF ERRORLEVEL 1 PAUSE
)
