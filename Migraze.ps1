#Requires -Version 5.1
<#
.SYNOPSIS
    Migraze v2.0 - Migration Management Platform

.DESCRIPTION
    A Windows GUI tool for managing mailbox and data migrations.
    Supports: Google Workspace to M365, M365 Tenant to Tenant Migration.
    Uses Microsoft Graph PowerShell SDK and Google Admin SDK REST API.

.NOTES
    Version : 2.0
    Tool    : Migraze
#>

# ---- Ensure STA threading model required by WPF ----
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $psFile  = $MyInvocation.MyCommand.Path
    $newArgs = "-STA -ExecutionPolicy Bypass -NoProfile -File `"$psFile`""
    Start-Process powershell.exe -ArgumentList $newArgs -WindowStyle Hidden
    exit
}

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

# ---- Helper: show a Windows Forms message box even before WPF loads ----
function Show-FatalError {
    param([string]$Message)
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show(
            $Message,
            "Migraze v2.0 - Startup Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } catch {
        $logPath = Join-Path $PSScriptRoot "migraze-error.log"
        "$(Get-Date)  $Message" | Add-Content -Path $logPath
    }
}

try {
    # ---- Load WPF assemblies ----
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName PresentationCore       -ErrorAction Stop
    Add-Type -AssemblyName WindowsBase            -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms   -ErrorAction Stop

    # ---- Resolve script root ----
    $script:AppRoot    = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent -LiteralPath $MyInvocation.MyCommand.Path }
    $script:AppVersion = "2.0"
    $script:AppTitle   = "Migraze"

    # ---- Source modules (order matters) ----
    $srcFiles = @(
        "src\GraphHelper.ps1",
        "src\GoogleHelper.ps1",
        "src\Discovery\GoogleDiscovery.ps1",
        "src\Discovery\M365Discovery.ps1",
        "src\Management\DistributionGroups.ps1",
        "src\Management\SharedMailbox.ps1",
        "src\Management\UserMailbox.ps1",
        "src\Scenarios\GW-to-M365.ps1",
        "src\Scenarios\M365-to-M365.ps1",
        "src\MainWindow.ps1"
    )

    foreach ($file in $srcFiles) {
        $path = Join-Path $script:AppRoot $file
        if (-not (Test-Path -LiteralPath $path)) {
            Show-FatalError "Required file missing:`n$path`n`nPlease reinstall Migraze."
            exit 1
        }
        . $path
    }

    # ---- Launch ----
    Show-MainWindow

} catch {
    $errMsg  = $_.Exception.Message
    $errLine = $_.InvocationInfo.ScriptLineNumber
    $errFile = $_.InvocationInfo.ScriptName
    Show-FatalError "Migraze encountered an error and could not start.`n`nError: $errMsg`nFile : $errFile`nLine : $errLine`n`nCheck migraze-error.log for details."
    try {
        $logPath = if ($script:AppRoot) { Join-Path $script:AppRoot "migraze-error.log" } else { "$env:TEMP\migraze-error.log" }
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  ERROR: $errMsg  [File: $errFile  Line: $errLine]`n$($_.ScriptStackTrace)" |
            Add-Content -Path $logPath
    } catch {}
    exit 1
}
