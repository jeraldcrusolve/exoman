#Requires -Version 5.1
<#
.SYNOPSIS
    ExoMan v1.0 - Exchange Online Management Tool

.DESCRIPTION
    A Windows GUI tool for managing Exchange Online post-migration tasks.
    Uses Microsoft Graph PowerShell SDK for all operations.
    Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users

.NOTES
    Version : 1.0
    Tool    : ExoMan
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
            "ExoMan v1.0 - Startup Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } catch {
        # Write to a log file as last resort
        $logPath = Join-Path $PSScriptRoot "exoman-error.log"
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
    $script:AppVersion = "1.0"
    $script:AppTitle   = "ExoMan v1.0"

    # ---- Source modules (order matters) ----
    $srcFiles = @(
        "src\GraphHelper.ps1",
        "src\DistributionGroups.ps1",
        "src\SharedMailbox.ps1",
        "src\UserMailbox.ps1",
        "src\MainWindow.ps1"
    )

    foreach ($file in $srcFiles) {
        $path = Join-Path $script:AppRoot $file
        if (-not (Test-Path -LiteralPath $path)) {
            Show-FatalError "Required file missing:`n$path`n`nPlease reinstall ExoMan."
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
    Show-FatalError "ExoMan encountered an error and could not start.`n`nError: $errMsg`nFile : $errFile`nLine : $errLine`n`nCheck exoman-error.log for details."
    # Write full error to log
    try {
        $logPath = if ($script:AppRoot) { Join-Path $script:AppRoot "exoman-error.log" } else { "$env:TEMP\exoman-error.log" }
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  ERROR: $errMsg  [File: $errFile  Line: $errLine]`n$($_.ScriptStackTrace)" |
            Add-Content -Path $logPath
    } catch {}
    exit 1
}
