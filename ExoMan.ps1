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
    $args = "-STA -ExecutionPolicy Bypass -NoProfile -File `"$PSCommandPath`""
    Start-Process powershell.exe -ArgumentList $args -NoNewWindow
    exit
}

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

# ---- Load WPF assemblies ----
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ---- Global app state ----
$script:AppRoot    = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
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
    if (-not (Test-Path $path)) {
        [System.Windows.MessageBox]::Show(
            "Required file missing:`n$path`n`nPlease reinstall ExoMan.",
            "ExoMan - Startup Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        exit 1
    }
    . $path
}

# ---- Launch ----
Show-MainWindow
