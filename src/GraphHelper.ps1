# GraphHelper.ps1 - Microsoft Graph authentication and shared utility functions

$script:IsGraphConnected       = $false
$script:GraphAccount           = $null
$script:GraphTenantId          = $null
$script:IsTargetGraphConnected = $false
$script:TargetGraphAccount     = $null
$script:TargetGraphTenantId    = $null

$script:GraphScopes = @(
    "User.Read",
    "User.ReadBasic.All",
    "Group.Read.All",
    "Group.ReadWrite.All",
    "GroupMember.Read.All",
    "GroupMember.ReadWrite.All",
    "Directory.Read.All",
    "OrgContact.Read.All"
)

$script:RequiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.Users"
)

$script:RequiredEXOModules = @("ExchangeOnlineManagement")
$script:IsEXOConnected = $false

$script:LogBox        = $null
$script:LogEntryCount = 0
$script:LogCountLabel = $null
$script:LogAutoScroll = $null

$script:IsTargetGraphConnected = $false
$script:TargetGraphAccount = $null
$script:TargetGraphTenantId = $null

function Write-MigrazeLog {
    param(
        [string]$Message,
        [ValidateSet("Info","Success","Error","Warning","Action")]
        [string]$Level = "Info"
    )
    if (-not $script:LogBox) { return }
    $ts     = Get-Date -Format "HH:mm:ss"
    $prefix = switch ($Level) {
        "Action"  { "ACTION " }
        "Info"    { "INFO   " }
        "Success" { "OK     " }
        "Warning" { "WARN   " }
        "Error"   { "ERROR  " }
    }
    $levelColor = switch ($Level) {
        "Action"  { "#38B2FF" }
        "Info"    { "#8AB4CC" }
        "Success" { "#00C853" }
        "Warning" { "#FFB300" }
        "Error"   { "#FF5252" }
    }
    $para            = [System.Windows.Documents.Paragraph]::new()
    $para.Margin     = [System.Windows.Thickness]::new(0)
    $para.LineHeight = 16
    $tsRun            = [System.Windows.Documents.Run]::new("[$ts] ")
    $tsRun.Foreground = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#3A5570"))
    $lvRun            = [System.Windows.Documents.Run]::new($prefix)
    $lvRun.Foreground = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString($levelColor))
    $lvRun.FontWeight = "Bold"
    $msgRun            = [System.Windows.Documents.Run]::new($Message)
    $msgRun.Foreground = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString(
        $(if ($Level -eq "Error") { "#FFAAAA" } elseif ($Level -eq "Success") { "#AAFFCC" } else { "#C8D8E8" })
    ))
    $para.Inlines.Add($tsRun)
    $para.Inlines.Add($lvRun)
    $para.Inlines.Add($msgRun)
    $script:LogBox.Document.Blocks.Add($para)
    $script:LogEntryCount++
    if ($script:LogCountLabel) { $script:LogCountLabel.Text = "  ($($script:LogEntryCount) entries)" }
    if ($script:LogAutoScroll -and $script:LogAutoScroll.IsChecked) { $script:LogBox.ScrollToEnd() }
}
Set-Alias -Name Write-ExoLog -Value Write-MigrazeLog -Scope Script

function Test-GraphModules {
    $missing = @()
    foreach ($m in $script:RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue)) { $missing += $m }
    }
    return $missing
}

function Install-GraphModules {
    param([string[]]$Modules)
    foreach ($m in $Modules) {
        Write-MigrazeLog "Installing module: $m ..." "Action"
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
        Write-MigrazeLog "Installed: $m" "Success"
    }
}

function Import-GraphModules {
    # EXO must be imported before Graph to win the MSAL AppDomain race
    Import-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue
    foreach ($m in $script:RequiredModules) {
        Import-Module $m -Force -ErrorAction SilentlyContinue
    }
}

function Connect-MigrazeGraph {
    try {
        # ── Step 1: Ensure ExchangeOnlineManagement is installed ──────────────
        $exoMissing = @($script:RequiredEXOModules | Where-Object {
            -not (Get-Module -ListAvailable -Name $_ -ErrorAction SilentlyContinue)
        })
        if ($exoMissing.Count -gt 0) {
            Write-MigrazeLog "Installing ExchangeOnlineManagement module..." "Action"
            Install-Module -Name "ExchangeOnlineManagement" -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-MigrazeLog "ExchangeOnlineManagement installed." "Success"
        }

        # ── Step 2: Ensure Microsoft.Graph modules are installed ──────────────
        $missing = Test-GraphModules
        if ($missing.Count -gt 0) {
            Write-MigrazeLog "Missing required modules: $($missing -join ', ')" "Warning"
            $answer = [System.Windows.MessageBox]::Show(
                "The following modules are required but not installed:`n`n$($missing -join "`n")`n`nInstall them now? (Requires internet access)",
                "Migraze - Missing Modules",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )
            if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { throw "Required modules not installed." }
            Install-GraphModules -Modules $missing
        }

        # ── Step 3: Import EXO FIRST so its MSAL wins the AppDomain race ─────
        Write-MigrazeLog "Loading modules..." "Info"
        Import-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue

        # ── Step 4: Connect to Exchange Online (browser prompt) ───────────────
        Write-MigrazeLog "Opening browser for Microsoft 365 authentication..." "Action"
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $exoInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        $upn = if ($exoInfo) { $exoInfo.UserPrincipalName } else { $null }
        $script:IsEXOConnected = $true
        Write-MigrazeLog "Connected to Exchange Online$(if ($upn) { " as $upn" })." "Success"

        # ── Step 5: Connect to Graph (SSO reuses cached AAD token) ────────────
        foreach ($m in $script:RequiredModules) {
            Import-Module $m -Force -ErrorAction SilentlyContinue
        }
        Write-MigrazeLog "Connecting to Microsoft Graph..." "Action"
        Connect-MgGraph -Scopes $script:GraphScopes -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        if ($ctx -and $ctx.Account) {
            $script:IsGraphConnected = $true
            $script:GraphAccount     = $ctx.Account
            $script:GraphTenantId    = $ctx.TenantId
            Write-MigrazeLog "Connected to Microsoft Graph." "Success"
            return @{ Success = $true; Account = $ctx.Account; TenantId = $ctx.TenantId }
        }
        throw "Graph login completed but no session context found."
    } catch {
        $script:IsGraphConnected = $false
        $script:GraphAccount     = $null
        $script:IsEXOConnected   = $false
        Write-MigrazeLog "M365 connection error: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Connect-MigrazeTargetGraph {
    try {
        $missing = Test-GraphModules
        if ($missing.Count -gt 0) {
            $answer = [System.Windows.MessageBox]::Show(
                "Required modules missing:`n`n$($missing -join "`n")`n`nInstall them now?",
                "Migraze - Missing Modules",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )
            if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { throw "Required modules not installed." }
            Install-GraphModules -Modules $missing
        }
        Import-GraphModules
        Write-MigrazeLog "Opening browser for TARGET M365 tenant authentication..." "Action"
        Connect-MgGraph -Scopes $script:GraphScopes -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        if ($ctx -and $ctx.Account) {
            $script:IsTargetGraphConnected = $true
            $script:TargetGraphAccount     = $ctx.Account
            $script:TargetGraphTenantId    = $ctx.TenantId
            Write-MigrazeLog "Connected to TARGET Microsoft 365 tenant." "Success"
            return @{ Success = $true; Account = $ctx.Account; TenantId = $ctx.TenantId }
        }
        throw "Login completed but no session context found."
    } catch {
        $script:IsTargetGraphConnected = $false
        $script:TargetGraphAccount     = $null
        Write-MigrazeLog "Target M365 connection error: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Disconnect-MigrazeGraph {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    $script:IsGraphConnected = $false
    $script:GraphAccount     = $null
    $script:GraphTenantId    = $null
    $script:IsEXOConnected   = $false
    Write-MigrazeLog "Disconnected from Microsoft 365." "Info"
}

function Disconnect-MigrazeTargetGraph {
    $script:IsTargetGraphConnected = $false
    $script:TargetGraphAccount     = $null
    $script:TargetGraphTenantId    = $null
    Write-MigrazeLog "Disconnected from target Microsoft 365 tenant." "Info"
}

function Get-MigrazeConnectionStatus {
    try {
        $ctx = Get-MgContext
        if ($ctx -and $ctx.Account) {
            $script:IsGraphConnected = $true
            $script:GraphAccount     = $ctx.Account
            $script:GraphTenantId    = $ctx.TenantId
            return @{ Connected = $true; Account = $ctx.Account; TenantId = $ctx.TenantId }
        }
    } catch {}
    $script:IsGraphConnected = $false
    return @{ Connected = $false }
}

# ---- Distribution Group operations (Exchange Online PowerShell) ------------

function Invoke-EXOCommand {
    # Ensure EXO module is imported before DG operations
    Import-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue
}

function Get-DGList {
    param([string]$SearchQuery = "")
    try {
        Invoke-EXOCommand
        if ($SearchQuery) {
            Write-MigrazeLog "Searching distribution groups: '$SearchQuery'..." "Action"
            $dgs = Get-DistributionGroup -Filter "DisplayName -like '*$SearchQuery*' -or PrimarySmtpAddress -like '*$SearchQuery*'" `
                -ResultSize 50 -ErrorAction Stop
        } else {
            Write-MigrazeLog "Fetching distribution groups..." "Action"
            $dgs = Get-DistributionGroup -ResultSize 50 -ErrorAction Stop
        }
        $normalized = @($dgs | ForEach-Object {
            [PSCustomObject]@{
                Id          = $_.PrimarySmtpAddress
                DisplayName = $_.DisplayName
                Mail        = $_.PrimarySmtpAddress
                Alias       = $_.Alias
                Description = $_.Notes
                ToString    = "$($_.DisplayName)  <$($_.PrimarySmtpAddress)>"
            }
        })
        Write-MigrazeLog "Found $($normalized.Count) distribution group(s)." "Success"
        return @{ Success = $true; Groups = $normalized }
    } catch {
        Write-MigrazeLog "Get-DGList failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-AllDGsForDiscovery {
    try {
        Invoke-EXOCommand
        Write-MigrazeLog "Discovering all distribution groups in tenant..." "Action"
        $dgs = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
        $normalized = @($dgs | ForEach-Object {
            [PSCustomObject]@{
                Id              = $_.PrimarySmtpAddress
                DisplayName     = $_.DisplayName
                Mail            = $_.PrimarySmtpAddress
                MailNickname    = $_.Alias
                Alias           = $_.Alias
                Description     = $_.Notes
                SecurityEnabled = ($_.GroupType -match "Security")
                MemberCount     = $_.GroupMemberCount
            }
        })
        Write-MigrazeLog "Discovery complete. Found $($normalized.Count) distribution group(s)." "Success"
        return @{ Success = $true; Groups = $normalized }
    } catch {
        Write-MigrazeLog "Discovery failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function New-DGGroup {
    param(
        [string]$DisplayName,
        [string]$MailNickname,
        [string]$Description     = "",
        [bool]  $SecurityEnabled = $false
    )
    try {
        Invoke-EXOCommand
        $type = if ($SecurityEnabled) { "Security" } else { "Distribution" }
        Write-MigrazeLog "Creating $type group '$DisplayName' (alias: $MailNickname)..." "Action"
        $group = New-DistributionGroup -Name $DisplayName -Alias $MailNickname -Type $type -ErrorAction Stop
        if ($Description) {
            Set-DistributionGroup -Identity $group.PrimarySmtpAddress -Notes $Description -ErrorAction SilentlyContinue
        }
        Write-MigrazeLog "Distribution group created: $($group.PrimarySmtpAddress)" "Success"
        return @{ Success = $true; Group = $group }
    } catch {
        Write-MigrazeLog "New-DGGroup failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Update-DGGroup {
    param([string]$GroupId, [string]$DisplayName, [string]$Description)
    try {
        Invoke-EXOCommand
        Write-MigrazeLog "Updating distribution group '$GroupId'..." "Action"
        $params = @{ Identity = $GroupId; ErrorAction = "Stop" }
        if ($DisplayName)           { $params.DisplayName = $DisplayName }
        if ($null -ne $Description -and $Description -ne "") { $params.Notes = $Description }
        Set-DistributionGroup @params
        Write-MigrazeLog "Group properties updated successfully." "Success"
        return @{ Success = $true }
    } catch {
        Write-MigrazeLog "Update-DGGroup failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Add-DGMember {
    param([string]$GroupId, [string]$UserId)
    try {
        Invoke-EXOCommand
        Write-MigrazeLog "Adding '$UserId' to group '$GroupId'..." "Action"
        Add-DistributionGroupMember -Identity $GroupId -Member $UserId -ErrorAction Stop
        Write-MigrazeLog "Member added successfully." "Success"
        return @{ Success = $true }
    } catch {
        Write-MigrazeLog "Add-DGMember failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Remove-DGMember {
    param([string]$GroupId, [string]$MemberId)
    try {
        Invoke-EXOCommand
        Write-MigrazeLog "Removing '$MemberId' from group '$GroupId'..." "Action"
        Remove-DistributionGroupMember -Identity $GroupId -Member $MemberId -Confirm:$false -ErrorAction Stop
        Write-MigrazeLog "Member removed successfully." "Success"
        return @{ Success = $true }
    } catch {
        Write-MigrazeLog "Remove-DGMember failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-DGProperties {
    param([string]$GroupId)
    try {
        Invoke-EXOCommand
        Write-MigrazeLog "Loading properties for group: $GroupId..." "Action"
        $dg = Get-DistributionGroup -Identity $GroupId -ErrorAction Stop
        Write-MigrazeLog "Loading members for '$($dg.DisplayName)'..." "Info"
        $members = Get-DistributionGroupMember -Identity $GroupId -ResultSize Unlimited -ErrorAction Stop
        $normalizedMembers = @($members | ForEach-Object {
            [PSCustomObject]@{
                Id          = $_.PrimarySmtpAddress
                DisplayName = $_.DisplayName
                ToString    = "$($_.DisplayName)  ($($_.PrimarySmtpAddress))"
            }
        })
        $normalizedGroup = [PSCustomObject]@{
            Id              = $dg.PrimarySmtpAddress
            DisplayName     = $dg.DisplayName
            Mail            = $dg.PrimarySmtpAddress
            Alias           = $dg.Alias
            Description     = $dg.Notes
            SecurityEnabled = ($dg.GroupType -match "Security")
        }
        Write-MigrazeLog "Loaded $($normalizedMembers.Count) member(s) for '$($dg.DisplayName)'." "Success"
        return @{ Success = $true; Group = $normalizedGroup; Members = $normalizedMembers }
    } catch {
        Write-MigrazeLog "Get-DGProperties failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Search-MigrazeUsers {
    param([string]$Query)
    try {
        Import-GraphModules
        Write-MigrazeLog "Searching users matching '$Query'..." "Action"
        $filter = "startsWith(displayName,'$Query') or startsWith(userPrincipalName,'$Query')"
        $users  = Get-MgUser -Filter $filter -Top 30 `
            -Property "Id,DisplayName,UserPrincipalName,Mail" -ErrorAction Stop
        Write-MigrazeLog "Found $($users.Count) user(s) matching '$Query'." "Success"
        return @{ Success = $true; Users = $users }
    } catch {
        Write-MigrazeLog "Search-MigrazeUsers failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Search-MigrazeUsers { param([string]$Query); Search-MigrazeUsers -Query $Query }
function Connect-MigrazeTargetGraph {
    <#
    .SYNOPSIS Opens a browser-based Microsoft 365 login for the target tenant.
    #>
    try {
        $missing = Test-GraphModules
        if ($missing.Count -gt 0) {
            Write-MigrazeLog "Missing required modules: $($missing -join ', ')" "Warning"
            $answer = [System.Windows.MessageBox]::Show(
                "The following modules are required but not installed:`n`n$($missing -join "`n")`n`nInstall them now? (Requires internet access)",
                "Migraze - Missing Modules",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )
            if ($answer -ne [System.Windows.MessageBoxResult]::Yes) {
                throw "Required modules not installed."
            }
            Install-GraphModules -Modules $missing
        }

        Write-MigrazeLog "Loading Microsoft Graph modules for target tenant..." "Info"
        Import-GraphModules

        Write-MigrazeLog "Opening browser for Target Tenant Microsoft 365 authentication..." "Action"
        Connect-MgGraph -Scopes $script:GraphScopes -NoWelcome -ErrorAction Stop

        $ctx = Get-MgContext
        if ($ctx -and $ctx.Account) {
            $script:IsTargetGraphConnected = $true
            $script:TargetGraphAccount     = $ctx.Account
            $script:TargetGraphTenantId    = $ctx.TenantId
            Write-MigrazeLog "Target tenant Graph connection established." "Success"
            return @{ Success = $true; Account = $ctx.Account; TenantId = $ctx.TenantId }
        }
        throw "Login completed but no session context found."

    } catch {
        $script:IsTargetGraphConnected = $false
        $script:TargetGraphAccount     = $null
        Write-MigrazeLog "Target connection error: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Disconnect-MigrazeTargetGraph {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    $script:IsTargetGraphConnected = $false
    $script:TargetGraphAccount     = $null
    $script:TargetGraphTenantId    = $null
    Write-MigrazeLog "Disconnected from target tenant." "Info"
}