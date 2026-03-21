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
        # Ensure ExchangeOnlineManagement is installed
        $exoMissing = @($script:RequiredEXOModules | Where-Object {
            -not (Get-Module -ListAvailable -Name $_ -ErrorAction SilentlyContinue)
        })
        if ($exoMissing.Count -gt 0) {
            Write-MigrazeLog "Installing ExchangeOnlineManagement module..." "Action"
            Install-Module -Name "ExchangeOnlineManagement" -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-MigrazeLog "ExchangeOnlineManagement installed." "Success"
        }
        Import-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue
        Write-MigrazeLog "Opening browser for Microsoft 365 authentication..." "Action"
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $exoInfo = Get-ConnectionInformation -ErrorAction Stop
        if ($exoInfo -and $exoInfo.UserPrincipalName) {
            $script:IsGraphConnected = $true
            $script:IsEXOConnected   = $true
            $script:GraphAccount     = $exoInfo.UserPrincipalName
            $script:GraphTenantId    = $exoInfo.TenantID
            Write-MigrazeLog "Connected to Microsoft 365 as $($exoInfo.UserPrincipalName)." "Success"
            return @{ Success = $true; Account = $exoInfo.UserPrincipalName; TenantId = $exoInfo.TenantID }
        }
        throw "Login completed but no Exchange Online session found."
    } catch {
        $script:IsGraphConnected = $false
        $script:IsEXOConnected   = $false
        $script:GraphAccount     = $null
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
    $script:IsGraphConnected = $false
    $script:IsEXOConnected   = $false
    $script:GraphAccount     = $null
    $script:GraphTenantId    = $null
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
        Import-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue
        $exoInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($exoInfo -and $exoInfo.State -eq "Connected" -and $exoInfo.UserPrincipalName) {
            $script:IsGraphConnected = $true
            $script:IsEXOConnected   = $true
            $script:GraphAccount     = $exoInfo.UserPrincipalName
            $script:GraphTenantId    = $exoInfo.TenantID
            return @{ Connected = $true; Account = $exoInfo.UserPrincipalName; TenantId = $exoInfo.TenantID }
        }
    } catch {}
    $script:IsGraphConnected = $false
    $script:IsEXOConnected   = $false
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
        if ($DisplayName) { $params.DisplayName = $DisplayName }
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
        Invoke-EXOCommand
        Write-MigrazeLog "Searching users matching '$Query'..." "Action"
        $recipients = Get-Recipient `
            -Filter "DisplayName -like '*$Query*' -or PrimarySmtpAddress -like '*$Query*'" `
            -RecipientTypeDetails UserMailbox -ResultSize 30 -ErrorAction Stop
        $users = @($recipients | ForEach-Object {
            [PSCustomObject]@{
                Id                = $_.PrimarySmtpAddress
                DisplayName       = $_.DisplayName
                UserPrincipalName = $_.PrimarySmtpAddress
                Mail              = $_.PrimarySmtpAddress
            }
        })
        Write-MigrazeLog "Found $($users.Count) user(s) matching '$Query'." "Success"
        return @{ Success = $true; Users = $users }
    } catch {
        Write-MigrazeLog "Search-MigrazeUsers failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Connect-MigrazeTargetGraph {
    # Reserved for future M365-to-M365 migration scenario
    Write-MigrazeLog "Target tenant connection not yet implemented." "Warning"
    return @{ Success = $false; Error = "Not implemented." }
}

function Disconnect-MigrazeTargetGraph {
    $script:IsTargetGraphConnected = $false
    $script:TargetGraphAccount     = $null
    $script:TargetGraphTenantId    = $null
    Write-MigrazeLog "Disconnected from target tenant." "Info"
}