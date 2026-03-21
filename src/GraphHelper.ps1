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

$script:LogBox        = $null
$script:LogEntryCount = 0
$script:LogCountLabel = $null
$script:LogAutoScroll = $null

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

function Write-ExoLog {
    param(
        [string]$Message,
        [ValidateSet("Info","Success","Error","Warning","Action")]
        [string]$Level = "Info"
    )
    Write-MigrazeLog -Message $Message -Level $Level
}

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
    foreach ($m in $script:RequiredModules) {
        Import-Module $m -Force -ErrorAction SilentlyContinue
    }
}

function Connect-MigrazeGraph {
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
            if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { throw "Required modules not installed." }
            Install-GraphModules -Modules $missing
        }
        Write-MigrazeLog "Loading Microsoft Graph modules..." "Info"
        Import-GraphModules
        Write-MigrazeLog "Opening browser for Microsoft 365 authentication..." "Action"
        Connect-MgGraph -Scopes $script:GraphScopes -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        if ($ctx -and $ctx.Account) {
            $script:IsGraphConnected = $true
            $script:GraphAccount     = $ctx.Account
            $script:GraphTenantId    = $ctx.TenantId
            Write-MigrazeLog "Connected to Microsoft 365." "Success"
            return @{ Success = $true; Account = $ctx.Account; TenantId = $ctx.TenantId }
        }
        throw "Login completed but no session context found."
    } catch {
        $script:IsGraphConnected = $false
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
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    $script:IsGraphConnected = $false
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

# ---- Distribution Group operations ----------------------------------------

function Get-DGList {
    param([string]$SearchQuery = "")
    try {
        Import-GraphModules
        if ($SearchQuery) {
            Write-MigrazeLog "Searching distribution groups: '$SearchQuery'..." "Action"
            $filter = "mailEnabled eq true and NOT groupTypes/any(c:c eq 'Unified') and (startsWith(displayName,'$SearchQuery') or startsWith(mail,'$SearchQuery'))"
        } else {
            Write-MigrazeLog "Fetching all distribution groups..." "Action"
            $filter = "mailEnabled eq true and NOT groupTypes/any(c:c eq 'Unified')"
        }
        $groups = Get-MgGroup -Filter $filter -Top 50 `
            -Property "Id,DisplayName,Mail,Description,MailNickname,MailEnabled,SecurityEnabled,CreatedDateTime" `
            -ErrorAction Stop
        Write-MigrazeLog "Found $($groups.Count) distribution group(s)." "Success"
        return @{ Success = $true; Groups = $groups }
    } catch {
        Write-MigrazeLog "Get-DGList failed: $($_.Exception.Message)" "Error"
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
        Import-GraphModules
        Write-MigrazeLog "Creating distribution group '$DisplayName' (alias: $MailNickname)..." "Action"
        $body = @{
            DisplayName     = $DisplayName
            MailNickname    = $MailNickname
            MailEnabled     = $true
            SecurityEnabled = $SecurityEnabled
            GroupTypes      = @()
        }
        if ($Description) { $body.Description = $Description }
        $group = New-MgGroup -BodyParameter $body -ErrorAction Stop
        Write-MigrazeLog "Distribution group created. ID: $($group.Id)" "Success"
        return @{ Success = $true; Group = $group }
    } catch {
        Write-MigrazeLog "New-DGGroup failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Update-DGGroup {
    param([string]$GroupId, [string]$DisplayName, [string]$Description)
    try {
        Import-GraphModules
        Write-MigrazeLog "Updating group properties for ID: $GroupId..." "Action"
        $body = @{}
        if ($DisplayName)           { $body.DisplayName   = $DisplayName }
        if ($null -ne $Description) { $body.Description   = $Description }
        Update-MgGroup -GroupId $GroupId -BodyParameter $body -ErrorAction Stop
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
        Import-GraphModules
        Write-MigrazeLog "Adding user $UserId to group $GroupId..." "Action"
        $ref = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" }
        New-MgGroupMemberByRef -GroupId $GroupId -BodyParameter $ref -ErrorAction Stop
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
        Import-GraphModules
        Write-MigrazeLog "Removing member $MemberId from group $GroupId..." "Action"
        Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $MemberId -ErrorAction Stop
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
        Import-GraphModules
        Write-MigrazeLog "Loading properties for group ID: $GroupId..." "Action"
        $group   = Get-MgGroup -GroupId $GroupId `
            -Property "Id,DisplayName,Mail,Description,MailNickname,MailEnabled,SecurityEnabled,GroupTypes,CreatedDateTime,Visibility" `
            -ErrorAction Stop
        Write-MigrazeLog "Loading members for '$($group.DisplayName)'..." "Info"
        $members = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop
        Write-MigrazeLog "Loaded $($members.Count) member(s) for '$($group.DisplayName)'." "Success"
        return @{ Success = $true; Group = $group; Members = $members }
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

function Search-ExoManUsers { param([string]$Query); Search-MigrazeUsers -Query $Query }