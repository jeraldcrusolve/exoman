# GraphHelper.ps1 - Microsoft Graph authentication and shared utility functions

$script:IsGraphConnected = $false
$script:GraphAccount     = $null
$script:GraphTenantId    = $null

$script:GraphScopes = @(
    "User.Read",
    "User.ReadBasic.All",
    "Group.Read.All",
    "Group.ReadWrite.All",
    "GroupMember.Read.All",
    "GroupMember.ReadWrite.All",
    "Directory.Read.All"
)

$script:RequiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.Users"
)

# ── Shared activity log (set by MainWindow after it loads) ──
$script:LogBox        = $null
$script:LogEntryCount = 0
$script:LogCountLabel = $null
$script:LogAutoScroll = $null

function Write-ExoLog {
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

    $para           = [System.Windows.Documents.Paragraph]::new()
    $para.Margin    = [System.Windows.Thickness]::new(0)
    $para.LineHeight = 16

    $tsRun           = [System.Windows.Documents.Run]::new("[$ts] ")
    $tsRun.Foreground = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#3A5570"))

    $lvRun           = [System.Windows.Documents.Run]::new($prefix)
    $lvRun.Foreground = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString($levelColor))
    $lvRun.FontWeight = "Bold"

    $msgRun           = [System.Windows.Documents.Run]::new($Message)
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

function Test-GraphModules {
    $missing = @()
    foreach ($m in $script:RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue)) {
            $missing += $m
        }
    }
    return $missing
}

function Install-GraphModules {
    param([string[]]$Modules)
    foreach ($m in $Modules) {
        Write-ExoLog "Installing module: $m ..." "Action"
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
        Write-ExoLog "Installed: $m" "Success"
    }
}

function Import-GraphModules {
    foreach ($m in $script:RequiredModules) {
        Import-Module $m -Force -ErrorAction SilentlyContinue
    }
}

function Connect-ExoManGraph {
    <#
    .SYNOPSIS Opens a browser-based Microsoft 365 login and connects to Microsoft Graph.
    #>
    try {
        $missing = Test-GraphModules
        if ($missing.Count -gt 0) {
            Write-ExoLog "Missing required modules: $($missing -join ', ')" "Warning"
            $answer = [System.Windows.MessageBox]::Show(
                "The following modules are required but not installed:`n`n$($missing -join "`n")`n`nInstall them now? (Requires internet access)",
                "ExoMan – Missing Modules",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )
            if ($answer -ne [System.Windows.MessageBoxResult]::Yes) {
                throw "Required modules not installed."
            }
            Install-GraphModules -Modules $missing
        }

        Write-ExoLog "Loading Microsoft Graph modules..." "Info"
        Import-GraphModules

        Write-ExoLog "Opening browser for Microsoft 365 authentication..." "Action"
        # Opens the default browser for interactive Microsoft 365 login
        Connect-MgGraph -Scopes $script:GraphScopes -NoWelcome -ErrorAction Stop

        $ctx = Get-MgContext
        if ($ctx -and $ctx.Account) {
            $script:IsGraphConnected = $true
            $script:GraphAccount     = $ctx.Account
            $script:GraphTenantId    = $ctx.TenantId
            Write-ExoLog "Graph connection established." "Success"
            return @{ Success = $true; Account = $ctx.Account; TenantId = $ctx.TenantId }
        }
        throw "Login completed but no session context found."

    } catch {
        $script:IsGraphConnected = $false
        $script:GraphAccount     = $null
        Write-ExoLog "Connection error: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Disconnect-ExoManGraph {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    $script:IsGraphConnected = $false
    $script:GraphAccount     = $null
    $script:GraphTenantId    = $null
}

function Get-ExoManConnectionStatus {
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
            Write-ExoLog "Searching distribution groups: '$SearchQuery'..." "Action"
            $filter = "mailEnabled eq true and NOT groupTypes/any(c:c eq 'Unified') and (startsWith(displayName,'$SearchQuery') or startsWith(mail,'$SearchQuery'))"
        } else {
            Write-ExoLog "Fetching all distribution groups..." "Action"
            $filter = "mailEnabled eq true and NOT groupTypes/any(c:c eq 'Unified')"
        }
        $groups = Get-MgGroup -Filter $filter -Top 50 `
            -Property "Id,DisplayName,Mail,Description,MailNickname,MailEnabled,SecurityEnabled,CreatedDateTime" `
            -ErrorAction Stop
        Write-ExoLog "Found $($groups.Count) distribution group(s)." "Success"
        return @{ Success = $true; Groups = $groups }
    } catch {
        Write-ExoLog "Get-DGList failed: $($_.Exception.Message)" "Error"
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
        Write-ExoLog "Creating distribution group '$DisplayName' (alias: $MailNickname)..." "Action"
        $body = @{
            DisplayName     = $DisplayName
            MailNickname    = $MailNickname
            MailEnabled     = $true
            SecurityEnabled = $SecurityEnabled
            GroupTypes      = @()
        }
        if ($Description) { $body.Description = $Description }
        $group = New-MgGroup -BodyParameter $body -ErrorAction Stop
        Write-ExoLog "Distribution group created. ID: $($group.Id)" "Success"
        return @{ Success = $true; Group = $group }
    } catch {
        Write-ExoLog "New-DGGroup failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Update-DGGroup {
    param(
        [string]$GroupId,
        [string]$DisplayName,
        [string]$Description
    )
    try {
        Import-GraphModules
        Write-ExoLog "Updating group properties for ID: $GroupId..." "Action"
        $body = @{}
        if ($DisplayName) { $body.DisplayName   = $DisplayName }
        if ($null -ne $Description) { $body.Description = $Description }
        Update-MgGroup -GroupId $GroupId -BodyParameter $body -ErrorAction Stop
        Write-ExoLog "Group properties updated successfully." "Success"
        return @{ Success = $true }
    } catch {
        Write-ExoLog "Update-DGGroup failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Add-DGMember {
    param([string]$GroupId, [string]$UserId)
    try {
        Import-GraphModules
        Write-ExoLog "Adding user $UserId to group $GroupId..." "Action"
        $ref = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" }
        New-MgGroupMemberByRef -GroupId $GroupId -BodyParameter $ref -ErrorAction Stop
        Write-ExoLog "Member added successfully." "Success"
        return @{ Success = $true }
    } catch {
        Write-ExoLog "Add-DGMember failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Remove-DGMember {
    param([string]$GroupId, [string]$MemberId)
    try {
        Import-GraphModules
        Write-ExoLog "Removing member $MemberId from group $GroupId..." "Action"
        Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $MemberId -ErrorAction Stop
        Write-ExoLog "Member removed successfully." "Success"
        return @{ Success = $true }
    } catch {
        Write-ExoLog "Remove-DGMember failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-DGProperties {
    param([string]$GroupId)
    try {
        Import-GraphModules
        Write-ExoLog "Loading properties for group ID: $GroupId..." "Action"
        $group   = Get-MgGroup -GroupId $GroupId `
            -Property "Id,DisplayName,Mail,Description,MailNickname,MailEnabled,SecurityEnabled,GroupTypes,CreatedDateTime,Visibility" `
            -ErrorAction Stop
        Write-ExoLog "Loading members for '$($group.DisplayName)'..." "Info"
        $members = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop
        Write-ExoLog "Loaded $($members.Count) member(s) for '$($group.DisplayName)'." "Success"
        return @{ Success = $true; Group = $group; Members = $members }
    } catch {
        Write-ExoLog "Get-DGProperties failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Search-ExoManUsers {
    param([string]$Query)
    try {
        Import-GraphModules
        Write-ExoLog "Searching users matching '$Query'..." "Action"
        $filter = "startsWith(displayName,'$Query') or startsWith(userPrincipalName,'$Query')"
        $users  = Get-MgUser -Filter $filter -Top 30 `
            -Property "Id,DisplayName,UserPrincipalName,Mail" -ErrorAction Stop
        Write-ExoLog "Found $($users.Count) user(s) matching '$Query'." "Success"
        return @{ Success = $true; Users = $users }
    } catch {
        Write-ExoLog "Search-ExoManUsers failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}
