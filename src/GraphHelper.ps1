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
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
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

        Import-GraphModules

        # Opens the default browser for interactive Microsoft 365 login
        Connect-MgGraph -Scopes $script:GraphScopes -NoWelcome -ErrorAction Stop

        $ctx = Get-MgContext
        if ($ctx -and $ctx.Account) {
            $script:IsGraphConnected = $true
            $script:GraphAccount     = $ctx.Account
            $script:GraphTenantId    = $ctx.TenantId
            return @{ Success = $true; Account = $ctx.Account; TenantId = $ctx.TenantId }
        }
        throw "Login completed but no session context found."

    } catch {
        $script:IsGraphConnected = $false
        $script:GraphAccount     = $null
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
        $filter = "mailEnabled eq true and NOT groupTypes/any(c:c eq 'Unified')"
        if ($SearchQuery) {
            $filter = "mailEnabled eq true and NOT groupTypes/any(c:c eq 'Unified') and (startsWith(displayName,'$SearchQuery') or startsWith(mail,'$SearchQuery'))"
        }
        $groups = Get-MgGroup -Filter $filter -Top 50 `
            -Property "Id,DisplayName,Mail,Description,MailNickname,MailEnabled,SecurityEnabled,CreatedDateTime" `
            -ErrorAction Stop
        return @{ Success = $true; Groups = $groups }
    } catch {
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
        $body = @{
            DisplayName     = $DisplayName
            MailNickname    = $MailNickname
            MailEnabled     = $true
            SecurityEnabled = $SecurityEnabled
            GroupTypes      = @()
        }
        if ($Description) { $body.Description = $Description }
        $group = New-MgGroup -BodyParameter $body -ErrorAction Stop
        return @{ Success = $true; Group = $group }
    } catch {
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
        $body = @{}
        if ($DisplayName) { $body.DisplayName   = $DisplayName }
        if ($null -ne $Description) { $body.Description = $Description }
        Update-MgGroup -GroupId $GroupId -BodyParameter $body -ErrorAction Stop
        return @{ Success = $true }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Add-DGMember {
    param([string]$GroupId, [string]$UserId)
    try {
        Import-GraphModules
        $ref = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" }
        New-MgGroupMemberByRef -GroupId $GroupId -BodyParameter $ref -ErrorAction Stop
        return @{ Success = $true }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Remove-DGMember {
    param([string]$GroupId, [string]$MemberId)
    try {
        Import-GraphModules
        Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $MemberId -ErrorAction Stop
        return @{ Success = $true }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-DGProperties {
    param([string]$GroupId)
    try {
        Import-GraphModules
        $group   = Get-MgGroup -GroupId $GroupId `
            -Property "Id,DisplayName,Mail,Description,MailNickname,MailEnabled,SecurityEnabled,GroupTypes,CreatedDateTime,Visibility" `
            -ErrorAction Stop
        $members = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop
        return @{ Success = $true; Group = $group; Members = $members }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Search-ExoManUsers {
    param([string]$Query)
    try {
        Import-GraphModules
        $filter = "startsWith(displayName,'$Query') or startsWith(userPrincipalName,'$Query')"
        $users  = Get-MgUser -Filter $filter -Top 30 `
            -Property "Id,DisplayName,UserPrincipalName,Mail" -ErrorAction Stop
        return @{ Success = $true; Users = $users }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}
