# GoogleHelper.ps1 - Google Workspace OAuth2 authentication

$script:GoogleAccessToken = $null
$script:GoogleTokenExpiry  = $null
$script:GoogleAdminEmail   = $null
$script:GoogleDomain       = $null
$script:GoogleAuthMethod   = $null

function Get-GoogleConnectionStatus {
    if ($script:GoogleAccessToken -and $script:GoogleTokenExpiry -and (Get-Date) -lt $script:GoogleTokenExpiry) {
        return @{ Connected = $true; Account = $script:GoogleAdminEmail; Domain = $script:GoogleDomain; Method = $script:GoogleAuthMethod }
    }
    return @{ Connected = $false }
}

function Disconnect-Google {
    $script:GoogleAccessToken = $null
    $script:GoogleTokenExpiry  = $null
    $script:GoogleAdminEmail   = $null
    $script:GoogleDomain       = $null
    $script:GoogleAuthMethod   = $null
    Write-MigrazeLog "Disconnected from Google Workspace." "Info"
}

function Show-GoogleLoginDialog {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Google Workspace Authentication" Width="520" Height="400"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="#F0F4F8">
  <Grid Margin="30">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Text="Connect to Google Workspace" FontSize="18" FontWeight="Bold"
               Foreground="#1A2D4A" Margin="0,0,0,8"/>
    <TextBlock Grid.Row="1" TextWrapping="Wrap" Foreground="#555555" Margin="0,0,0,20"
               Text="Choose how to authenticate. Service Account is recommended for bulk admin operations."/>
    <StackPanel Grid.Row="2">
      <Border x:Name="CardSA" Background="White" BorderBrush="#DDEAF7" BorderThickness="1"
              CornerRadius="8" Margin="0,0,0,12" Padding="16" Cursor="Hand">
        <StackPanel>
          <TextBlock Text="Service Account JSON Key" FontSize="14" FontWeight="SemiBold" Foreground="#1A2D4A"/>
          <TextBlock Text="Browse for a service account .json key file from Google Cloud Console."
                     Foreground="#666666" FontSize="11" Margin="0,4,0,0" TextWrapping="Wrap"/>
          <TextBlock Text="Recommended for domain-wide admin access" Foreground="#2D6A4F" FontSize="10" Margin="0,6,0,0"/>
        </StackPanel>
      </Border>
      <Border x:Name="CardOAuth" Background="White" BorderBrush="#DDEAF7" BorderThickness="1"
              CornerRadius="8" Padding="16" Cursor="Hand">
        <StackPanel>
          <TextBlock Text="Interactive Browser Login (OAuth2)" FontSize="14" FontWeight="SemiBold" Foreground="#1A2D4A"/>
          <TextBlock Text="Sign in via browser with your Google admin account. Requires an OAuth2 Client ID."
                     Foreground="#666666" FontSize="11" Margin="0,4,0,0" TextWrapping="Wrap"/>
          <TextBlock Text="Requires OAuth2 Client ID from Google Cloud Console" Foreground="#0078D4" FontSize="10" Margin="0,6,0,0"/>
        </StackPanel>
      </Border>
    </StackPanel>
    <Button x:Name="BtnCancel" Grid.Row="3" Content="Cancel"
            HorizontalAlignment="Right" Padding="16,8" Margin="0,16,0,0"
            Background="#E0E0E0" Foreground="#333333" BorderThickness="0" Cursor="Hand"/>
  </Grid>
</Window>
"@
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $dialog = [Windows.Markup.XamlReader]::Load($reader)

    $dialog.FindName("CardSA").Add_MouseLeftButtonUp({
        $dialog.Tag = "ServiceAccount"; $dialog.Close()
    })
    $dialog.FindName("CardOAuth").Add_MouseLeftButtonUp({
        $dialog.Tag = "OAuth2"; $dialog.Close()
    })
    $dialog.FindName("BtnCancel").Add_Click({
        $dialog.Tag = "Cancel"; $dialog.Close()
    })

    $dialog.ShowDialog() | Out-Null

    switch ($dialog.Tag) {
        "ServiceAccount" { return Connect-GoogleServiceAccount }
        "OAuth2"         { return Connect-GoogleOAuth2 }
        default          { Write-MigrazeLog "Google auth cancelled." "Warning"; return $false }
    }
}

function Connect-GoogleServiceAccount {
    Add-Type -AssemblyName System.Windows.Forms
    $ofd = [System.Windows.Forms.OpenFileDialog]::new()
    $ofd.Title  = "Select Google Service Account JSON Key File"
    $ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-MigrazeLog "Google service account auth cancelled." "Warning"
        return $false
    }
    try {
        $saKey = Get-Content -LiteralPath $ofd.FileName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($saKey.type -ne "service_account") {
            throw "Not a service account key file (type: $($saKey.type))"
        }
        Write-MigrazeLog "Service account key loaded: $($saKey.client_email)" "Info"
        $token = Get-GoogleServiceAccountToken -ServiceAccountKey $saKey
        if ($token) {
            $script:GoogleAccessToken = $token
            $script:GoogleTokenExpiry  = (Get-Date).AddSeconds(3500)
            $script:GoogleAdminEmail   = $saKey.client_email
            $script:GoogleDomain       = ($saKey.client_email -split "@")[1]
            $script:GoogleAuthMethod   = "ServiceAccount"
            Write-MigrazeLog "Google Workspace connected via Service Account." "Success"
            Write-MigrazeLog "Account: $($saKey.client_email)" "Info"
            return $true
        }
        return $false
    } catch {
        Write-MigrazeLog "Service account auth failed: $($_.Exception.Message)" "Error"
        [System.Windows.MessageBox]::Show(
            "Google auth failed:`n`n$($_.Exception.Message)",
            "Migraze - Google Auth Error", "OK", "Error") | Out-Null
        return $false
    }
}

function Get-GoogleServiceAccountToken {
    param([object]$ServiceAccountKey)
    try {
        $now     = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $header  = '{"alg":"RS256","typ":"JWT"}'
        $payload = "{`"iss`":`"$($ServiceAccountKey.client_email)`",`"scope`":`"https://www.googleapis.com/auth/admin.directory.user.readonly https://www.googleapis.com/auth/admin.directory.group.readonly https://www.googleapis.com/auth/admin.directory.domain.readonly https://www.googleapis.com/auth/admin.directory.orgunit.readonly`",`"aud`":`"https://oauth2.googleapis.com/token`",`"exp`":$($now+3600),`"iat`":$now}"

        $toB64Url = { param($s) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($s)).TrimEnd("=").Replace("+","-").Replace("/","_") }
        $headerB64  = & $toB64Url $header
        $payloadB64 = & $toB64Url $payload
        $sigInput   = "$headerB64.$payloadB64"

        $pemKey   = $ServiceAccountKey.private_key -replace "-----BEGIN PRIVATE KEY-----","" -replace "-----END PRIVATE KEY-----","" -replace "`n","" -replace "`r","" -replace " ",""
        $keyBytes = [Convert]::FromBase64String($pemKey)
        $rsa      = [System.Security.Cryptography.RSA]::Create()
        $rsa.ImportPkcs8PrivateKey($keyBytes, [ref]$null)

        $sigBytes = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($sigInput),
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $sigB64   = [Convert]::ToBase64String($sigBytes).TrimEnd("=").Replace("+","-").Replace("/","_")
        $jwt      = "$sigInput.$sigB64"

        $body     = "grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=$jwt"
        $response = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method POST `
            -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
        Write-MigrazeLog "Google access token obtained (expires in ~1 hour)." "Info"
        return $response.access_token
    } catch {
        Write-MigrazeLog "JWT token request failed: $($_.Exception.Message)" "Error"
        throw
    }
}

function Connect-GoogleOAuth2 {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Google OAuth2 Setup" Width="480" Height="280"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize" Background="#F0F4F8">
  <Grid Margin="30">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Text="Google OAuth2 Authentication" FontSize="16" FontWeight="Bold"
               Foreground="#1A2D4A" Margin="0,0,0,8"/>
    <TextBlock Grid.Row="1" TextWrapping="Wrap" Foreground="#555555" Margin="0,0,0,16"
               Text="Enter your OAuth2 Client ID (Desktop app type) from Google Cloud Console. Create one at console.cloud.google.com."/>
    <TextBlock Grid.Row="2" Text="OAuth2 Client ID:" FontWeight="SemiBold"
               Foreground="#1A2D4A" Margin="0,0,0,6"/>
    <TextBox Grid.Row="3" x:Name="TxtClientId" Padding="8,6" FontSize="12"
             BorderBrush="#B0C8E0" BorderThickness="1"/>
    <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
      <Button x:Name="BtnOK" Content="Open Browser" Padding="16,8" Margin="0,0,10,0"
              Background="#0078D4" Foreground="White" BorderThickness="0" Cursor="Hand"/>
      <Button x:Name="BtnCancel" Content="Cancel" Padding="16,8"
              Background="#E0E0E0" Foreground="#333333" BorderThickness="0" Cursor="Hand"/>
    </StackPanel>
  </Grid>
</Window>
"@
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $d = [Windows.Markup.XamlReader]::Load($reader)
    $d.FindName("BtnOK").Add_Click({ $d.DialogResult = $true; $d.Close() })
    $d.FindName("BtnCancel").Add_Click({ $d.DialogResult = $false; $d.Close() })
    if (-not $d.ShowDialog()) {
        Write-MigrazeLog "Google OAuth2 auth cancelled." "Warning"
        return $false
    }
    $clientId = $d.FindName("TxtClientId").Text.Trim()
    if (-not $clientId) {
        Write-MigrazeLog "No Client ID provided. Auth cancelled." "Warning"
        return $false
    }
    try {
        $redirectUri = "http://localhost:8765/callback"
        $scopes      = "https://www.googleapis.com/auth/admin.directory.user.readonly https://www.googleapis.com/auth/admin.directory.group.readonly https://www.googleapis.com/auth/admin.directory.domain.readonly https://www.googleapis.com/auth/admin.directory.orgunit.readonly"
        $state       = [System.Guid]::NewGuid().ToString("N")
        $authUrl = "https://accounts.google.com/o/oauth2/v2/auth?client_id=$([Uri]::EscapeDataString($clientId))&redirect_uri=$([Uri]::EscapeDataString($redirectUri))&response_type=code&scope=$([Uri]::EscapeDataString($scopes))&state=$state&access_type=offline"

        Write-MigrazeLog "Opening browser for Google login..." "Action"
        Start-Process $authUrl

        $listener = [System.Net.HttpListener]::new()
        $listener.Prefixes.Add("http://localhost:8765/")
        $listener.Start()
        Write-MigrazeLog "Waiting for Google OAuth2 callback (120s timeout)..." "Info"

        $task    = $listener.GetContextAsync()
        $timeout = (Get-Date).AddSeconds(120)
        while (-not $task.IsCompleted -and (Get-Date) -lt $timeout) {
            Start-Sleep -Milliseconds 300
        }
        if (-not $task.IsCompleted) {
            $listener.Stop()
            throw "OAuth2 callback timed out after 120 seconds."
        }

        $ctx  = $task.Result
        Add-Type -AssemblyName System.Web
        $code = [System.Web.HttpUtility]::ParseQueryString($ctx.Request.Url.Query)["code"]

        $html = "<html><body style='font-family:Arial;text-align:center;margin-top:80px'><h2>Authenticated!</h2><p>You can close this tab and return to Migraze.</p></body></html>"
        $buf  = [Text.Encoding]::UTF8.GetBytes($html)
        $ctx.Response.ContentLength64 = $buf.Length
        $ctx.Response.OutputStream.Write($buf, 0, $buf.Length)
        $ctx.Response.Close()
        $listener.Stop()

        if (-not $code) { throw "No authorization code received from Google." }

        $body = @{ code=$code; client_id=$clientId; redirect_uri=$redirectUri; grant_type="authorization_code" }
        $resp = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method POST `
            -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop

        $script:GoogleAccessToken = $resp.access_token
        $script:GoogleTokenExpiry  = (Get-Date).AddSeconds($resp.expires_in - 60)
        $script:GoogleAuthMethod   = "OAuth2"

        try {
            $me = Invoke-RestMethod -Uri "https://www.googleapis.com/oauth2/v2/userinfo" `
                -Headers @{ Authorization = "Bearer $($script:GoogleAccessToken)" } -ErrorAction Stop
            $script:GoogleAdminEmail = $me.email
            $script:GoogleDomain     = ($me.email -split "@")[1]
        } catch {}

        Write-MigrazeLog "Google Workspace connected via OAuth2." "Success"
        Write-MigrazeLog "Signed in as: $($script:GoogleAdminEmail)" "Info"
        return $true
    } catch {
        Write-MigrazeLog "Google OAuth2 failed: $($_.Exception.Message)" "Error"
        [System.Windows.MessageBox]::Show(
            "Google OAuth2 failed:`n`n$($_.Exception.Message)",
            "Migraze - Google Auth Error", "OK", "Error") | Out-Null
        return $false
    }
}

function Invoke-WithGoogleAuth {
    param([scriptblock]$Action)
    $status = Get-GoogleConnectionStatus
    if (-not $status.Connected) {
        Write-MigrazeLog "Google authentication required..." "Info"
        $ok = Show-GoogleLoginDialog
        if (-not $ok) { return }
    }
    & $Action
}

function Invoke-GoogleAPI {
    param([string]$Url, [string]$Method = "GET", [hashtable]$Body = $null)
    if (-not $script:GoogleAccessToken) { throw "Not authenticated to Google Workspace." }
    $headers = @{ Authorization = "Bearer $script:GoogleAccessToken" }
    $params  = @{ Uri = $Url; Method = $Method; Headers = $headers; ErrorAction = "Stop" }
    if ($Body) { $params.Body = ($Body | ConvertTo-Json -Depth 10); $params.ContentType = "application/json" }
    return Invoke-RestMethod @params
}