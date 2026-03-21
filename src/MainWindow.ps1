# MainWindow.ps1 - ExoMan v1.0 main application window

function Show-MainWindow {

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="ExoMan v1.0 – Exchange Online Management"
    Width="820" Height="700"
    MinWidth="720" MinHeight="600"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResizeWithGrip"
    Background="#F0F4F8">

  <Window.Resources>

    <!-- Primary action button -->
    <Style x:Key="PrimaryBtn" TargetType="Button">
      <Setter Property="Background"   Value="#0078D4"/>
      <Setter Property="Foreground"   Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"      Value="18,9"/>
      <Setter Property="FontSize"     Value="13"/>
      <Setter Property="Cursor"       Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="5"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#106EBE"/>
              </Trigger>
              <Trigger Property="IsPressed"   Value="True">
                <Setter Property="Background" Value="#005A9E"/>
              </Trigger>
              <Trigger Property="IsEnabled"   Value="False">
                <Setter Property="Background" Value="#B0BEC5"/>
                <Setter Property="Foreground"  Value="#ECEFF1"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Secondary / danger button -->
    <Style x:Key="SecondaryBtn" TargetType="Button" BasedOn="{StaticResource PrimaryBtn}">
      <Setter Property="Background" Value="#D13438"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#A4262C"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- Feature tile button -->
    <Style x:Key="TileBtn" TargetType="Button">
      <Setter Property="Background"   Value="White"/>
      <Setter Property="Foreground"   Value="#1A2D4A"/>
      <Setter Property="BorderBrush"  Value="#DDEAF7"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor"       Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="10" Padding="18,20">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background"  Value="#E5F1FB"/>
                <Setter Property="BorderBrush" Value="#0078D4"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="#CCE4F7"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Background"  Value="#F5F5F5"/>
                <Setter Property="BorderBrush" Value="#E0E0E0"/>
                <Setter Property="Foreground"  Value="#AAAAAA"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="72"/>   <!-- Header -->
      <RowDefinition Height="Auto"/> <!-- Status bar -->
      <RowDefinition Height="*" MinHeight="180"/>    <!-- Feature tiles -->
      <RowDefinition Height="5"/>    <!-- Splitter -->
      <RowDefinition Height="180" MinHeight="120"/>  <!-- Activity log -->
      <RowDefinition Height="30"/>   <!-- Footer -->
    </Grid.RowDefinitions>

    <!-- ═══════════════════ HEADER ═══════════════════ -->
    <Border Grid.Row="0" Background="#0F2B50">
      <Grid Margin="24,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="ExoMan" FontSize="30" FontWeight="Bold"
                     Foreground="White" VerticalAlignment="Center"/>
          <TextBlock Text=" v1.0" FontSize="16" Foreground="#AAC8E8"
                     VerticalAlignment="Bottom" Margin="0,0,0,4"/>
          <Rectangle Width="1" Fill="#3A5878" Margin="18,8,18,8"/>
          <TextBlock Text="Exchange Online Management Tool" FontSize="13"
                     Foreground="#AAC8E8" VerticalAlignment="Center"/>
        </StackPanel>
        <!-- Version badge -->
        <Border Grid.Column="1" Background="#1A3F6A" CornerRadius="4"
                Padding="10,4" VerticalAlignment="Center">
          <TextBlock Text="Post-Migration Admin" Foreground="#AAC8E8" FontSize="11"/>
        </Border>
      </Grid>
    </Border>

    <!-- ═══════════════════ CONNECTION STATUS ═══════════════════ -->
    <Border Grid.Row="1" Background="White" BorderBrush="#E0E8F0"
            BorderThickness="0,0,0,1" Padding="24,12">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <!-- Indicator + text -->
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse x:Name="StatusDot" Width="13" Height="13"
                   Fill="#D13438" VerticalAlignment="Center"/>
          <StackPanel Margin="12,0,0,0">
            <TextBlock x:Name="StatusTitle" Text="Not Connected"
                       FontSize="14" FontWeight="SemiBold" Foreground="#1A2D4A"/>
            <TextBlock x:Name="StatusSub" Text="Sign in to enable management features."
                       FontSize="12" Foreground="#666666"/>
          </StackPanel>
        </StackPanel>

        <!-- Disconnect button (hidden when disconnected) -->
        <Button x:Name="DisconnectBtn" Grid.Column="1"
                Content="Disconnect" Style="{StaticResource SecondaryBtn}"
                Margin="0,0,10,0" Visibility="Collapsed"/>

        <!-- Connect button -->
        <Button x:Name="ConnectBtn" Grid.Column="2"
                Content="🔑  Connect to Exchange Online"
                Style="{StaticResource PrimaryBtn}"/>
      </Grid>
    </Border>

    <!-- ═══════════════════ FEATURE TILES ═══════════════════ -->
    <Grid Grid.Row="2" Margin="24,24,24,20">
      <Grid.ColumnDefinitions>
        <ColumnDefinition/>
        <ColumnDefinition/>
        <ColumnDefinition/>
      </Grid.ColumnDefinitions>

      <!-- Distribution Groups -->
      <Button x:Name="BtnDG" Grid.Column="0" Style="{StaticResource TileBtn}"
              Margin="0,0,12,0" IsEnabled="False">
        <StackPanel HorizontalAlignment="Center">
          <TextBlock Text="📋" FontSize="40" HorizontalAlignment="Center" Margin="0,0,0,10"/>
          <TextBlock Text="Manage Distribution" FontSize="15" FontWeight="SemiBold"
                     HorizontalAlignment="Center"/>
          <TextBlock Text="Groups" FontSize="15" FontWeight="SemiBold"
                     HorizontalAlignment="Center"/>
          <TextBlock Text="Create, update, and manage&#10;members of DGs"
                     FontSize="11" Foreground="#666666" HorizontalAlignment="Center"
                     TextAlignment="Center" Margin="0,8,0,0"/>
        </StackPanel>
      </Button>

      <!-- Shared Mailbox -->
      <Button x:Name="BtnSharedMB" Grid.Column="1" Style="{StaticResource TileBtn}"
              Margin="6,0,6,0" IsEnabled="False">
        <StackPanel HorizontalAlignment="Center">
          <TextBlock Text="📬" FontSize="40" HorizontalAlignment="Center" Margin="0,0,0,10"/>
          <TextBlock Text="Manage Shared" FontSize="15" FontWeight="SemiBold"
                     HorizontalAlignment="Center"/>
          <TextBlock Text="Mailbox" FontSize="15" FontWeight="SemiBold"
                     HorizontalAlignment="Center"/>
          <TextBlock Text="Configure and manage&#10;shared mailboxes"
                     FontSize="11" Foreground="#666666" HorizontalAlignment="Center"
                     TextAlignment="Center" Margin="0,8,0,0"/>
        </StackPanel>
      </Button>

      <!-- User Mailbox -->
      <Button x:Name="BtnUserMB" Grid.Column="2" Style="{StaticResource TileBtn}"
              Margin="12,0,0,0" IsEnabled="False">
        <StackPanel HorizontalAlignment="Center">
          <TextBlock Text="👤" FontSize="40" HorizontalAlignment="Center" Margin="0,0,0,10"/>
          <TextBlock Text="Manage User" FontSize="15" FontWeight="SemiBold"
                     HorizontalAlignment="Center"/>
          <TextBlock Text="Mailbox" FontSize="15" FontWeight="SemiBold"
                     HorizontalAlignment="Center"/>
          <TextBlock Text="Manage individual user&#10;mailbox settings"
                     FontSize="11" Foreground="#666666" HorizontalAlignment="Center"
                     TextAlignment="Center" Margin="0,8,0,0"/>
        </StackPanel>
      </Button>
    </Grid>

    <!-- ═══════════════════ SPLITTER ═══════════════════ -->
    <GridSplitter Grid.Row="3" Height="5" HorizontalAlignment="Stretch"
                  Background="#2A4A7C" Cursor="SizeNS" VerticalAlignment="Center"/>

    <!-- ═══════════════════ ACTIVITY LOG ═══════════════════ -->
    <Border Grid.Row="4" Background="#070F1A" BorderBrush="#1A3050" BorderThickness="0,1,0,0">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="26"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Log header bar -->
        <Border Grid.Row="0" Background="#0C1E35">
          <Grid Margin="10,0">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <Ellipse Width="7" Height="7" Fill="#00C853" VerticalAlignment="Center" Margin="0,0,7,0"/>
              <TextBlock Text="Activity Log" Foreground="#7AABCC" FontSize="11"
                         FontWeight="SemiBold" VerticalAlignment="Center"/>
              <TextBlock x:Name="LogCount" Text="  (0 entries)" Foreground="#4A6A88"
                         FontSize="10" VerticalAlignment="Center"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right"
                        VerticalAlignment="Center">
              <CheckBox x:Name="LogAutoScroll" Content="Auto-scroll" IsChecked="True"
                        Foreground="#5A8AAA" FontSize="10" VerticalAlignment="Center"
                        Margin="0,0,12,0"/>
              <Button x:Name="BtnClearLog" Content="Clear"
                      Background="#1A3050" Foreground="#7AABCC"
                      BorderBrush="#2A4A7C" BorderThickness="1"
                      Padding="8,2" FontSize="10" Cursor="Hand"/>
            </StackPanel>
          </Grid>
        </Border>

        <!-- Log content (dark terminal style) -->
        <RichTextBox x:Name="LogBox" Grid.Row="1"
                     Background="#070F1A" BorderThickness="0"
                     IsReadOnly="True" IsDocumentEnabled="True"
                     FontFamily="Consolas,Courier New" FontSize="11.5"
                     Foreground="#C8D8E8"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Auto"
                     Padding="10,4,10,4"/>
      </Grid>
    </Border>

    <!-- ═══════════════════ FOOTER ═══════════════════ -->
    <Border Grid.Row="5" Background="#0C1E35">
      <TextBlock x:Name="FooterText"
                 Text="ExoMan v1.0  |  Ready"
                 Foreground="#4A7AAA" FontSize="10"
                 VerticalAlignment="Center" Margin="14,0"/>
    </Border>
  </Grid>
</Window>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # ── Get controls ──
    $statusDot     = $window.FindName("StatusDot")
    $statusTitle   = $window.FindName("StatusTitle")
    $statusSub     = $window.FindName("StatusSub")
    $connectBtn    = $window.FindName("ConnectBtn")
    $disconnectBtn = $window.FindName("DisconnectBtn")
    $btnDG         = $window.FindName("BtnDG")
    $btnSharedMB   = $window.FindName("BtnSharedMB")
    $btnUserMB     = $window.FindName("BtnUserMB")
    $footerText    = $window.FindName("FooterText")

    # ── Wire up the shared activity log ──
    $script:LogBox       = $window.FindName("LogBox")
    $logCount            = $window.FindName("LogCount")
    $logAutoScroll       = $window.FindName("LogAutoScroll")
    $btnClearLog         = $window.FindName("BtnClearLog")
    $script:LogEntryCount = 0

    # Remove default FlowDocument padding
    $script:LogBox.Document.PagePadding = [System.Windows.Thickness]::new(0)

    # Update entry count in header whenever Write-ExoLog runs
    $script:LogCountLabel = $logCount
    $script:LogAutoScroll = $logAutoScroll

    $btnClearLog.Add_Click({
        $script:LogBox.Document.Blocks.Clear()
        $script:LogEntryCount = 0
        $script:LogCountLabel.Text = "  (0 entries)"
    })

    # ── Helper: update UI to reflect connection state ──
    function Update-ConnectionUI {
        param([hashtable]$Status)
        if ($Status.Connected) {
            $statusDot.Fill       = [Windows.Media.Brushes]::Green
            $statusTitle.Text     = "Connected  ·  $($Status.Account)"
            $statusSub.Text       = "Tenant: $($Status.TenantId)"
            $connectBtn.Content   = "✔  Connected"
            $connectBtn.IsEnabled = $false
            $disconnectBtn.Visibility = "Visible"
            $btnDG.IsEnabled      = $true
            $btnSharedMB.IsEnabled = $true
            $btnUserMB.IsEnabled  = $true
            $footerText.Text      = "ExoMan v1.0  |  Signed in as $($Status.Account)"
        } else {
            $statusDot.Fill       = [Windows.Media.Brushes]::Crimson
            $statusTitle.Text     = "Not Connected"
            $statusSub.Text       = "Sign in to enable management features."
            $connectBtn.Content   = "🔑  Connect to Exchange Online"
            $connectBtn.IsEnabled = $true
            $disconnectBtn.Visibility = "Collapsed"
            $btnDG.IsEnabled      = $false
            $btnSharedMB.IsEnabled = $false
            $btnUserMB.IsEnabled  = $false
            $footerText.Text      = "ExoMan v1.0  |  Ready"
        }
    }

    # ── Check if already connected on launch ──
    $initialStatus = Get-ExoManConnectionStatus
    Update-ConnectionUI $initialStatus
    Write-ExoLog "ExoMan v1.0 started" "Action"
    if ($initialStatus.Connected) {
        Write-ExoLog "Already signed in as $($initialStatus.Account)" "Success"
    } else {
        Write-ExoLog "Not connected. Click 'Connect to Exchange Online' to sign in." "Info"
    }

    # ── Connect button ──
    $connectBtn.Add_Click({
        $connectBtn.Content   = "Connecting…"
        $connectBtn.IsEnabled = $false
        $footerText.Text      = "ExoMan v1.0  |  Opening Microsoft 365 login…"
        $window.Cursor        = [System.Windows.Input.Cursors]::Wait
        Write-ExoLog "Initiating Microsoft 365 browser login..." "Action"

        $result = Connect-ExoManGraph

        $window.Cursor = $null
        if ($result.Success) {
            Update-ConnectionUI @{ Connected = $true; Account = $result.Account; TenantId = $result.TenantId }
            Write-ExoLog "Connected successfully as $($result.Account)" "Success"
            Write-ExoLog "Tenant ID: $($result.TenantId)" "Info"
        } else {
            Update-ConnectionUI @{ Connected = $false }
            Write-ExoLog "Connection failed: $($result.Error)" "Error"
            [System.Windows.MessageBox]::Show(
                "Connection failed:`n`n$($result.Error)",
                "ExoMan – Connection Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            ) | Out-Null
            $footerText.Text = "ExoMan v1.0  |  Connection failed"
        }
    })

    # ── Disconnect button ──
    $disconnectBtn.Add_Click({
        Write-ExoLog "Disconnecting from Microsoft Graph..." "Action"
        Disconnect-ExoManGraph
        Update-ConnectionUI @{ Connected = $false }
        Write-ExoLog "Disconnected." "Info"
    })

    # ── Feature tile buttons ──
    $btnDG.Add_Click({
        Write-ExoLog "Opening Distribution Groups manager..." "Action"
        Show-DistributionGroupsWindow -Owner $window
    })

    $btnSharedMB.Add_Click({
        Write-ExoLog "Opening Shared Mailbox manager..." "Action"
        Show-SharedMailboxWindow -Owner $window
    })

    $btnUserMB.Add_Click({
        Write-ExoLog "Opening User Mailbox manager..." "Action"
        Show-UserMailboxWindow -Owner $window
    })

    $window.ShowDialog() | Out-Null
}
