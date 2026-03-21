# MainWindow.ps1 - ExoMan v1.0 main application window

function Show-MainWindow {

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="ExoMan v1.0 – Exchange Online Management"
    Width="780" Height="530"
    MinWidth="700" MinHeight="480"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanMinimize"
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
      <RowDefinition Height="*"/>    <!-- Feature tiles -->
      <RowDefinition Height="36"/>   <!-- Footer -->
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

    <!-- ═══════════════════ FOOTER ═══════════════════ -->
    <Border Grid.Row="3" Background="#1A2D4A">
      <TextBlock x:Name="FooterText"
                 Text="ExoMan v1.0  |  Ready"
                 Foreground="#AAC8E8" FontSize="11"
                 VerticalAlignment="Center" Margin="20,0"/>
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

    # ── Connect button ──
    $connectBtn.Add_Click({
        $connectBtn.Content   = "Connecting…"
        $connectBtn.IsEnabled = $false
        $footerText.Text      = "ExoMan v1.0  |  Opening Microsoft 365 login…"
        $window.Cursor        = [System.Windows.Input.Cursors]::Wait

        $result = Connect-ExoManGraph

        $window.Cursor = $null
        if ($result.Success) {
            Update-ConnectionUI @{ Connected = $true; Account = $result.Account; TenantId = $result.TenantId }
        } else {
            Update-ConnectionUI @{ Connected = $false }
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
        Disconnect-ExoManGraph
        Update-ConnectionUI @{ Connected = $false }
    })

    # ── Feature tile buttons ──
    $btnDG.Add_Click({
        Show-DistributionGroupsWindow -Owner $window
    })

    $btnSharedMB.Add_Click({
        Show-SharedMailboxWindow -Owner $window
    })

    $btnUserMB.Add_Click({
        Show-UserMailboxWindow -Owner $window
    })

    $window.ShowDialog() | Out-Null
}
