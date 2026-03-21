# GW-to-M365.ps1 - Google Workspace to Microsoft 365 migration scenario window

function Show-GWtoM365Window {
    param([System.Windows.Window]$Owner)

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Migraze - Google Workspace to Microsoft 365"
    Width="1120" Height="800"
    MinWidth="900" MinHeight="640"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResizeWithGrip"
    Background="#F0F4F8">

  <Window.Resources>
    <Style x:Key="ColBtn" TargetType="Button">
      <Setter Property="Height"          Value="40"/>
      <Setter Property="HorizontalAlignment" Value="Stretch"/>
      <Setter Property="Margin"          Value="0,0,0,6"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Foreground"      Value="White"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="5"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center" Margin="10,0"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Opacity" Value="0.85"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Opacity" Value="0.7"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="56"/>
      <RowDefinition Height="*" MinHeight="300"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="165" MinHeight="100"/>
      <RowDefinition Height="30"/>
    </Grid.RowDefinitions>

    <!-- HEADER -->
    <Border Grid.Row="0" Background="#0F2B50">
      <Grid Margin="16,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="BtnBack" Grid.Column="0"
                Content="&#x2190; Back" Padding="12,6"
                Background="#1A3F6A" Foreground="White"
                BorderThickness="0" Cursor="Hand"
                VerticalAlignment="Center" FontSize="12"/>
        <TextBlock Grid.Column="1"
                   Text="Google Workspace  &#x2192;  Microsoft 365"
                   FontSize="16" FontWeight="Bold" Foreground="White"
                   VerticalAlignment="Center" Margin="20,0,0,0"/>
      </Grid>
    </Border>

    <!-- TWO-COLUMN BODY -->
    <Grid Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="1"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- LEFT: Google Workspace -->
      <Grid Grid.Column="0">
        <Grid.RowDefinitions>
          <RowDefinition Height="46"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="1"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#2D6A4F">
          <Grid Margin="16,0">
            <TextBlock Text="Google Workspace" FontSize="14" FontWeight="Bold"
                       Foreground="White" VerticalAlignment="Center"/>
            <Border HorizontalAlignment="Right" Background="#1A4A30" CornerRadius="3"
                    Padding="8,2" VerticalAlignment="Center">
              <TextBlock Text="Source" Foreground="#90EEC0" FontSize="10"/>
            </Border>
          </Grid>
        </Border>

        <!-- Auth status -->
        <StackPanel Grid.Row="1" Margin="16,12,16,8">
          <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
            <Ellipse x:Name="GwStatusDot" Width="10" Height="10" Fill="#D13438"
                     VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBlock x:Name="GwStatusText" Text="Not Connected"
                       FontSize="12" FontWeight="SemiBold" Foreground="#333333"
                       VerticalAlignment="Center"/>
          </StackPanel>
          <TextBlock x:Name="GwAccountText" Text="" FontSize="11"
                     Foreground="#556677" Margin="0,0,0,8"/>
          <StackPanel Orientation="Horizontal">
            <Button x:Name="BtnGwLogin" Content="Connect Google Workspace"
                    Background="#2D6A4F" Foreground="White"
                    BorderThickness="0" Padding="14,7"
                    FontSize="12" Cursor="Hand" Margin="0,0,8,0"/>
            <Button x:Name="BtnGwDisconnect" Content="Disconnect"
                    Background="#D13438" Foreground="White"
                    BorderThickness="0" Padding="14,7"
                    FontSize="12" Cursor="Hand" Visibility="Collapsed"/>
          </StackPanel>
        </StackPanel>

        <Rectangle Grid.Row="2" Fill="#DDEAF7"/>

        <!-- Discovery buttons -->
        <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" Margin="16,12,16,12">
          <StackPanel>
            <TextBlock Text="DISCOVERY" FontSize="10" FontWeight="Bold"
                       Foreground="#889AAA" Margin="0,0,0,10"/>
            <Button x:Name="BtnDiscoverUsers"    Content="&#x1F465;  Discover Users"
                    Background="#2D6A4F" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnDiscoverGroups"   Content="&#x1F4CB;  Discover Distribution Groups"
                    Background="#2D6A4F" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnDiscoverCollab"   Content="&#x1F4EC;  Discover Collaboration Mailboxes"
                    Background="#2D6A4F" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnDiscoverDomains"  Content="&#x1F310;  Discover Domains"
                    Background="#2D6A4F" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnDiscoverOUs"      Content="&#x1F3E2;  Discover Org Units"
                    Background="#2D6A4F" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnDiscoverDrives"   Content="&#x1F4BE;  Discover Shared Drives"
                    Background="#2D6A4F" Style="{StaticResource ColBtn}"/>
            <TextBlock Text="All discovery results can be exported to CSV."
                       FontSize="10" Foreground="#889AAA" TextWrapping="Wrap" Margin="0,8,0,0"/>
          </StackPanel>
        </ScrollViewer>
      </Grid>

      <!-- DIVIDER -->
      <Rectangle Grid.Column="1" Fill="#DDEAF7"/>

      <!-- RIGHT: Microsoft 365 -->
      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="46"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="1"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#0078D4">
          <Grid Margin="16,0">
            <TextBlock Text="Microsoft 365" FontSize="14" FontWeight="Bold"
                       Foreground="White" VerticalAlignment="Center"/>
            <Border HorizontalAlignment="Right" Background="#005A9E" CornerRadius="3"
                    Padding="8,2" VerticalAlignment="Center">
              <TextBlock Text="Target" Foreground="#A8D4FF" FontSize="10"/>
            </Border>
          </Grid>
        </Border>

        <!-- M365 auth status -->
        <StackPanel Grid.Row="1" Margin="16,12,16,8">
          <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
            <Ellipse x:Name="M365StatusDot" Width="10" Height="10" Fill="#D13438"
                     VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBlock x:Name="M365StatusText" Text="Not Connected"
                       FontSize="12" FontWeight="SemiBold" Foreground="#333333"
                       VerticalAlignment="Center"/>
          </StackPanel>
          <TextBlock x:Name="M365AccountText" Text="" FontSize="11"
                     Foreground="#556677" Margin="0,0,0,8"/>
          <StackPanel Orientation="Horizontal">
            <Button x:Name="BtnM365Login" Content="Connect Microsoft 365"
                    Background="#0078D4" Foreground="White"
                    BorderThickness="0" Padding="14,7"
                    FontSize="12" Cursor="Hand" Margin="0,0,8,0"/>
            <Button x:Name="BtnM365Disconnect" Content="Disconnect"
                    Background="#D13438" Foreground="White"
                    BorderThickness="0" Padding="14,7"
                    FontSize="12" Cursor="Hand" Visibility="Collapsed"/>
          </StackPanel>
        </StackPanel>

        <Rectangle Grid.Row="2" Fill="#DDEAF7"/>

        <!-- Create / Manage buttons -->
        <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" Margin="16,12,16,12">
          <StackPanel>
            <TextBlock Text="CREATE / MIGRATE" FontSize="10" FontWeight="Bold"
                       Foreground="#889AAA" Margin="0,0,0,10"/>
            <Button x:Name="BtnCreateUsers"    Content="&#x1F465;  Create Users from Google"
                    Background="#0078D4" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnCreateGroups"   Content="&#x1F4CB;  Create Distribution Groups"
                    Background="#0078D4" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnCreateSharedMB" Content="&#x1F4EC;  Create Shared Mailboxes"
                    Background="#0078D4" Style="{StaticResource ColBtn}"/>
            <TextBlock Text="MANAGEMENT" FontSize="10" FontWeight="Bold"
                       Foreground="#889AAA" Margin="0,12,0,10"/>
            <Button x:Name="BtnManageDG"       Content="&#x2699;  Manage Distribution Groups"
                    Background="#106EBE" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnManageSharedMB" Content="&#x2699;  Manage Shared Mailboxes"
                    Background="#106EBE" Style="{StaticResource ColBtn}"/>
          </StackPanel>
        </ScrollViewer>
      </Grid>
    </Grid>

    <!-- SPLITTER -->
    <GridSplitter Grid.Row="2" Height="5" HorizontalAlignment="Stretch"
                  Background="#2A4A7C" Cursor="SizeNS" VerticalAlignment="Center"/>

    <!-- ACTIVITY LOG -->
    <Border Grid.Row="3" Background="#070F1A" BorderBrush="#1A3050" BorderThickness="0,1,0,0">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="26"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border Grid.Row="0" Background="#0C1E35">
          <Grid Margin="10,0">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <Ellipse Width="7" Height="7" Fill="#00C853" VerticalAlignment="Center" Margin="0,0,7,0"/>
              <TextBlock Text="Activity Log" Foreground="#7AABCC" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
              <TextBlock x:Name="LogCount" Text="  (0 entries)" Foreground="#4A6A88" FontSize="10" VerticalAlignment="Center"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
              <CheckBox x:Name="LogAutoScroll" Content="Auto-scroll" IsChecked="True"
                        Foreground="#5A8AAA" FontSize="10" VerticalAlignment="Center" Margin="0,0,12,0"/>
              <Button x:Name="BtnClearLog" Content="Clear"
                      Background="#1A3050" Foreground="#7AABCC"
                      BorderBrush="#2A4A7C" BorderThickness="1"
                      Padding="8,2" FontSize="10" Cursor="Hand"/>
            </StackPanel>
          </Grid>
        </Border>
        <RichTextBox x:Name="LogBox" Grid.Row="1"
                     Background="#070F1A" BorderThickness="0"
                     IsReadOnly="True" IsDocumentEnabled="True"
                     FontFamily="Consolas,Courier New" FontSize="11.5"
                     Foreground="#C8D8E8"
                     VerticalScrollBarVisibility="Auto"
                     Padding="10,4,10,4"/>
      </Grid>
    </Border>

    <!-- FOOTER -->
    <Border Grid.Row="4" Background="#0C1E35">
      <TextBlock Text="Migraze v2.0  |  Google Workspace to Microsoft 365"
                 Foreground="#4A7AAA" FontSize="10"
                 VerticalAlignment="Center" Margin="14,0"/>
    </Border>
  </Grid>
</Window>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    if ($Owner) { $window.Owner = $Owner }

    # ── Save/restore log box ──
    $savedLogBox        = $script:LogBox
    $savedLogCount      = $script:LogCountLabel
    $savedLogScroll     = $script:LogAutoScroll
    $savedLogEntryCount = $script:LogEntryCount

    $script:LogBox        = $window.FindName("LogBox")
    $script:LogCountLabel = $window.FindName("LogCount")
    $script:LogAutoScroll = $window.FindName("LogAutoScroll")
    $script:LogEntryCount = 0
    $script:LogBox.Document.PagePadding = [System.Windows.Thickness]::new(0)

    $window.Add_Closed({
        $script:LogBox        = $savedLogBox
        $script:LogCountLabel = $savedLogCount
        $script:LogAutoScroll = $savedLogScroll
        $script:LogEntryCount = $savedLogEntryCount
    })

    $window.FindName("BtnClearLog").Add_Click({
        $script:LogBox.Document.Blocks.Clear()
        $script:LogEntryCount = 0
        $script:LogCountLabel.Text = "  (0 entries)"
    })

    # ── Helper: update Google status UI ──
    function Update-GwUI {
        $s = Get-GoogleConnectionStatus
        if ($s.Connected) {
            $window.FindName("GwStatusDot").Fill      = [Windows.Media.Brushes]::Green
            $window.FindName("GwStatusText").Text     = "Connected"
            $window.FindName("GwAccountText").Text    = "$($s.Account)  ($($s.Method))"
            $window.FindName("BtnGwLogin").Visibility = "Collapsed"
            $window.FindName("BtnGwDisconnect").Visibility = "Visible"
        } else {
            $window.FindName("GwStatusDot").Fill      = [Windows.Media.Brushes]::Crimson
            $window.FindName("GwStatusText").Text     = "Not Connected"
            $window.FindName("GwAccountText").Text    = ""
            $window.FindName("BtnGwLogin").Visibility = "Visible"
            $window.FindName("BtnGwDisconnect").Visibility = "Collapsed"
        }
    }

    # ── Helper: update M365 status UI ──
    function Update-M365UI {
        $s = Get-MigrazeConnectionStatus
        if ($s.Connected) {
            $window.FindName("M365StatusDot").Fill      = [Windows.Media.Brushes]::Green
            $window.FindName("M365StatusText").Text     = "Connected"
            $window.FindName("M365AccountText").Text    = $s.Account
            $window.FindName("BtnM365Login").Visibility = "Collapsed"
            $window.FindName("BtnM365Disconnect").Visibility = "Visible"
        } else {
            $window.FindName("M365StatusDot").Fill      = [Windows.Media.Brushes]::Crimson
            $window.FindName("M365StatusText").Text     = "Not Connected"
            $window.FindName("M365AccountText").Text    = ""
            $window.FindName("BtnM365Login").Visibility = "Visible"
            $window.FindName("BtnM365Disconnect").Visibility = "Collapsed"
        }
    }

    # ── Back button ──
    $window.FindName("BtnBack").Add_Click({ $window.Close() })

    # ── Google auth buttons ──
    $window.FindName("BtnGwLogin").Add_Click({
        Write-MigrazeLog "Initiating Google Workspace authentication..." "Action"
        $ok = Show-GoogleLoginDialog
        Update-GwUI
    })
    $window.FindName("BtnGwDisconnect").Add_Click({
        Disconnect-Google; Update-GwUI
    })

    # ── M365 auth buttons ──
    $window.FindName("BtnM365Login").Add_Click({
        Write-MigrazeLog "Initiating Microsoft 365 authentication..." "Action"
        $result = Connect-MigrazeGraph
        Update-M365UI
    })
    $window.FindName("BtnM365Disconnect").Add_Click({
        Disconnect-MigrazeGraph; Update-M365UI
    })

    # ── Google Discovery buttons ──
    $window.FindName("BtnDiscoverUsers").Add_Click({
        Get-GoogleUsers -ExportCSV
        Update-GwUI
    })
    $window.FindName("BtnDiscoverGroups").Add_Click({
        Get-GoogleGroups -ExportCSV
        Update-GwUI
    })
    $window.FindName("BtnDiscoverCollab").Add_Click({
        Get-GoogleCollabMailboxes -ExportCSV
        Update-GwUI
    })
    $window.FindName("BtnDiscoverDomains").Add_Click({
        Get-GoogleDomains -ExportCSV
        Update-GwUI
    })
    $window.FindName("BtnDiscoverOUs").Add_Click({
        Get-GoogleOrgUnits -ExportCSV
        Update-GwUI
    })
    $window.FindName("BtnDiscoverDrives").Add_Click({
        Get-GoogleSharedDrives -ExportCSV
        Update-GwUI
    })

    # ── M365 action buttons ──
    $comingSoon = {
        param($feature)
        Write-MigrazeLog "$feature - Full migration wizard coming in a future release." "Info"
        [System.Windows.MessageBox]::Show(
            "$feature`n`nThis feature will be available in a future release of Migraze.`nFor now, please use the discovery tools on the left to export data to CSV.",
            "Migraze - Coming Soon", "OK", "Information") | Out-Null
    }
    $window.FindName("BtnCreateUsers").Add_Click({
        & $comingSoon "Create Users from Google Workspace"
    })
    $window.FindName("BtnCreateGroups").Add_Click({
        & $comingSoon "Create Distribution Groups from Google"
    })
    $window.FindName("BtnCreateSharedMB").Add_Click({
        & $comingSoon "Create Shared Mailboxes from Google Groups"
    })
    $window.FindName("BtnManageDG").Add_Click({
        Write-MigrazeLog "Opening Distribution Groups manager..." "Action"
        Show-DistributionGroupsWindow -Owner $window
    })
    $window.FindName("BtnManageSharedMB").Add_Click({
        Write-MigrazeLog "Opening Shared Mailbox manager..." "Action"
        Show-SharedMailboxWindow -Owner $window
    })

    # ── Initial status check ──
    Update-GwUI
    Update-M365UI
    Write-MigrazeLog "Google Workspace to M365 scenario opened." "Action"
    Write-MigrazeLog "Connect to Google Workspace (left) and Microsoft 365 (right) to begin." "Info"

    $window.ShowDialog() | Out-Null
}