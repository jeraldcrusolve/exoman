# M365-to-M365.ps1 - M365 Tenant to Tenant Migration scenario window

function Show-M365toM365Window {
    param([System.Windows.Window]$Owner)

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Migraze - M365 Tenant to Tenant Migration"
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
    <Border Grid.Row="0" Background="#1A2D4A">
      <Grid Margin="16,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="BtnBack" Grid.Column="0"
                Content="&#x2190; Back" Padding="12,6"
                Background="#2A3F5A" Foreground="White"
                BorderThickness="0" Cursor="Hand"
                VerticalAlignment="Center" FontSize="12"/>
        <TextBlock Grid.Column="1"
                   Text="M365 Tenant to Tenant Migration"
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

      <!-- LEFT: Source M365 Tenant -->
      <Grid Grid.Column="0">
        <Grid.RowDefinitions>
          <RowDefinition Height="46"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="1"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#37474F">
          <Grid Margin="16,0">
            <TextBlock Text="Source M365 Tenant" FontSize="14" FontWeight="Bold"
                       Foreground="White" VerticalAlignment="Center"/>
            <Border HorizontalAlignment="Right" Background="#263238" CornerRadius="3"
                    Padding="8,2" VerticalAlignment="Center">
              <TextBlock Text="Source" Foreground="#90CAF9" FontSize="10"/>
            </Border>
          </Grid>
        </Border>

        <!-- Source auth -->
        <StackPanel Grid.Row="1" Margin="16,12,16,8">
          <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
            <Ellipse x:Name="SrcStatusDot" Width="10" Height="10" Fill="#D13438"
                     VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBlock x:Name="SrcStatusText" Text="Not Connected"
                       FontSize="12" FontWeight="SemiBold" Foreground="#333333" VerticalAlignment="Center"/>
          </StackPanel>
          <TextBlock x:Name="SrcAccountText" Text="" FontSize="11" Foreground="#556677" Margin="0,0,0,8"/>
          <StackPanel Orientation="Horizontal">
            <Button x:Name="BtnSrcLogin" Content="Connect Source Tenant"
                    Background="#455A64" Foreground="White"
                    BorderThickness="0" Padding="14,7"
                    FontSize="12" Cursor="Hand" Margin="0,0,8,0"/>
            <Button x:Name="BtnSrcDisconnect" Content="Disconnect"
                    Background="#D13438" Foreground="White"
                    BorderThickness="0" Padding="14,7"
                    FontSize="12" Cursor="Hand" Visibility="Collapsed"/>
          </StackPanel>
        </StackPanel>

        <Rectangle Grid.Row="2" Fill="#DDEAF7"/>

        <!-- Source Discovery buttons -->
        <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" Margin="16,12,16,12">
          <StackPanel>
            <TextBlock Text="DISCOVERY" FontSize="10" FontWeight="Bold"
                       Foreground="#889AAA" Margin="0,0,0,10"/>
            <Button x:Name="BtnDiscSrcUsers"      Content="&#x1F465;  Discover Users"
                    Background="#455A64" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnDiscSrcSecGroups"  Content="&#x1F512;  Discover Security Groups"
                    Background="#455A64" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnDiscSrcDGs"        Content="&#x1F4CB;  Discover Distribution Groups"
                    Background="#455A64" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnDiscSrcSharedMBs"  Content="&#x1F4EC;  Discover Shared Mailboxes"
                    Background="#455A64" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnDiscSrcUserMBs"    Content="&#x1F4E5;  Discover User Mailboxes"
                    Background="#455A64" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnDiscSrcRooms"      Content="&#x1F3E2;  Discover Rooms and Resources"
                    Background="#455A64" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnDiscSrcContacts"   Content="&#x1F4C7;  Discover Contacts"
                    Background="#455A64" Style="{StaticResource ColBtn}"/>
            <TextBlock Text="All discovery results can be exported to CSV."
                       FontSize="10" Foreground="#889AAA" TextWrapping="Wrap" Margin="0,8,0,0"/>
          </StackPanel>
        </ScrollViewer>
      </Grid>

      <!-- DIVIDER -->
      <Rectangle Grid.Column="1" Fill="#DDEAF7"/>

      <!-- RIGHT: Target M365 Tenant -->
      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="46"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="1"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#0078D4">
          <Grid Margin="16,0">
            <TextBlock Text="Target M365 Tenant" FontSize="14" FontWeight="Bold"
                       Foreground="White" VerticalAlignment="Center"/>
            <Border HorizontalAlignment="Right" Background="#005A9E" CornerRadius="3"
                    Padding="8,2" VerticalAlignment="Center">
              <TextBlock Text="Target" Foreground="#A8D4FF" FontSize="10"/>
            </Border>
          </Grid>
        </Border>

        <!-- Target auth -->
        <StackPanel Grid.Row="1" Margin="16,12,16,8">
          <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
            <Ellipse x:Name="TgtStatusDot" Width="10" Height="10" Fill="#D13438"
                     VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBlock x:Name="TgtStatusText" Text="Not Connected"
                       FontSize="12" FontWeight="SemiBold" Foreground="#333333" VerticalAlignment="Center"/>
          </StackPanel>
          <TextBlock x:Name="TgtAccountText" Text="" FontSize="11" Foreground="#556677" Margin="0,0,0,8"/>
          <StackPanel Orientation="Horizontal">
            <Button x:Name="BtnTgtLogin" Content="Connect Target Tenant"
                    Background="#0078D4" Foreground="White"
                    BorderThickness="0" Padding="14,7"
                    FontSize="12" Cursor="Hand" Margin="0,0,8,0"/>
            <Button x:Name="BtnTgtDisconnect" Content="Disconnect"
                    Background="#D13438" Foreground="White"
                    BorderThickness="0" Padding="14,7"
                    FontSize="12" Cursor="Hand" Visibility="Collapsed"/>
          </StackPanel>
        </StackPanel>

        <Rectangle Grid.Row="2" Fill="#DDEAF7"/>

        <!-- Target Create/Manage buttons -->
        <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" Margin="16,12,16,12">
          <StackPanel>
            <TextBlock Text="CREATE / MIGRATE" FontSize="10" FontWeight="Bold"
                       Foreground="#889AAA" Margin="0,0,0,10"/>
            <Button x:Name="BtnCreateTgtUsers"     Content="&#x1F465;  Create Users"
                    Background="#0078D4" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnCreateTgtSecGroups" Content="&#x1F512;  Create Security Groups"
                    Background="#0078D4" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnCreateTgtDGs"       Content="&#x1F4CB;  Create Distribution Groups"
                    Background="#0078D4" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnCreateTgtSharedMBs" Content="&#x1F4EC;  Create Shared Mailboxes"
                    Background="#0078D4" Style="{StaticResource ColBtn}"/>
            <TextBlock Text="MANAGEMENT" FontSize="10" FontWeight="Bold"
                       Foreground="#889AAA" Margin="0,12,0,10"/>
            <Button x:Name="BtnManageTgtDG"        Content="&#x2699;  Manage Distribution Groups"
                    Background="#106EBE" Style="{StaticResource ColBtn}"/>
            <Button x:Name="BtnManageTgtSharedMB"  Content="&#x2699;  Manage Shared Mailboxes"
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
      <TextBlock Text="Migraze v2.0  |  M365 Tenant to Tenant Migration"
                 Foreground="#4A7AAA" FontSize="10"
                 VerticalAlignment="Center" Margin="14,0"/>
    </Border>
  </Grid>
</Window>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    if ($Owner) { $window.Owner = $Owner }

    # Save/restore log box
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

    # ── Source status helper ──
    function Update-SrcUI {
        $s = Get-MigrazeConnectionStatus
        if ($s.Connected) {
            $window.FindName("SrcStatusDot").Fill      = [Windows.Media.Brushes]::Green
            $window.FindName("SrcStatusText").Text     = "Connected"
            $window.FindName("SrcAccountText").Text    = $s.Account
            $window.FindName("BtnSrcLogin").Visibility = "Collapsed"
            $window.FindName("BtnSrcDisconnect").Visibility = "Visible"
        } else {
            $window.FindName("SrcStatusDot").Fill      = [Windows.Media.Brushes]::Crimson
            $window.FindName("SrcStatusText").Text     = "Not Connected"
            $window.FindName("SrcAccountText").Text    = ""
            $window.FindName("BtnSrcLogin").Visibility = "Visible"
            $window.FindName("BtnSrcDisconnect").Visibility = "Collapsed"
        }
    }

    # ── Target status helper ──
    function Update-TgtUI {
        if ($script:IsTargetGraphConnected) {
            $window.FindName("TgtStatusDot").Fill      = [Windows.Media.Brushes]::Green
            $window.FindName("TgtStatusText").Text     = "Connected"
            $window.FindName("TgtAccountText").Text    = $script:TargetGraphAccount
            $window.FindName("BtnTgtLogin").Visibility = "Collapsed"
            $window.FindName("BtnTgtDisconnect").Visibility = "Visible"
        } else {
            $window.FindName("TgtStatusDot").Fill      = [Windows.Media.Brushes]::Crimson
            $window.FindName("TgtStatusText").Text     = "Not Connected"
            $window.FindName("TgtAccountText").Text    = ""
            $window.FindName("BtnTgtLogin").Visibility = "Visible"
            $window.FindName("BtnTgtDisconnect").Visibility = "Collapsed"
        }
    }

    $window.FindName("BtnBack").Add_Click({ $window.Close() })

    # Auth
    $window.FindName("BtnSrcLogin").Add_Click({
        Write-MigrazeLog "Connecting to SOURCE M365 tenant..." "Action"
        $result = Connect-MigrazeGraph; Update-SrcUI
    })
    $window.FindName("BtnSrcDisconnect").Add_Click({ Disconnect-MigrazeGraph; Update-SrcUI })
    $window.FindName("BtnTgtLogin").Add_Click({
        Write-MigrazeLog "Connecting to TARGET M365 tenant..." "Action"
        $result = Connect-MigrazeTargetGraph; Update-TgtUI
    })
    $window.FindName("BtnTgtDisconnect").Add_Click({ Disconnect-MigrazeTargetGraph; Update-TgtUI })

    # Source discovery - auth check wrapper
    $requireSrcAuth = {
        param([scriptblock]$Action)
        $s = Get-MigrazeConnectionStatus
        if (-not $s.Connected) {
            Write-MigrazeLog "Source M365 authentication required. Click Connect Source Tenant first." "Warning"
            [System.Windows.MessageBox]::Show(
                "Please connect to the Source M365 Tenant first.",
                "Migraze - Authentication Required", "OK", "Warning") | Out-Null
            return
        }
        & $Action
    }

    $window.FindName("BtnDiscSrcUsers").Add_Click({
        & $requireSrcAuth { Get-M365Users -ExportCSV }
    })
    $window.FindName("BtnDiscSrcSecGroups").Add_Click({
        & $requireSrcAuth { Get-M365SecurityGroups -ExportCSV }
    })
    $window.FindName("BtnDiscSrcDGs").Add_Click({
        & $requireSrcAuth { Get-M365DistributionGroupsSrc -ExportCSV }
    })
    $window.FindName("BtnDiscSrcSharedMBs").Add_Click({
        & $requireSrcAuth { Get-M365SharedMailboxesSrc -ExportCSV }
    })
    $window.FindName("BtnDiscSrcUserMBs").Add_Click({
        & $requireSrcAuth { Get-M365UserMailboxesSrc -ExportCSV }
    })
    $window.FindName("BtnDiscSrcRooms").Add_Click({
        & $requireSrcAuth { Get-M365RoomsResources -ExportCSV }
    })
    $window.FindName("BtnDiscSrcContacts").Add_Click({
        & $requireSrcAuth { Get-M365Contacts -ExportCSV }
    })

    # Target create buttons - coming soon
    $comingSoon = {
        param($feature)
        Write-MigrazeLog "$feature - Coming in a future release." "Info"
        [System.Windows.MessageBox]::Show(
            "$feature`n`nThis feature will be available in a future release of Migraze.",
            "Migraze - Coming Soon", "OK", "Information") | Out-Null
    }
    $window.FindName("BtnCreateTgtUsers").Add_Click({     & $comingSoon "Create Users in Target Tenant" })
    $window.FindName("BtnCreateTgtSecGroups").Add_Click({ & $comingSoon "Create Security Groups in Target Tenant" })
    $window.FindName("BtnCreateTgtDGs").Add_Click({       & $comingSoon "Create Distribution Groups in Target Tenant" })
    $window.FindName("BtnCreateTgtSharedMBs").Add_Click({ & $comingSoon "Create Shared Mailboxes in Target Tenant" })

    $window.FindName("BtnManageTgtDG").Add_Click({
        Write-MigrazeLog "Opening Distribution Groups manager..." "Action"
        Show-DistributionGroupsWindow -Owner $window
    })
    $window.FindName("BtnManageTgtSharedMB").Add_Click({
        Write-MigrazeLog "Opening Shared Mailbox manager..." "Action"
        Show-SharedMailboxWindow -Owner $window
    })

    Update-SrcUI
    Update-TgtUI
    Write-MigrazeLog "M365 Tenant to Tenant Migration scenario opened." "Action"
    Write-MigrazeLog "Connect to Source Tenant (left) and Target Tenant (right) to begin." "Info"

    $window.ShowDialog() | Out-Null
}