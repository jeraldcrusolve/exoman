# MainWindow.ps1 - Migraze v2.0 main application window (single-window navigation)

function Show-MainWindow {

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Migraze v2.0 - Cloud Management Platform"
    Width="860" Height="620"
    MinWidth="720" MinHeight="500"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResizeWithGrip"
    Background="#F0F4F8">

  <Window.Resources>
    <Style x:Key="FeatureCard" TargetType="Button">
      <Setter Property="Background"             Value="White"/>
      <Setter Property="BorderBrush"            Value="#DDEAF7"/>
      <Setter Property="BorderThickness"        Value="1"/>
      <Setter Property="Padding"                Value="0"/>
      <Setter Property="Cursor"                 Value="Hand"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="CB" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="10" Padding="22,20">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="CB" Property="BorderBrush"  Value="#0078D4"/>
                <Setter TargetName="CB" Property="Background"   Value="#F0F7FF"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="CB" Property="Background"   Value="#E3F0FF"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="CB" Property="Background"   Value="#F8F9FA"/>
                <Setter TargetName="CB" Property="BorderBrush"  Value="#E0E8F0"/>
                <Setter TargetName="CB" Property="Opacity"      Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="68"/>
      <RowDefinition Height="*" MinHeight="200"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="170" MinHeight="110"/>
      <RowDefinition Height="28"/>
    </Grid.RowDefinitions>

    <!-- ═══════════════════════════════ HEADER ═══════════════════════════════ -->
    <Border Grid.Row="0" Background="#0F2B50">
      <Grid Margin="20,0,24,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <!-- Back button (hidden on home, visible on sub-views) -->
        <Button x:Name="BtnBack" Grid.Column="0"
                Background="Transparent" BorderThickness="0"
                Foreground="#AAC8E8" FontSize="12" Cursor="Hand"
                Padding="8,0,14,0" VerticalAlignment="Stretch"
                Visibility="Collapsed">
          <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
            <TextBlock Text="&#x2190;" FontSize="16" VerticalAlignment="Center" Margin="0,0,5,0"/>
            <TextBlock Text="Back" VerticalAlignment="Center"/>
          </StackPanel>
        </Button>

        <!-- Title area -->
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Migraze" FontSize="26" FontWeight="Bold"
                     Foreground="White" VerticalAlignment="Center"/>
          <TextBlock Text=" v2.0" FontSize="13" Foreground="#AAC8E8"
                     VerticalAlignment="Bottom" Margin="2,0,0,5"/>
          <Rectangle Width="1" Fill="#3A5878" Margin="16,10,16,10"/>
          <TextBlock x:Name="HeaderSubtitle" Text="Cloud Management Platform"
                     FontSize="12" Foreground="#AAC8E8" VerticalAlignment="Center"/>
        </StackPanel>

        <!-- Right badge (connection status shown on M365 view) -->
        <Border Grid.Column="2" x:Name="ConnStatusBadge"
                Background="#2A1010" CornerRadius="4" Padding="12,5"
                VerticalAlignment="Center" Visibility="Collapsed">
          <StackPanel Orientation="Horizontal">
            <Ellipse x:Name="ConnDot" Width="8" Height="8" Fill="#FF5252"
                     VerticalAlignment="Center" Margin="0,0,7,0"/>
            <TextBlock x:Name="ConnStatusText" Text="Not Connected"
                       Foreground="#FFAAAA" FontSize="11" FontWeight="SemiBold"/>
          </StackPanel>
        </Border>
      </Grid>
    </Border>

    <!-- ═══════════════════════════════ CONTENT ══════════════════════════════ -->
    <Grid Grid.Row="1">

      <!-- ─────────────────── VIEW: HOME ─────────────────── -->
      <ScrollViewer x:Name="ViewHome" VerticalScrollBarVisibility="Auto" Visibility="Visible">
        <StackPanel Margin="36,32,36,20">
          <TextBlock Text="Select Your Environment" FontSize="22" FontWeight="Bold"
                     Foreground="#1A2D4A" Margin="0,0,0,6"/>
          <TextBlock Text="Choose the platform you want to manage."
                     FontSize="12" Foreground="#667788" Margin="0,0,0,28"/>

          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="24"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Google Workspace (Coming Soon) -->
            <Border Grid.Column="0" Background="#FAFAFA" BorderBrush="#E0E8F0"
                    BorderThickness="1" CornerRadius="10" Padding="28,26">
              <StackPanel>
                <TextBlock Text="&#x1F4E7;" FontSize="44" HorizontalAlignment="Center"
                           Margin="0,0,0,14" Opacity="0.4"/>
                <TextBlock Text="Google Workspace" FontSize="16" FontWeight="Bold"
                           Foreground="#9AAAB8" HorizontalAlignment="Center" Margin="0,0,0,10"/>
                <TextBlock TextWrapping="Wrap" FontSize="12" Foreground="#AABBCC"
                           HorizontalAlignment="Center" TextAlignment="Center"
                           Text="Manage users, groups, shared drives and Gmail in your Google Workspace environment."
                           Margin="0,0,0,20"/>
                <Border Background="#E8EFF5" CornerRadius="20" Padding="16,6"
                        HorizontalAlignment="Center">
                  <TextBlock Text="Coming Soon" Foreground="#8899AA"
                             FontSize="11" FontWeight="SemiBold"/>
                </Border>
              </StackPanel>
            </Border>

            <!-- Microsoft 365 (Active) -->
            <Button x:Name="BtnM365" Grid.Column="2" Cursor="Hand"
                    Background="White" BorderBrush="#DDEAF7" BorderThickness="1"
                    Padding="0" HorizontalContentAlignment="Stretch">
              <Button.Template>
                <ControlTemplate TargetType="Button">
                  <Border x:Name="CB" Background="{TemplateBinding Background}"
                          BorderBrush="{TemplateBinding BorderBrush}"
                          BorderThickness="{TemplateBinding BorderThickness}"
                          CornerRadius="10" Padding="28,26">
                    <ContentPresenter/>
                  </Border>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                      <Setter TargetName="CB" Property="BorderBrush" Value="#0078D4"/>
                      <Setter TargetName="CB" Property="Background"  Value="#F0F7FF"/>
                    </Trigger>
                    <Trigger Property="IsPressed" Value="True">
                      <Setter TargetName="CB" Property="Background"  Value="#E3F0FF"/>
                    </Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </Button.Template>
              <StackPanel>
                <TextBlock Text="&#x2601;" FontSize="44" Foreground="#0078D4"
                           HorizontalAlignment="Center" Margin="0,0,0,14"/>
                <TextBlock Text="Microsoft 365" FontSize="16" FontWeight="Bold"
                           Foreground="#1A2D4A" HorizontalAlignment="Center" Margin="0,0,0,10"/>
                <TextBlock TextWrapping="Wrap" FontSize="12" Foreground="#556677"
                           HorizontalAlignment="Center" TextAlignment="Center"
                           Text="Manage distribution groups, shared mailboxes and user mailboxes in your Microsoft 365 tenant."
                           Margin="0,0,0,20"/>
                <DockPanel LastChildFill="False">
                  <Border Background="#EBF3FD" CornerRadius="4" Padding="10,4" DockPanel.Dock="Left">
                    <TextBlock Text="Exchange Online" Foreground="#0078D4"
                               FontSize="10" FontWeight="SemiBold"/>
                  </Border>
                  <TextBlock Text="&#x2192;" FontSize="20" Foreground="#0078D4"
                             DockPanel.Dock="Right" VerticalAlignment="Center"/>
                </DockPanel>
              </StackPanel>
            </Button>
          </Grid>
        </StackPanel>
      </ScrollViewer>

      <!-- ─────────────────── VIEW: MICROSOFT 365 ─────────────────── -->
      <ScrollViewer x:Name="ViewM365" VerticalScrollBarVisibility="Auto" Visibility="Collapsed">
        <StackPanel Margin="36,24,36,20">

          <!-- Section heading -->
          <TextBlock Text="Microsoft 365 Management" FontSize="20" FontWeight="Bold"
                     Foreground="#1A2D4A" Margin="0,0,0,4"/>
          <TextBlock Text="Connect to your tenant, then select a feature to manage."
                     FontSize="12" Foreground="#667788" Margin="0,0,0,20"/>

          <!-- Connect strip -->
          <Border Background="White" BorderBrush="#DDEAF7" BorderThickness="1"
                  CornerRadius="8" Padding="18,14" Margin="0,0,0,24">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock x:Name="TenantInfoText" Grid.Column="0"
                         Text="Connect to your Microsoft 365 tenant to manage Exchange Online objects."
                         FontSize="12" Foreground="#445566" TextWrapping="Wrap"
                         VerticalAlignment="Center"/>
              <Button x:Name="BtnConnect" Grid.Column="1"
                      Cursor="Hand" Margin="16,0,0,0" Padding="18,10"
                      FontSize="13" FontWeight="SemiBold"
                      Background="#0078D4" Foreground="White" BorderThickness="0">
                <Button.Template>
                  <ControlTemplate TargetType="Button">
                    <Border x:Name="BB" Background="{TemplateBinding Background}"
                            CornerRadius="6" Padding="{TemplateBinding Padding}">
                      <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsMouseOver" Value="True">
                        <Setter TargetName="BB" Property="Background" Value="#006CBF"/>
                      </Trigger>
                      <Trigger Property="IsPressed" Value="True">
                        <Setter TargetName="BB" Property="Background" Value="#005AA3"/>
                      </Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </Button.Template>
                <TextBlock x:Name="BtnConnectText" Text="Connect to M365 Tenant"/>
              </Button>
            </Grid>
          </Border>

          <!-- Feature heading -->
          <TextBlock Text="Manage Objects" FontSize="15" FontWeight="Bold"
                     Foreground="#1A2D4A" Margin="0,0,0,4"/>
          <TextBlock Text="Select a feature below to get started."
                     FontSize="12" Foreground="#667788" Margin="0,0,0,16"/>

          <!-- Feature tiles -->
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="20"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="20"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Button x:Name="BtnDG" Grid.Column="0" Style="{StaticResource FeatureCard}">
              <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="&#x1F465;" FontSize="34" HorizontalAlignment="Center" Margin="0,0,0,10"/>
                <TextBlock Text="Distribution Groups" FontSize="13" FontWeight="Bold"
                           Foreground="#1A2D4A" HorizontalAlignment="Center"
                           TextAlignment="Center" TextWrapping="Wrap" Margin="0,0,0,8"/>
                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#667788"
                           HorizontalAlignment="Center" TextAlignment="Center"
                           Text="Create, update and manage distribution group members"
                           Margin="0,0,0,12"/>
                <Border Background="#EBF3FD" CornerRadius="4" Padding="8,3" HorizontalAlignment="Center">
                  <TextBlock Text="Create / Edit / Members" Foreground="#0078D4"
                             FontSize="10" FontWeight="SemiBold"/>
                </Border>
              </StackPanel>
            </Button>

            <Button x:Name="BtnSM" Grid.Column="2" Style="{StaticResource FeatureCard}">
              <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="&#x1F4EC;" FontSize="34" HorizontalAlignment="Center" Margin="0,0,0,10"/>
                <TextBlock Text="Shared Mailbox" FontSize="13" FontWeight="Bold"
                           Foreground="#1A2D4A" HorizontalAlignment="Center"
                           TextAlignment="Center" TextWrapping="Wrap" Margin="0,0,0,8"/>
                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#667788"
                           HorizontalAlignment="Center" TextAlignment="Center"
                           Text="Create and manage shared mailboxes and permissions"
                           Margin="0,0,0,12"/>
                <Border Background="#EBF3FD" CornerRadius="4" Padding="8,3" HorizontalAlignment="Center">
                  <TextBlock Text="Create / Permissions" Foreground="#0078D4"
                             FontSize="10" FontWeight="SemiBold"/>
                </Border>
              </StackPanel>
            </Button>

            <Button x:Name="BtnUM" Grid.Column="4" Style="{StaticResource FeatureCard}">
              <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="&#x1F464;" FontSize="34" HorizontalAlignment="Center" Margin="0,0,0,10"/>
                <TextBlock Text="User Mailbox" FontSize="13" FontWeight="Bold"
                           Foreground="#1A2D4A" HorizontalAlignment="Center"
                           TextAlignment="Center" TextWrapping="Wrap" Margin="0,0,0,8"/>
                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#667788"
                           HorizontalAlignment="Center" TextAlignment="Center"
                           Text="Manage user mailbox settings, quotas and properties"
                           Margin="0,0,0,12"/>
                <Border Background="#EBF3FD" CornerRadius="4" Padding="8,3" HorizontalAlignment="Center">
                  <TextBlock Text="Settings / Properties" Foreground="#0078D4"
                             FontSize="10" FontWeight="SemiBold"/>
                </Border>
              </StackPanel>
            </Button>
          </Grid>

        </StackPanel>
      </ScrollViewer>

    </Grid><!-- end content Grid -->

    <!-- ═══════════════════════════════ SPLITTER ═════════════════════════════ -->
    <GridSplitter Grid.Row="2" Height="5" HorizontalAlignment="Stretch"
                  Background="#2A4A7C" Cursor="SizeNS" VerticalAlignment="Center"/>

    <!-- ═══════════════════════════════ ACTIVITY LOG ═════════════════════════ -->
    <Border Grid.Row="3" Background="#1E1E1E" BorderBrush="#1A3050" BorderThickness="0,1,0,0">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="26"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border Grid.Row="0" Background="#0C1E35">
          <Grid Margin="10,0">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <Ellipse Width="7" Height="7" Fill="#00C853" VerticalAlignment="Center" Margin="0,0,7,0"/>
              <TextBlock Text="Activity Log" Foreground="#7AABCC" FontSize="11"
                         FontWeight="SemiBold" VerticalAlignment="Center"/>
              <TextBlock x:Name="LogCount" Text="  (0 entries)" Foreground="#4A6A88"
                         FontSize="10" VerticalAlignment="Center"/>
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
                     Background="#1E1E1E" BorderThickness="0"
                     IsReadOnly="True" IsDocumentEnabled="True"
                     FontFamily="Consolas,Courier New" FontSize="11.5"
                     Foreground="#C8D8E8"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Auto"
                     Padding="10,4,10,4"/>
      </Grid>
    </Border>

    <!-- ═══════════════════════════════ FOOTER ═══════════════════════════════ -->
    <Border Grid.Row="4" Background="#0C1E35">
      <TextBlock x:Name="FooterText" Text="Migraze v2.0  |  Ready"
                 Foreground="#4A7AAA" FontSize="10"
                 VerticalAlignment="Center" Margin="14,0"/>
    </Border>
  </Grid>
</Window>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # ── Shared log wiring ──────────────────────────────────────────────────────
    $script:LogBox        = $window.FindName("LogBox")
    $script:LogCount      = $window.FindName("LogCount")
    $script:LogAutoScroll = $window.FindName("LogAutoScroll")
    $script:LogEntryCount = 0
    $script:LogBox.Document.PagePadding = [System.Windows.Thickness]::new(0)
    $script:LogCountLabel = $script:LogCount

    $window.FindName("BtnClearLog").Add_Click({
        $script:LogBox.Document.Blocks.Clear()
        $script:LogEntryCount = 0
        if ($script:LogCountLabel) { $script:LogCountLabel.Text = "  (0 entries)" }
    })

    # ── Named element references ───────────────────────────────────────────────
    $viewHome      = $window.FindName("ViewHome")
    $viewM365      = $window.FindName("ViewM365")
    $btnBack       = $window.FindName("BtnBack")
    $headerSub     = $window.FindName("HeaderSubtitle")
    $connBadge     = $window.FindName("ConnStatusBadge")
    $connDot       = $window.FindName("ConnDot")
    $connText      = $window.FindName("ConnStatusText")
    $tenantInfo    = $window.FindName("TenantInfoText")
    $btnConnect    = $window.FindName("BtnConnect")
    $btnConnText   = $window.FindName("BtnConnectText")
    $footerText    = $window.FindName("FooterText")
    $btnDG         = $window.FindName("BtnDG")
    $btnSM         = $window.FindName("BtnSM")
    $btnUM         = $window.FindName("BtnUM")

    # ── Navigation helpers ─────────────────────────────────────────────────────
    function Show-HomeView {
        $viewHome.Visibility  = "Visible"
        $viewM365.Visibility  = "Collapsed"
        $btnBack.Visibility   = "Collapsed"
        $connBadge.Visibility = "Collapsed"
        $headerSub.Text       = "Cloud Management Platform"
        $footerText.Text      = "Migraze v2.0  |  Ready"
    }

    function Show-M365View {
        $viewHome.Visibility  = "Collapsed"
        $viewM365.Visibility  = "Visible"
        $btnBack.Visibility   = "Visible"
        $connBadge.Visibility = "Visible"
        $headerSub.Text       = "Microsoft 365"
        Update-M365ConnStatus
        Write-MigrazeLog "Microsoft 365 management opened." "Info"
    }

    function Update-M365ConnStatus {
        $connected = Get-MigrazeConnectionStatus
        if ($connected) {
            $connDot.Fill         = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#00C853"))
            $connText.Text        = "Connected"
            $connText.Foreground  = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#AAFFCC"))
            $connBadge.Background = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0C2A10"))
            $btnConnText.Text     = "Disconnect"
            $btnConnect.Background= [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#C62828"))
            $infoMsg = if ($script:M365TenantDomain) { "Connected to: $($script:M365TenantDomain)" } else { "Connected to Microsoft 365." }
            $tenantInfo.Text      = $infoMsg
            $footerText.Text      = "Migraze v2.0  |  Microsoft 365  |  $($script:M365TenantDomain)"
            $btnDG.IsEnabled = $true; $btnSM.IsEnabled = $true; $btnUM.IsEnabled = $true
        } else {
            $connDot.Fill         = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#FF5252"))
            $connText.Text        = "Not Connected"
            $connText.Foreground  = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#FFAAAA"))
            $connBadge.Background = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#2A1010"))
            $btnConnText.Text     = "Connect to M365 Tenant"
            $btnConnect.Background= [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0078D4"))
            $tenantInfo.Text      = "Connect to your Microsoft 365 tenant to manage Exchange Online objects."
            $footerText.Text      = "Migraze v2.0  |  Microsoft 365  |  Not Connected"
            $btnDG.IsEnabled = $false; $btnSM.IsEnabled = $false; $btnUM.IsEnabled = $false
        }
    }

    # ── Event handlers ─────────────────────────────────────────────────────────
    $window.FindName("BtnM365").Add_Click({ Show-M365View })

    $btnBack.Add_Click({ Show-HomeView })

    $btnConnect.Add_Click({
        if (Get-MigrazeConnectionStatus) {
            Write-MigrazeLog "Disconnecting from Microsoft 365..." "Action"
            Disconnect-MigrazeGraph
        } else {
            Write-MigrazeLog "Opening browser for Microsoft 365 login..." "Action"
            $ok = Connect-MigrazeGraph
            if ($ok) { Write-MigrazeLog "Successfully connected to Microsoft 365." "Success" }
        }
        Update-M365ConnStatus
    })

    $btnDG.Add_Click({ Show-DistributionGroupsWindow -Owner $window })
    $btnSM.Add_Click({ Show-SharedMailboxWindow      -Owner $window })
    $btnUM.Add_Click({ Show-UserMailboxWindow        -Owner $window })

    # ── Start ──────────────────────────────────────────────────────────────────
    Write-MigrazeLog "Migraze v2.0 started." "Action"
    Write-MigrazeLog "Select an environment to get started." "Info"

    $window.ShowDialog() | Out-Null
}