# M365-Home.ps1 - Microsoft 365 management home window

function Show-M365HomeWindow {
    param([System.Windows.Window]$Owner)

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Migraze v2.0 - Microsoft 365"
    Width="860" Height="620"
    MinWidth="700" MinHeight="500"
    WindowStartupLocation="CenterOwner"
    ResizeMode="CanResizeWithGrip"
    Background="#F0F4F8">

  <Window.Resources>
    <Style x:Key="FeatureCard" TargetType="Button">
      <Setter Property="Background"   Value="White"/>
      <Setter Property="BorderBrush"  Value="#DDEAF7"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"      Value="0"/>
      <Setter Property="Cursor"       Value="Hand"/>
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
                <Setter TargetName="CB" Property="Opacity"      Value="0.55"/>
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

    <!-- HEADER -->
    <Border Grid.Row="0" Background="#0F2B50">
      <Grid Margin="16,0,24,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="BtnBack" Grid.Column="0"
                Background="Transparent" BorderThickness="0"
                Foreground="#AAC8E8" FontSize="13" Cursor="Hand"
                Padding="10,0,16,0" VerticalAlignment="Stretch"
                ToolTip="Back to environment selection">
          <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
            <TextBlock Text="&#x2190;" FontSize="16" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBlock Text="Back" FontSize="12" VerticalAlignment="Center"/>
          </StackPanel>
        </Button>
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="&#x2601;" FontSize="20" Foreground="#5BB3F0"
                     VerticalAlignment="Center" Margin="0,0,10,0"/>
          <TextBlock Text="Microsoft 365" FontSize="20" FontWeight="Bold"
                     Foreground="White" VerticalAlignment="Center"/>
          <Rectangle Width="1" Fill="#3A5878" Margin="16,10,16,10"/>
          <TextBlock Text="Exchange Online Management" FontSize="12"
                     Foreground="#AAC8E8" VerticalAlignment="Center"/>
        </StackPanel>
        <Border Grid.Column="2" x:Name="ConnStatusBadge"
                Background="#1B3A20" CornerRadius="4" Padding="12,5" VerticalAlignment="Center">
          <StackPanel Orientation="Horizontal">
            <Ellipse x:Name="ConnDot" Width="8" Height="8" Fill="#FF5252"
                     VerticalAlignment="Center" Margin="0,0,7,0"/>
            <TextBlock x:Name="ConnStatusText" Text="Not Connected"
                       Foreground="#FFAAAA" FontSize="11" FontWeight="SemiBold"/>
          </StackPanel>
        </Border>
      </Grid>
    </Border>

    <!-- MAIN CONTENT -->
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
      <StackPanel Margin="36,28,36,20">

        <!-- Connect Section -->
        <Border Background="White" BorderBrush="#DDEAF7" BorderThickness="1"
                CornerRadius="8" Padding="20,16" Margin="0,0,0,28">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" VerticalAlignment="Center">
              <TextBlock x:Name="TenantInfoText"
                         Text="Connect to your Microsoft 365 tenant to manage Exchange Online objects."
                         FontSize="12" Foreground="#445566" TextWrapping="Wrap"/>
            </StackPanel>
            <Button x:Name="BtnConnect" Grid.Column="1"
                    Cursor="Hand" Margin="16,0,0,0"
                    Padding="20,10" FontSize="13" FontWeight="SemiBold"
                    Background="#0078D4" Foreground="White"
                    BorderThickness="0">
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
              <StackPanel Orientation="Horizontal">
                <TextBlock x:Name="BtnConnectText" Text="Connect to M365 Tenant"
                           VerticalAlignment="Center"/>
              </StackPanel>
            </Button>
          </Grid>
        </Border>

        <!-- Feature Heading -->
        <TextBlock Text="Manage Objects" FontSize="16" FontWeight="Bold"
                   Foreground="#1A2D4A" Margin="0,0,0,6"/>
        <TextBlock Text="Select a feature to manage your Exchange Online environment."
                   FontSize="12" Foreground="#667788" Margin="0,0,0,20"/>

        <!-- Feature Tiles - 3 columns -->
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="20"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="20"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <!-- Distribution Groups -->
          <Button x:Name="BtnDG" Grid.Column="0" Style="{StaticResource FeatureCard}">
            <StackPanel HorizontalAlignment="Center">
              <TextBlock Text="&#x1F465;" FontSize="36"
                         HorizontalAlignment="Center" Margin="0,0,0,12"/>
              <TextBlock Text="Distribution Groups" FontSize="14" FontWeight="Bold"
                         Foreground="#1A2D4A" HorizontalAlignment="Center"
                         TextAlignment="Center" TextWrapping="Wrap" Margin="0,0,0,8"/>
              <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#667788"
                         HorizontalAlignment="Center" TextAlignment="Center"
                         Text="Create, update and manage distribution group members"
                         Margin="0,0,0,14"/>
              <Border Background="#EBF3FD" CornerRadius="4" Padding="10,4"
                      HorizontalAlignment="Center">
                <TextBlock Text="Create / Edit / Members" Foreground="#0078D4"
                           FontSize="10" FontWeight="SemiBold"/>
              </Border>
            </StackPanel>
          </Button>

          <!-- Shared Mailbox -->
          <Button x:Name="BtnSM" Grid.Column="2" Style="{StaticResource FeatureCard}">
            <StackPanel HorizontalAlignment="Center">
              <TextBlock Text="&#x1F4EC;" FontSize="36"
                         HorizontalAlignment="Center" Margin="0,0,0,12"/>
              <TextBlock Text="Shared Mailbox" FontSize="14" FontWeight="Bold"
                         Foreground="#1A2D4A" HorizontalAlignment="Center"
                         TextAlignment="Center" TextWrapping="Wrap" Margin="0,0,0,8"/>
              <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#667788"
                         HorizontalAlignment="Center" TextAlignment="Center"
                         Text="Create and manage shared mailboxes and permissions"
                         Margin="0,0,0,14"/>
              <Border Background="#EBF3FD" CornerRadius="4" Padding="10,4"
                      HorizontalAlignment="Center">
                <TextBlock Text="Create / Permissions" Foreground="#0078D4"
                           FontSize="10" FontWeight="SemiBold"/>
              </Border>
            </StackPanel>
          </Button>

          <!-- User Mailbox -->
          <Button x:Name="BtnUM" Grid.Column="4" Style="{StaticResource FeatureCard}">
            <StackPanel HorizontalAlignment="Center">
              <TextBlock Text="&#x1F464;" FontSize="36"
                         HorizontalAlignment="Center" Margin="0,0,0,12"/>
              <TextBlock Text="User Mailbox" FontSize="14" FontWeight="Bold"
                         Foreground="#1A2D4A" HorizontalAlignment="Center"
                         TextAlignment="Center" TextWrapping="Wrap" Margin="0,0,0,8"/>
              <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#667788"
                         HorizontalAlignment="Center" TextAlignment="Center"
                         Text="Manage user mailbox settings, quotas and properties"
                         Margin="0,0,0,14"/>
              <Border Background="#EBF3FD" CornerRadius="4" Padding="10,4"
                      HorizontalAlignment="Center">
                <TextBlock Text="Settings / Properties" Foreground="#0078D4"
                           FontSize="10" FontWeight="SemiBold"/>
              </Border>
            </StackPanel>
          </Button>

        </Grid>
      </StackPanel>
    </ScrollViewer>

    <!-- SPLITTER -->
    <GridSplitter Grid.Row="2" Height="5" HorizontalAlignment="Stretch"
                  Background="#2A4A7C" Cursor="SizeNS" VerticalAlignment="Center"/>

    <!-- ACTIVITY LOG -->
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
              <TextBlock x:Name="LogCount2" Text="" Foreground="#4A6A88"
                         FontSize="10" VerticalAlignment="Center"/>
            </StackPanel>
          </Grid>
        </Border>
        <RichTextBox x:Name="LogBox2" Grid.Row="1"
                     Background="#1E1E1E" BorderThickness="0"
                     IsReadOnly="True" IsDocumentEnabled="True"
                     FontFamily="Consolas,Courier New" FontSize="11.5"
                     Foreground="#C8D8E8"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Auto"
                     Padding="10,4,10,4"/>
      </Grid>
    </Border>

    <!-- FOOTER -->
    <Border Grid.Row="4" Background="#0C1E35">
      <TextBlock x:Name="FooterText" Text="Migraze v2.0  |  Microsoft 365"
                 Foreground="#4A7AAA" FontSize="10"
                 VerticalAlignment="Center" Margin="14,0"/>
    </Border>
  </Grid>
</Window>
"@

    $reader  = [System.Xml.XmlNodeReader]::new($xaml)
    $window  = [Windows.Markup.XamlReader]::Load($reader)
    if ($Owner) { $window.Owner = $Owner }

    # Save and redirect the shared log to this window's log box
    $savedLogBox        = $script:LogBox
    $savedLogCountLabel = $script:LogCountLabel
    $savedLogAutoScroll = $script:LogAutoScroll
    $savedLogEntryCount = $script:LogEntryCount

    $localLogBox = $window.FindName("LogBox2")
    $localLogBox.Document.PagePadding = [System.Windows.Thickness]::new(0)
    $script:LogBox        = $localLogBox
    $script:LogCountLabel = $window.FindName("LogCount2")
    $script:LogAutoScroll = $null

    # Connection status helpers
    $connDot    = $window.FindName("ConnDot")
    $connText   = $window.FindName("ConnStatusText")
    $connBadge  = $window.FindName("ConnStatusBadge")
    $tenantInfo = $window.FindName("TenantInfoText")
    $btnConnect = $window.FindName("BtnConnect")
    $btnConnTxt = $window.FindName("BtnConnectText")
    $footer     = $window.FindName("FooterText")

    function Update-ConnStatus {
        $connected = Get-MigrazeConnectionStatus
        if ($connected) {
            $connDot.Fill   = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#00C853"))
            $connText.Text  = "Connected"
            $connText.Foreground = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#AAFFCC"))
            $connBadge.Background = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0C2A10"))
            $btnConnTxt.Text = "Disconnect"
            $btnConnect.Background = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#C62828"))
            if ($script:M365TenantDomain) {
                $tenantInfo.Text = "Connected to: $($script:M365TenantDomain)"
                $footer.Text = "Migraze v2.0  |  Microsoft 365  |  $($script:M365TenantDomain)"
            }
            $window.FindName("BtnDG").IsEnabled = $true
            $window.FindName("BtnSM").IsEnabled = $true
            $window.FindName("BtnUM").IsEnabled = $true
        } else {
            $connDot.Fill   = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#FF5252"))
            $connText.Text  = "Not Connected"
            $connText.Foreground = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#FFAAAA"))
            $connBadge.Background = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#2A1010"))
            $btnConnTxt.Text = "Connect to M365 Tenant"
            $btnConnect.Background = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0078D4"))
            $tenantInfo.Text = "Connect to your Microsoft 365 tenant to manage Exchange Online objects."
            $footer.Text = "Migraze v2.0  |  Microsoft 365  |  Not Connected"
            $window.FindName("BtnDG").IsEnabled = $false
            $window.FindName("BtnSM").IsEnabled = $false
            $window.FindName("BtnUM").IsEnabled = $false
        }
    }

    # Initial state
    Update-ConnStatus

    # Connect / Disconnect button
    $btnConnect.Add_Click({
        if (Get-MigrazeConnectionStatus) {
            Write-MigrazeLog "Disconnecting from Microsoft 365..." "Action"
            Disconnect-MigrazeGraph
            Update-ConnStatus
        } else {
            Write-MigrazeLog "Connecting to Microsoft 365..." "Action"
            $ok = Connect-MigrazeGraph
            if ($ok) { Write-MigrazeLog "Connected to Microsoft 365." "Success" }
            Update-ConnStatus
        }
    })

    # Back button
    $window.FindName("BtnBack").Add_Click({ $window.Close() })

    # Feature buttons
    $window.FindName("BtnDG").Add_Click({
        Show-DistributionGroupsWindow -Owner $window
    })

    $window.FindName("BtnSM").Add_Click({
        Show-SharedMailboxWindow -Owner $window
    })

    $window.FindName("BtnUM").Add_Click({
        Show-UserMailboxWindow -Owner $window
    })

    Write-MigrazeLog "Microsoft 365 management opened." "Info"
    $window.ShowDialog() | Out-Null

    # Restore shared log to home window
    $script:LogBox        = $savedLogBox
    $script:LogCountLabel = $savedLogCountLabel
    $script:LogAutoScroll = $savedLogAutoScroll
    $script:LogEntryCount = $savedLogEntryCount
}
