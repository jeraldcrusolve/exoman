# MainWindow.ps1 - Migraze v2.0 home screen with scenario selection

function Show-MainWindow {

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Migraze v2.0 - Migration Management Platform"
    Width="920" Height="700"
    MinWidth="780" MinHeight="580"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResizeWithGrip"
    Background="#F0F4F8">

  <Window.Resources>
    <Style x:Key="ScenarioCard" TargetType="Button">
      <Setter Property="Background"      Value="White"/>
      <Setter Property="Foreground"      Value="#1A2D4A"/>
      <Setter Property="BorderBrush"     Value="#DDEAF7"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="12" Padding="24,22">
              <ContentPresenter HorizontalAlignment="Stretch" VerticalAlignment="Stretch"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background"  Value="#E8F4FF"/>
                <Setter Property="BorderBrush" Value="#0078D4"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="#D0E8F8"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="72"/>
      <RowDefinition Height="*" MinHeight="220"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="190" MinHeight="120"/>
      <RowDefinition Height="30"/>
    </Grid.RowDefinitions>

    <!-- HEADER -->
    <Border Grid.Row="0" Background="#0F2B50">
      <Grid Margin="24,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Migraze" FontSize="30" FontWeight="Bold"
                     Foreground="White" VerticalAlignment="Center"/>
          <TextBlock Text=" v2.0" FontSize="16" Foreground="#AAC8E8"
                     VerticalAlignment="Bottom" Margin="0,0,0,4"/>
          <Rectangle Width="1" Fill="#3A5878" Margin="18,8,18,8"/>
          <TextBlock Text="Migration Management Platform" FontSize="13"
                     Foreground="#AAC8E8" VerticalAlignment="Center"/>
        </StackPanel>
        <Border Grid.Column="1" Background="#1A3F6A" CornerRadius="4"
                Padding="10,4" VerticalAlignment="Center">
          <TextBlock Text="Migration Toolkit" Foreground="#AAC8E8" FontSize="11"/>
        </Border>
      </Grid>
    </Border>

    <!-- SCENARIO SELECTION -->
    <Grid Grid.Row="1" Margin="32,28,32,20">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <TextBlock Grid.Row="0" Text="Select Migration Scenario"
                 FontSize="20" FontWeight="Bold" Foreground="#1A2D4A"
                 Margin="0,0,0,6"/>
      <TextBlock Grid.Row="1"
                 Text="Choose your migration type to begin discovery and migration tasks."
                 FontSize="12" Foreground="#667788" Margin="0,0,0,22"/>

      <Grid Grid.Row="2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition/>
          <ColumnDefinition Width="20"/>
          <ColumnDefinition/>
        </Grid.ColumnDefinitions>

        <!-- Card 1: Google Workspace to M365 -->
        <Button x:Name="BtnScenarioGW" Grid.Column="0" Style="{StaticResource ScenarioCard}">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Text="&#x1F30E;" FontSize="44"
                       HorizontalAlignment="Center" Margin="0,0,0,12"/>
            <StackPanel Grid.Row="1" HorizontalAlignment="Center" Margin="0,0,0,4">
              <TextBlock Text="Google Workspace" FontSize="17" FontWeight="Bold"
                         Foreground="#1A2D4A" HorizontalAlignment="Center"/>
              <TextBlock Text="&#x2192;  Microsoft 365" FontSize="14"
                         Foreground="#0078D4" HorizontalAlignment="Center" Margin="0,2,0,0"/>
            </StackPanel>
            <TextBlock Grid.Row="2"
                       Text="Discover users, groups, mailboxes and drives from Google Workspace, then create and migrate objects to Microsoft 365."
                       FontSize="11" Foreground="#556677" TextAlignment="Center"
                       TextWrapping="Wrap" HorizontalAlignment="Center"
                       Margin="8,10,8,14"/>
            <Border Grid.Row="4" Background="#EAF4FF" CornerRadius="4"
                    Padding="10,4" HorizontalAlignment="Center">
              <TextBlock Text="Discovery + Migration" FontSize="10"
                         Foreground="#0078D4" FontWeight="SemiBold"/>
            </Border>
          </Grid>
        </Button>

        <!-- Card 2: M365 Tenant to Tenant -->
        <Button x:Name="BtnScenarioM365" Grid.Column="2" Style="{StaticResource ScenarioCard}">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Text="&#x1F504;" FontSize="44"
                       HorizontalAlignment="Center" Margin="0,0,0,12"/>
            <StackPanel Grid.Row="1" HorizontalAlignment="Center" Margin="0,0,0,4">
              <TextBlock Text="M365 Tenant to Tenant" FontSize="17" FontWeight="Bold"
                         Foreground="#1A2D4A" HorizontalAlignment="Center"/>
              <TextBlock Text="Migration" FontSize="14"
                         Foreground="#0078D4" HorizontalAlignment="Center" Margin="0,2,0,0"/>
            </StackPanel>
            <TextBlock Grid.Row="2"
                       Text="Discover users, groups, mailboxes and contacts from a source M365 tenant, then recreate and migrate to the target tenant."
                       FontSize="11" Foreground="#556677" TextAlignment="Center"
                       TextWrapping="Wrap" HorizontalAlignment="Center"
                       Margin="8,10,8,14"/>
            <Border Grid.Row="4" Background="#EAF4FF" CornerRadius="4"
                    Padding="10,4" HorizontalAlignment="Center">
              <TextBlock Text="Tenant to Tenant" FontSize="10"
                         Foreground="#0078D4" FontWeight="SemiBold"/>
            </Border>
          </Grid>
        </Button>

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
              <Ellipse Width="7" Height="7" Fill="#00C853"
                       VerticalAlignment="Center" Margin="0,0,7,0"/>
              <TextBlock Text="Activity Log" Foreground="#7AABCC"
                         FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
              <TextBlock x:Name="LogCount" Text="  (0 entries)"
                         Foreground="#4A6A88" FontSize="10" VerticalAlignment="Center"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
              <CheckBox x:Name="LogAutoScroll" Content="Auto-scroll" IsChecked="True"
                        Foreground="#5A8AAA" FontSize="10"
                        VerticalAlignment="Center" Margin="0,0,12,0"/>
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
                     HorizontalScrollBarVisibility="Auto"
                     Padding="10,4,10,4"/>
      </Grid>
    </Border>

    <!-- FOOTER -->
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

    $footerText  = $window.FindName("FooterText")

    # Wire up activity log
    $script:LogBox        = $window.FindName("LogBox")
    $script:LogCountLabel = $window.FindName("LogCount")
    $script:LogAutoScroll = $window.FindName("LogAutoScroll")
    $script:LogEntryCount = 0
    $script:LogBox.Document.PagePadding = [System.Windows.Thickness]::new(0)

    $window.FindName("BtnClearLog").Add_Click({
        $script:LogBox.Document.Blocks.Clear()
        $script:LogEntryCount = 0
        $script:LogCountLabel.Text = "  (0 entries)"
    })

    # Scenario buttons
    $window.FindName("BtnScenarioGW").Add_Click({
        Write-MigrazeLog "Opening Google Workspace to M365 scenario..." "Action"
        $footerText.Text = "Migraze v2.0  |  Google Workspace Migration"
        Show-GWtoM365Window -Owner $window
        $footerText.Text = "Migraze v2.0  |  Ready"
    })

    $window.FindName("BtnScenarioM365").Add_Click({
        Write-MigrazeLog "Opening M365 Tenant to Tenant Migration scenario..." "Action"
        $footerText.Text = "Migraze v2.0  |  M365 Tenant to Tenant Migration"
        Show-M365toM365Window -Owner $window
        $footerText.Text = "Migraze v2.0  |  Ready"
    })

    # Startup log
    Write-MigrazeLog "Migraze v2.0 started." "Action"
    Write-MigrazeLog "Select a migration scenario to begin." "Info"

    $window.ShowDialog() | Out-Null
}