# MainWindow.ps1 - Migraze v2.0 main application window

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
      <Grid Margin="24,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Migraze" FontSize="26" FontWeight="Bold"
                     Foreground="White" VerticalAlignment="Center"/>
          <TextBlock Text=" v2.0" FontSize="13" Foreground="#AAC8E8"
                     VerticalAlignment="Bottom" Margin="2,0,0,5"/>
          <Rectangle Width="1" Fill="#3A5878" Margin="16,10,16,10"/>
          <TextBlock Text="Cloud Management Platform" FontSize="12"
                     Foreground="#AAC8E8" VerticalAlignment="Center"/>
        </StackPanel>
        <Border Grid.Column="1" Background="#1B4F8A" CornerRadius="4"
                Padding="12,5" VerticalAlignment="Center">
          <TextBlock Text="Admin Toolkit" Foreground="White" FontSize="11"/>
        </Border>
      </Grid>
    </Border>

    <!-- ENVIRONMENT SELECTION -->
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
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

          <!-- Card 1: Google Workspace (Coming Soon) -->
          <Border Grid.Column="0" Background="#FAFAFA" BorderBrush="#E0E8F0"
                  BorderThickness="1" CornerRadius="10" Padding="28,24">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <TextBlock Grid.Row="0" Text="&#x1F4E7;" FontSize="44"
                         HorizontalAlignment="Center" Margin="0,0,0,14" Opacity="0.45"/>
              <TextBlock Grid.Row="1" Text="Google Workspace" FontSize="16" FontWeight="Bold"
                         Foreground="#9AAAB8" HorizontalAlignment="Center" Margin="0,0,0,10"/>
              <TextBlock Grid.Row="2" TextWrapping="Wrap" FontSize="12" Foreground="#AABBCC"
                         HorizontalAlignment="Center" TextAlignment="Center"
                         Text="Manage users, groups, shared drives and Gmail settings in your Google Workspace environment."
                         Margin="0,0,0,18"/>
              <Border Grid.Row="3" Background="#E8EFF5" CornerRadius="20"
                      Padding="16,6" HorizontalAlignment="Center">
                <TextBlock Text="Coming Soon" Foreground="#8899AA" FontSize="11"
                           FontWeight="SemiBold"/>
              </Border>
            </Grid>
          </Border>

          <!-- Card 2: Microsoft 365 (Active) -->
          <Button x:Name="BtnM365" Grid.Column="2" Cursor="Hand"
                  Background="White" BorderBrush="#DDEAF7" BorderThickness="1"
                  Padding="0" HorizontalContentAlignment="Stretch">
            <Button.Template>
              <ControlTemplate TargetType="Button">
                <Border x:Name="CB" Background="{TemplateBinding Background}"
                        BorderBrush="{TemplateBinding BorderBrush}"
                        BorderThickness="{TemplateBinding BorderThickness}"
                        CornerRadius="10" Padding="28,24">
                  <ContentPresenter/>
                </Border>
                <ControlTemplate.Triggers>
                  <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="CB" Property="BorderBrush" Value="#0078D4"/>
                    <Setter TargetName="CB" Property="Background" Value="#F0F7FF"/>
                  </Trigger>
                  <Trigger Property="IsPressed" Value="True">
                    <Setter TargetName="CB" Property="Background" Value="#E3F0FF"/>
                  </Trigger>
                </ControlTemplate.Triggers>
              </ControlTemplate>
            </Button.Template>
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <TextBlock Grid.Row="0" Text="&#x2601;" FontSize="44"
                         HorizontalAlignment="Center" Margin="0,0,0,14"
                         Foreground="#0078D4"/>
              <TextBlock Grid.Row="1" Text="Microsoft 365" FontSize="16" FontWeight="Bold"
                         Foreground="#1A2D4A" HorizontalAlignment="Center" Margin="0,0,0,10"/>
              <TextBlock Grid.Row="2" TextWrapping="Wrap" FontSize="12" Foreground="#556677"
                         HorizontalAlignment="Center" TextAlignment="Center"
                         Text="Manage distribution groups, shared mailboxes and user mailboxes in your Microsoft 365 tenant."
                         Margin="0,0,0,18"/>
              <DockPanel Grid.Row="3" LastChildFill="False">
                <Border Background="#EBF3FD" CornerRadius="4" Padding="10,4" DockPanel.Dock="Left">
                  <TextBlock Text="Exchange Online" Foreground="#0078D4" FontSize="10" FontWeight="SemiBold"/>
                </Border>
                <TextBlock Text="&#x2192;" FontSize="20" Foreground="#0078D4"
                           DockPanel.Dock="Right" VerticalAlignment="Center"/>
              </DockPanel>
            </Grid>
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

    <!-- FOOTER -->
    <Border Grid.Row="4" Background="#0C1E35">
      <TextBlock Text="Migraze v2.0  |  Ready"
                 Foreground="#4A7AAA" FontSize="10"
                 VerticalAlignment="Center" Margin="14,0"/>
    </Border>
  </Grid>
</Window>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Wire up shared activity log
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

    $window.FindName("BtnM365").Add_Click({
        Show-M365HomeWindow -Owner $window
    })

    Write-MigrazeLog "Migraze v2.0 started" "Action"
    Write-MigrazeLog "Select an environment to get started." "Info"

    $window.ShowDialog() | Out-Null
}