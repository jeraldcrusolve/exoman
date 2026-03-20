# SharedMailbox.ps1 - Shared Mailbox management window (coming soon)

function Show-SharedMailboxWindow {
    param([System.Windows.Window]$Owner)

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="ExoMan v1.0 – Shared Mailbox"
    Width="680" Height="420"
    WindowStartupLocation="CenterOwner"
    ResizeMode="CanMinimize"
    Background="#F0F4F8">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="68"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#0F2B50">
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="24,0">
        <TextBlock Text="ExoMan" FontSize="26" FontWeight="Bold" Foreground="White"/>
        <TextBlock Text=" v1.0" FontSize="14" Foreground="#AAC8E8" VerticalAlignment="Bottom" Margin="0,0,0,3"/>
        <Rectangle Width="1" Fill="#3A5878" Margin="16,8,16,8"/>
        <TextBlock Text="Shared Mailbox Management" FontSize="13" Foreground="#AAC8E8" VerticalAlignment="Center"/>
      </StackPanel>
    </Border>

    <StackPanel Grid.Row="1" VerticalAlignment="Center" HorizontalAlignment="Center">
      <TextBlock Text="📬" FontSize="64" HorizontalAlignment="Center" Margin="0,0,0,16"/>
      <TextBlock Text="Shared Mailbox Management" FontSize="22" FontWeight="Bold"
                 Foreground="#0F2B50" HorizontalAlignment="Center"/>
      <TextBlock Text="This feature is coming in a future version of ExoMan."
                 FontSize="14" Foreground="#667788" HorizontalAlignment="Center" Margin="0,10,0,4"/>
      <TextBlock Text="Planned: Create, configure, delegate access, and manage shared mailboxes."
                 FontSize="12" Foreground="#889AAA" HorizontalAlignment="Center"
                 TextAlignment="Center" Margin="40,6,40,0"/>
    </StackPanel>
  </Grid>
</Window>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    if ($Owner) { $window.Owner = $Owner }
    $window.ShowDialog() | Out-Null
}
