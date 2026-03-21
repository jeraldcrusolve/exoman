# UserMailbox.ps1 - User Mailbox management window (coming soon)

function Show-UserMailboxWindow {
    param([System.Windows.Window]$Owner)

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="ExoMan v1.0 – User Mailbox"
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
        <TextBlock Text="User Mailbox Management" FontSize="13" Foreground="#AAC8E8" VerticalAlignment="Center"/>
      </StackPanel>
    </Border>

    <StackPanel Grid.Row="1" VerticalAlignment="Center" HorizontalAlignment="Center">
      <TextBlock Text="👤" FontSize="64" HorizontalAlignment="Center" Margin="0,0,0,16"/>
      <TextBlock Text="User Mailbox Management" FontSize="22" FontWeight="Bold"
                 Foreground="#0F2B50" HorizontalAlignment="Center"/>
      <TextBlock Text="This feature is coming in a future version of ExoMan."
                 FontSize="14" Foreground="#667788" HorizontalAlignment="Center" Margin="0,10,0,4"/>
      <TextBlock Text="Planned: Manage mailbox settings, forwarding, quotas, permissions, and more."
                 FontSize="12" Foreground="#889AAA" HorizontalAlignment="Center"
                 TextAlignment="Center" Margin="40,6,40,0"/>
      <Button x:Name="BtnClose" Content="← Back to Home"
              Background="#0078D4" Foreground="White" BorderThickness="0"
              Padding="20,9" FontSize="13" Cursor="Hand" Margin="0,28,0,0"
              HorizontalAlignment="Center"/>
    </StackPanel>
  </Grid>
</Window>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    if ($Owner) { $window.Owner = $Owner }
    $window.FindName("BtnClose").Add_Click({ $window.Close() })
    $window.ShowDialog() | Out-Null
}
