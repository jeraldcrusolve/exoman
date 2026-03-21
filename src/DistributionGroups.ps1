# DistributionGroups.ps1 - Distribution Group management window

function Show-DistributionGroupsWindow {
    param([System.Windows.Window]$Owner)

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="ExoMan v1.0 – Distribution Groups"
    Width="980" Height="680"
    MinWidth="820" MinHeight="560"
    WindowStartupLocation="CenterOwner"
    ResizeMode="CanResizeWithGrip"
    Background="#F0F4F8">

  <Window.Resources>

    <Style x:Key="NavBtn" TargetType="Button">
      <Setter Property="Background"              Value="Transparent"/>
      <Setter Property="Foreground"              Value="#C8DCEE"/>
      <Setter Property="BorderThickness"         Value="0"/>
      <Setter Property="Padding"                 Value="16,11"/>
      <Setter Property="FontSize"                Value="13"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
      <Setter Property="Cursor"                  Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="5"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#1E3F6A"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="#142E52"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="ActionBtn" TargetType="Button">
      <Setter Property="Background"     Value="#0078D4"/>
      <Setter Property="Foreground"     Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"        Value="20,9"/>
      <Setter Property="FontSize"       Value="13"/>
      <Setter Property="Cursor"         Value="Hand"/>
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
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="#005A9E"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Background" Value="#B0BEC5"/>
                <Setter Property="Foreground" Value="#ECEFF1"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="SmallBtn" TargetType="Button" BasedOn="{StaticResource ActionBtn}">
      <Setter Property="Padding"   Value="12,7"/>
      <Setter Property="FontSize"  Value="12"/>
    </Style>

    <Style x:Key="Lbl" TargetType="TextBlock">
      <Setter Property="FontSize"     Value="12"/>
      <Setter Property="FontWeight"   Value="SemiBold"/>
      <Setter Property="Foreground"   Value="#2C3E50"/>
      <Setter Property="Margin"       Value="0,12,0,4"/>
    </Style>

    <Style x:Key="TB" TargetType="TextBox">
      <Setter Property="FontSize"       Value="13"/>
      <Setter Property="Padding"        Value="8,6"/>
      <Setter Property="BorderBrush"    Value="#C0CDD8"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Background"     Value="White"/>
    </Style>

    <Style x:Key="LB" TargetType="ListBox">
      <Setter Property="BorderBrush"    Value="#C0CDD8"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize"       Value="13"/>
      <Setter Property="Background"     Value="White"/>
      <Setter Property="Padding"        Value="2"/>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="230"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <!-- ═══ SIDEBAR ═══ -->
    <Border Grid.Column="0" Background="#0F2B50">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="68"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#081D36" Padding="16,0">
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="Distribution Groups" Foreground="White"
                       FontSize="14" FontWeight="Bold" TextWrapping="Wrap"/>
            <TextBlock Text="Exchange Online" Foreground="#7AAFD4" FontSize="11"/>
          </StackPanel>
        </Border>

        <StackPanel Grid.Row="1" Margin="8,14,8,0">
          <Button x:Name="NavCreate"        Content="➕   Create Distribution Group"  Style="{StaticResource NavBtn}"/>
          <Button x:Name="NavUpdate"        Content="✏️   Update DG Properties"       Style="{StaticResource NavBtn}" Margin="0,3,0,0"/>
          <Button x:Name="NavAddMembers"    Content="👤+  Add Members"                Style="{StaticResource NavBtn}" Margin="0,3,0,0"/>
          <Button x:Name="NavRemoveMembers" Content="👤–  Remove Members"             Style="{StaticResource NavBtn}" Margin="0,3,0,0"/>
          <Button x:Name="NavReadProps"     Content="📋   Read Current Properties"    Style="{StaticResource NavBtn}" Margin="0,3,0,0"/>
        </StackPanel>

        <Border Grid.Row="2" BorderBrush="#1A3A5C" BorderThickness="0,1,0,0">
          <StackPanel Margin="8,8,8,10">
            <Button x:Name="NavClose" Content="← Back to Home"
                    Style="{StaticResource NavBtn}"
                    Foreground="#FF8A80" FontSize="12"/>
            <TextBlock Text="Graph PowerShell SDK" Foreground="#4A7FAE"
                       FontSize="10" Margin="6,6,0,0"/>
          </StackPanel>
        </Border>
      </Grid>
    </Border>

    <!-- ═══ CONTENT AREA ═══ -->
    <Grid Grid.Column="1">

      <!-- ─── Panel: Create DG ─── -->
      <ScrollViewer x:Name="PanelCreate" Visibility="Visible"
                    VerticalScrollBarVisibility="Auto" Padding="30,26,30,20">
        <StackPanel>
          <TextBlock Text="Create Distribution Group" FontSize="22" FontWeight="Bold" Foreground="#0F2B50"/>
          <TextBlock Text="Creates a new mail-enabled distribution group in Exchange Online."
                     FontSize="12" Foreground="#667788" Margin="0,4,0,22"/>

          <TextBlock Text="Display Name *" Style="{StaticResource Lbl}"/>
          <TextBox x:Name="C_DisplayName" Style="{StaticResource TB}"
                   ToolTip="Name shown in the Global Address List"/>

          <TextBlock Text="Email Alias (MailNickname) *" Style="{StaticResource Lbl}"/>
          <TextBox x:Name="C_MailNickname" Style="{StaticResource TB}"
                   ToolTip="Part before @ in the email address. No spaces or special characters."/>

          <TextBlock Text="Description" Style="{StaticResource Lbl}"/>
          <TextBox x:Name="C_Description" Style="{StaticResource TB}"
                   Height="75" TextWrapping="Wrap" AcceptsReturn="True"
                   VerticalScrollBarVisibility="Auto"/>

          <CheckBox x:Name="C_SecurityEnabled"
                    Content="Also enable as a Security Group"
                    Margin="0,14,0,0" FontSize="13" Foreground="#2C3E50"/>

          <TextBlock x:Name="C_Status" Visibility="Collapsed"
                     FontSize="13" Margin="0,16,0,0"/>

          <Button x:Name="BtnCreate" Content="  Create Distribution Group  "
                  Style="{StaticResource ActionBtn}" HorizontalAlignment="Left" Margin="0,20,0,8"/>
        </StackPanel>
      </ScrollViewer>

      <!-- ─── Panel: Update DG Properties ─── -->
      <ScrollViewer x:Name="PanelUpdate" Visibility="Collapsed"
                    VerticalScrollBarVisibility="Auto" Padding="30,26,30,20">
        <StackPanel>
          <TextBlock Text="Update DG Properties" FontSize="22" FontWeight="Bold" Foreground="#0F2B50"/>
          <TextBlock Text="Search for a distribution group and edit its properties."
                     FontSize="12" Foreground="#667788" Margin="0,4,0,22"/>

          <TextBlock Text="Search Distribution Group" Style="{StaticResource Lbl}" Margin="0,0,0,4"/>
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="U_Search" Style="{StaticResource TB}"
                     ToolTip="Enter display name or email prefix"/>
            <Button x:Name="BtnUSearch" Grid.Column="1" Content="Search"
                    Style="{StaticResource SmallBtn}" Margin="8,0,0,0"/>
          </Grid>
          <ListBox x:Name="U_DGList" Style="{StaticResource LB}" Height="110" Margin="0,6,0,0"/>

          <Border x:Name="U_FieldsPanel" Visibility="Collapsed"
                  Background="White" CornerRadius="7" Padding="16,14"
                  Margin="0,16,0,0" BorderBrush="#D0DCE8" BorderThickness="1">
            <StackPanel>
              <TextBlock Text="Edit Properties" FontSize="14" FontWeight="SemiBold"
                         Foreground="#0F2B50" Margin="0,0,0,8"/>
              <TextBlock Text="Display Name" Style="{StaticResource Lbl}" Margin="0,0,0,4"/>
              <TextBox x:Name="U_DisplayName" Style="{StaticResource TB}"/>
              <TextBlock Text="Description" Style="{StaticResource Lbl}"/>
              <TextBox x:Name="U_Description" Style="{StaticResource TB}"
                       Height="70" TextWrapping="Wrap" AcceptsReturn="True"/>
              <TextBlock x:Name="U_Status" Visibility="Collapsed" FontSize="13" Margin="0,14,0,0"/>
              <Button x:Name="BtnUSave" Content="  Save Changes  "
                      Style="{StaticResource ActionBtn}" HorizontalAlignment="Left" Margin="0,16,0,4"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </ScrollViewer>

      <!-- ─── Panel: Add Members ─── -->
      <ScrollViewer x:Name="PanelAddMembers" Visibility="Collapsed"
                    VerticalScrollBarVisibility="Auto" Padding="30,26,30,20">
        <StackPanel>
          <TextBlock Text="Add Members" FontSize="22" FontWeight="Bold" Foreground="#0F2B50"/>
          <TextBlock Text="Add a user to a distribution group."
                     FontSize="12" Foreground="#667788" Margin="0,4,0,22"/>

          <TextBlock Text="Step 1 – Select Distribution Group" Style="{StaticResource Lbl}" Margin="0,0,0,4"/>
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <TextBox x:Name="AM_GrpSearch" Style="{StaticResource TB}"/>
            <Button x:Name="BtnAMGrpSearch" Grid.Column="1" Content="Search"
                    Style="{StaticResource SmallBtn}" Margin="8,0,0,0"/>
          </Grid>
          <ListBox x:Name="AM_GrpList" Style="{StaticResource LB}" Height="100" Margin="0,6,0,0"/>

          <TextBlock Text="Step 2 – Search User to Add" Style="{StaticResource Lbl}" Margin="0,16,0,4"/>
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <TextBox x:Name="AM_UsrSearch" Style="{StaticResource TB}"
                     ToolTip="Enter display name or UPN prefix"/>
            <Button x:Name="BtnAMUsrSearch" Grid.Column="1" Content="Search"
                    Style="{StaticResource SmallBtn}" Margin="8,0,0,0"/>
          </Grid>
          <ListBox x:Name="AM_UsrList" Style="{StaticResource LB}" Height="120" Margin="0,6,0,0"/>

          <TextBlock x:Name="AM_Status" Visibility="Collapsed" FontSize="13" Margin="0,14,0,0"/>
          <Button x:Name="BtnAMAdd" Content="  Add Selected User to Group  "
                  Style="{StaticResource ActionBtn}" HorizontalAlignment="Left"
                  Margin="0,16,0,8" IsEnabled="False"/>
        </StackPanel>
      </ScrollViewer>

      <!-- ─── Panel: Remove Members ─── -->
      <ScrollViewer x:Name="PanelRemoveMembers" Visibility="Collapsed"
                    VerticalScrollBarVisibility="Auto" Padding="30,26,30,20">
        <StackPanel>
          <TextBlock Text="Remove Members" FontSize="22" FontWeight="Bold" Foreground="#0F2B50"/>
          <TextBlock Text="Select a group, load its members, then remove one or more."
                     FontSize="12" Foreground="#667788" Margin="0,4,0,22"/>

          <TextBlock Text="Step 1 – Select Distribution Group" Style="{StaticResource Lbl}" Margin="0,0,0,4"/>
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <TextBox x:Name="RM_GrpSearch" Style="{StaticResource TB}"/>
            <Button x:Name="BtnRMGrpSearch" Grid.Column="1" Content="Search"
                    Style="{StaticResource SmallBtn}" Margin="8,0,0,0"/>
          </Grid>
          <ListBox x:Name="RM_GrpList" Style="{StaticResource LB}" Height="100" Margin="0,6,0,0"/>

          <TextBlock Text="Step 2 – Load and Select Members to Remove" Style="{StaticResource Lbl}" Margin="0,16,0,4"/>
          <Button x:Name="BtnRMLoadMembers" Content="Load Members"
                  Style="{StaticResource SmallBtn}" HorizontalAlignment="Left"
                  Margin="0,0,0,6" IsEnabled="False"/>
          <ListBox x:Name="RM_MbrList" Style="{StaticResource LB}" Height="150"
                   SelectionMode="Extended" Margin="0,0,0,0"
                   ToolTip="Hold Ctrl or Shift to select multiple members"/>

          <TextBlock x:Name="RM_Status" Visibility="Collapsed" FontSize="13" Margin="0,14,0,0"/>
          <Button x:Name="BtnRMRemove" Content="  Remove Selected Member(s)  "
                  Style="{StaticResource ActionBtn}" HorizontalAlignment="Left"
                  Margin="0,16,0,8" IsEnabled="False"/>
        </StackPanel>
      </ScrollViewer>

      <!-- ─── Panel: Read Properties ─── -->
      <Grid x:Name="PanelReadProps" Visibility="Collapsed" Margin="30,26,30,20">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0">
          <TextBlock Text="Read Current Properties" FontSize="22" FontWeight="Bold" Foreground="#0F2B50"/>
          <TextBlock Text="View all properties and current members of a distribution group."
                     FontSize="12" Foreground="#667788" Margin="0,4,0,22"/>
          <TextBlock Text="Search Distribution Group" Style="{StaticResource Lbl}" Margin="0,0,0,4"/>
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <TextBox x:Name="R_Search" Style="{StaticResource TB}"/>
            <Button x:Name="BtnRSearch" Grid.Column="1" Content="Search"
                    Style="{StaticResource SmallBtn}" Margin="8,0,0,0"/>
          </Grid>
        </StackPanel>

        <ListBox x:Name="R_DGList" Grid.Row="1" Style="{StaticResource LB}"
                 Height="100" Margin="0,6,0,0"/>

        <Button x:Name="BtnRLoad" Grid.Row="2" Content="Load Properties"
                Style="{StaticResource SmallBtn}" HorizontalAlignment="Left"
                Margin="0,10,0,10" IsEnabled="False"/>

        <Border x:Name="R_PropsBox" Grid.Row="3" Background="White" CornerRadius="7"
                BorderBrush="#D0DCE8" BorderThickness="1" Padding="14,12"
                Visibility="Collapsed" Margin="0,0,0,10">
          <StackPanel x:Name="R_PropsContent"/>
        </Border>

        <Border Grid.Row="4" Background="White" CornerRadius="7"
                BorderBrush="#D0DCE8" BorderThickness="1" Padding="14,12"
                x:Name="R_MembersBox" Visibility="Collapsed">
          <StackPanel>
            <TextBlock x:Name="R_MemberHeader" Text="Members (0)"
                       FontSize="13" FontWeight="SemiBold" Foreground="#0F2B50" Margin="0,0,0,6"/>
            <ListBox x:Name="R_MbrList" Style="{StaticResource LB}" MaxHeight="180"/>
          </StackPanel>
        </Border>
      </Grid>

    </Grid><!-- end content grid -->
  </Grid>
</Window>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    if ($Owner) { $window.Owner = $Owner }

    # ── Resolve controls ──
    $navCreate        = $window.FindName("NavCreate")
    $navUpdate        = $window.FindName("NavUpdate")
    $navAddMembers    = $window.FindName("NavAddMembers")
    $navRemoveMembers = $window.FindName("NavRemoveMembers")
    $navReadProps     = $window.FindName("NavReadProps")
    $navClose         = $window.FindName("NavClose")

    $pCreate        = $window.FindName("PanelCreate")
    $pUpdate        = $window.FindName("PanelUpdate")
    $pAddMembers    = $window.FindName("PanelAddMembers")
    $pRemoveMembers = $window.FindName("PanelRemoveMembers")
    $pReadProps     = $window.FindName("PanelReadProps")

    $allPanels  = @($pCreate, $pUpdate, $pAddMembers, $pRemoveMembers, $pReadProps)
    $allNavBtns = @($navCreate, $navUpdate, $navAddMembers, $navRemoveMembers, $navReadProps)

    # ── Nav panel switching ──
    $activeNavColor   = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0078D4"))
    $inactiveNavColor = [Windows.Media.Brushes]::Transparent

    function Switch-DGPanel {
        param([int]$Idx)
        for ($i = 0; $i -lt $allPanels.Count; $i++) {
            $allPanels[$i].Visibility    = if ($i -eq $Idx) { "Visible" } else { "Collapsed" }
            $allNavBtns[$i].Background   = if ($i -eq $Idx) { $activeNavColor } else { $inactiveNavColor }
            $allNavBtns[$i].FontWeight   = if ($i -eq $Idx) { "SemiBold" } else { "Normal" }
        }
    }
    Switch-DGPanel 0

    $navCreate.Add_Click({        Switch-DGPanel 0 })
    $navUpdate.Add_Click({        Switch-DGPanel 1 })
    $navAddMembers.Add_Click({    Switch-DGPanel 2 })
    $navRemoveMembers.Add_Click({ Switch-DGPanel 3 })
    $navReadProps.Add_Click({     Switch-DGPanel 4 })
    $navClose.Add_Click({         Write-ExoLog "Closed Distribution Groups window." "Info"; $window.Close() })

    # ─────────────────────────────────────────────────────
    # Helper: set status text with colour
    # ─────────────────────────────────────────────────────
    function Set-StatusText {
        param($ctrl, [string]$Msg, [string]$Type = "info")
        $ctrl.Text = $Msg
        $ctrl.Foreground = switch ($Type) {
            "success" { [Windows.Media.Brushes]::DarkGreen }
            "error"   { [Windows.Media.Brushes]::Crimson   }
            default   { [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0078D4")) }
        }
        $ctrl.Visibility = "Visible"
    }

    # Helper: populate a ListBox with group objects
    function Fill-GroupList {
        param($listBox, [array]$Groups)
        $listBox.Items.Clear()
        foreach ($g in $Groups) {
            $item = [PSCustomObject]@{
                Id          = $g.Id
                DisplayName = $g.DisplayName
                Mail        = $g.Mail
                ToString    = "$($g.DisplayName)  <$($g.Mail)>"
            }
            $listBox.Items.Add($item) | Out-Null
        }
        $listBox.DisplayMemberPath = "ToString"
    }

    # ─────────────────────────────────────────────────────
    # ── CREATE DG ──
    # ─────────────────────────────────────────────────────
    $cDisplayName    = $window.FindName("C_DisplayName")
    $cMailNickname   = $window.FindName("C_MailNickname")
    $cDescription    = $window.FindName("C_Description")
    $cSecurityEnabled= $window.FindName("C_SecurityEnabled")
    $cStatus         = $window.FindName("C_Status")
    $btnCreate       = $window.FindName("BtnCreate")

    $btnCreate.Add_Click({
        $dn  = $cDisplayName.Text.Trim()
        $mn  = $cMailNickname.Text.Trim()
        $desc= $cDescription.Text.Trim()
        $sec = $cSecurityEnabled.IsChecked -eq $true

        if (-not $dn) { Set-StatusText $cStatus "Display Name is required." "error"; return }
        if (-not $mn) { Set-StatusText $cStatus "Email Alias (MailNickname) is required." "error"; return }
        if ($mn -match '\s|[^a-zA-Z0-9._\-]') {
            Set-StatusText $cStatus "MailNickname must not contain spaces or special characters." "error"; return
        }

        Set-StatusText $cStatus "Creating distribution group…" "info"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $btnCreate.IsEnabled = $false

        $result = New-DGGroup -DisplayName $dn -MailNickname $mn -Description $desc -SecurityEnabled $sec

        $window.Cursor = $null
        $btnCreate.IsEnabled = $true

        if ($result.Success) {
            Set-StatusText $cStatus "✔  Distribution group '$dn' created successfully!" "success"
            $cDisplayName.Clear(); $cMailNickname.Clear(); $cDescription.Clear()
            $cSecurityEnabled.IsChecked = $false
        } else {
            Set-StatusText $cStatus "✖  Error: $($result.Error)" "error"
        }
    })

    # ─────────────────────────────────────────────────────
    # ── UPDATE DG PROPERTIES ──
    # ─────────────────────────────────────────────────────
    $uSearch       = $window.FindName("U_Search")
    $uDGList       = $window.FindName("U_DGList")
    $uFieldsPanel  = $window.FindName("U_FieldsPanel")
    $uDisplayName  = $window.FindName("U_DisplayName")
    $uDescription  = $window.FindName("U_Description")
    $uStatus       = $window.FindName("U_Status")
    $btnUSearch    = $window.FindName("BtnUSearch")
    $btnUSave      = $window.FindName("BtnUSave")

    $btnUSearch.Add_Click({
        $q = $uSearch.Text.Trim()
        if (-not $q) { return }
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGList -SearchQuery $q
        $window.Cursor = $null
        if ($result.Success) {
            Fill-GroupList $uDGList $result.Groups
            if ($result.Groups.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No distribution groups found for '$q'.", "No Results",
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
            }
        } else {
            [System.Windows.MessageBox]::Show("Search failed:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    $uDGList.Add_SelectionChanged({
        if ($uDGList.SelectedItem) {
            $uDisplayName.Text  = $uDGList.SelectedItem.DisplayName
            $uFieldsPanel.Visibility = "Visible"
            $uStatus.Visibility = "Collapsed"
        }
    })

    $btnUSave.Add_Click({
        $selected = $uDGList.SelectedItem
        if (-not $selected) { Set-StatusText $uStatus "Please select a group first." "error"; return }
        $newDN   = $uDisplayName.Text.Trim()
        $newDesc = $uDescription.Text.Trim()
        if (-not $newDN) { Set-StatusText $uStatus "Display Name cannot be empty." "error"; return }

        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $btnUSave.IsEnabled = $false
        $result = Update-DGGroup -GroupId $selected.Id -DisplayName $newDN -Description $newDesc
        $window.Cursor = $null
        $btnUSave.IsEnabled = $true

        if ($result.Success) {
            Set-StatusText $uStatus "✔  Properties updated successfully." "success"
        } else {
            Set-StatusText $uStatus "✖  Error: $($result.Error)" "error"
        }
    })

    # ─────────────────────────────────────────────────────
    # ── ADD MEMBERS ──
    # ─────────────────────────────────────────────────────
    $amGrpSearch  = $window.FindName("AM_GrpSearch")
    $amGrpList    = $window.FindName("AM_GrpList")
    $amUsrSearch  = $window.FindName("AM_UsrSearch")
    $amUsrList    = $window.FindName("AM_UsrList")
    $amStatus     = $window.FindName("AM_Status")
    $btnAMGrpSrc  = $window.FindName("BtnAMGrpSearch")
    $btnAMUsrSrc  = $window.FindName("BtnAMUsrSearch")
    $btnAMAdd     = $window.FindName("BtnAMAdd")

    $btnAMGrpSrc.Add_Click({
        $q = $amGrpSearch.Text.Trim()
        if (-not $q) { return }
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGList -SearchQuery $q
        $window.Cursor = $null
        if ($result.Success) { Fill-GroupList $amGrpList $result.Groups }
        else {
            [System.Windows.MessageBox]::Show("Search failed:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    $btnAMUsrSrc.Add_Click({
        $q = $amUsrSearch.Text.Trim()
        if (-not $q) { return }
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Search-ExoManUsers -Query $q
        $window.Cursor = $null
        $amUsrList.Items.Clear()
        if ($result.Success) {
            foreach ($u in $result.Users) {
                $item = [PSCustomObject]@{
                    Id          = $u.Id
                    DisplayName = $u.DisplayName
                    UPN         = $u.UserPrincipalName
                    ToString    = "$($u.DisplayName)  ($($u.UserPrincipalName))"
                }
                $amUsrList.Items.Add($item) | Out-Null
            }
            $amUsrList.DisplayMemberPath = "ToString"
        } else {
            [System.Windows.MessageBox]::Show("User search failed:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    # Enable Add button only when both list items are selected
    $amGrpList.Add_SelectionChanged({ $btnAMAdd.IsEnabled = ($amGrpList.SelectedItem -and $amUsrList.SelectedItem) })
    $amUsrList.Add_SelectionChanged({ $btnAMAdd.IsEnabled = ($amGrpList.SelectedItem -and $amUsrList.SelectedItem) })

    $btnAMAdd.Add_Click({
        $grp = $amGrpList.SelectedItem
        $usr = $amUsrList.SelectedItem
        if (-not $grp -or -not $usr) { Set-StatusText $amStatus "Select both a group and a user." "error"; return }
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $btnAMAdd.IsEnabled = $false
        $result = Add-DGMember -GroupId $grp.Id -UserId $usr.Id
        $window.Cursor = $null
        $btnAMAdd.IsEnabled = $true
        if ($result.Success) {
            Set-StatusText $amStatus "✔  $($usr.DisplayName) added to $($grp.DisplayName) successfully." "success"
        } else {
            Set-StatusText $amStatus "✖  Error: $($result.Error)" "error"
        }
    })

    # ─────────────────────────────────────────────────────
    # ── REMOVE MEMBERS ──
    # ─────────────────────────────────────────────────────
    $rmGrpSearch    = $window.FindName("RM_GrpSearch")
    $rmGrpList      = $window.FindName("RM_GrpList")
    $rmMbrList      = $window.FindName("RM_MbrList")
    $rmStatus       = $window.FindName("RM_Status")
    $btnRMGrpSrc    = $window.FindName("BtnRMGrpSearch")
    $btnRMLoadMbrs  = $window.FindName("BtnRMLoadMembers")
    $btnRMRemove    = $window.FindName("BtnRMRemove")

    $btnRMGrpSrc.Add_Click({
        $q = $rmGrpSearch.Text.Trim()
        if (-not $q) { return }
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGList -SearchQuery $q
        $window.Cursor = $null
        if ($result.Success) {
            Fill-GroupList $rmGrpList $result.Groups
            $btnRMLoadMbrs.IsEnabled = ($rmGrpList.Items.Count -gt 0)
        } else {
            [System.Windows.MessageBox]::Show("Search failed:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    $rmGrpList.Add_SelectionChanged({ $btnRMLoadMbrs.IsEnabled = ($null -ne $rmGrpList.SelectedItem) })

    $btnRMLoadMbrs.Add_Click({
        $grp = $rmGrpList.SelectedItem
        if (-not $grp) { return }
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGProperties -GroupId $grp.Id
        $window.Cursor = $null
        $rmMbrList.Items.Clear()
        if ($result.Success) {
            foreach ($m in $result.Members) {
                $item = [PSCustomObject]@{
                    Id          = $m.Id
                    DisplayName = if ($m.AdditionalProperties.displayName) { $m.AdditionalProperties.displayName } else { $m.Id }
                    ToString    = if ($m.AdditionalProperties.displayName) { "$($m.AdditionalProperties.displayName)  ($($m.AdditionalProperties.userPrincipalName))" } else { $m.Id }
                }
                $rmMbrList.Items.Add($item) | Out-Null
            }
            $rmMbrList.DisplayMemberPath = "ToString"
            Set-StatusText $rmStatus "$($result.Members.Count) member(s) loaded." "info"
        } else {
            Set-StatusText $rmStatus "✖  Failed to load members: $($result.Error)" "error"
        }
    })

    $rmMbrList.Add_SelectionChanged({ $btnRMRemove.IsEnabled = ($rmMbrList.SelectedItems.Count -gt 0) })

    $btnRMRemove.Add_Click({
        $grp = $rmGrpList.SelectedItem
        if (-not $grp) { Set-StatusText $rmStatus "No group selected." "error"; return }
        $selected = @($rmMbrList.SelectedItems)
        if ($selected.Count -eq 0) { Set-StatusText $rmStatus "No member selected." "error"; return }

        $confirm = [System.Windows.MessageBox]::Show(
            "Remove $($selected.Count) member(s) from '$($grp.DisplayName)'?",
            "Confirm Removal",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $btnRMRemove.IsEnabled = $false
        $errors = @()
        foreach ($mbr in $selected) {
            $res = Remove-DGMember -GroupId $grp.Id -MemberId $mbr.Id
            if (-not $res.Success) { $errors += $mbr.DisplayName }
        }
        $window.Cursor = $null
        $btnRMRemove.IsEnabled = $true

        if ($errors.Count -eq 0) {
            Set-StatusText $rmStatus "✔  $($selected.Count) member(s) removed successfully." "success"
            # Reload member list
            $btnRMLoadMbrs.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
        } else {
            Set-StatusText $rmStatus "⚠  Some removals failed: $($errors -join ', ')" "error"
        }
    })

    # ─────────────────────────────────────────────────────
    # ── READ PROPERTIES ──
    # ─────────────────────────────────────────────────────
    $rSearch      = $window.FindName("R_Search")
    $rDGList      = $window.FindName("R_DGList")
    $rPropsBox    = $window.FindName("R_PropsBox")
    $rPropsContent= $window.FindName("R_PropsContent")
    $rMembersBox  = $window.FindName("R_MembersBox")
    $rMbrList     = $window.FindName("R_MbrList")
    $rMbrHeader   = $window.FindName("R_MemberHeader")
    $btnRSearch   = $window.FindName("BtnRSearch")
    $btnRLoad     = $window.FindName("BtnRLoad")

    $btnRSearch.Add_Click({
        $q = $rSearch.Text.Trim()
        if (-not $q) { return }
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGList -SearchQuery $q
        $window.Cursor = $null
        if ($result.Success) {
            Fill-GroupList $rDGList $result.Groups
            $btnRLoad.IsEnabled = ($rDGList.Items.Count -gt 0)
        } else {
            [System.Windows.MessageBox]::Show("Search failed:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    $rDGList.Add_SelectionChanged({ $btnRLoad.IsEnabled = ($null -ne $rDGList.SelectedItem) })

    $btnRLoad.Add_Click({
        $grp = $rDGList.SelectedItem
        if (-not $grp) { return }
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGProperties -GroupId $grp.Id
        $window.Cursor = $null
        if (-not $result.Success) {
            [System.Windows.MessageBox]::Show("Failed to load properties:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            return
        }
        $g = $result.Group

        # Build properties grid
        $rPropsContent.Children.Clear()
        $props = [ordered]@{
            "Display Name"    = $g.DisplayName
            "Email Address"   = $g.Mail
            "Mail Nickname"   = $g.MailNickname
            "Description"     = if ($g.Description) { $g.Description } else { "(none)" }
            "Mail Enabled"    = $g.MailEnabled
            "Security Enabled"= $g.SecurityEnabled
            "Group ID"        = $g.Id
            "Created"         = if ($g.CreatedDateTime) { $g.CreatedDateTime.ToString("yyyy-MM-dd HH:mm UTC") } else { "—" }
        }
        foreach ($k in $props.Keys) {
            $row = [System.Windows.Controls.Grid]::new()
            $col1 = [System.Windows.Controls.ColumnDefinition]::new(); $col1.Width = [System.Windows.GridLength]::new(160)
            $col2 = [System.Windows.Controls.ColumnDefinition]::new(); $col2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $row.ColumnDefinitions.Add($col1); $row.ColumnDefinitions.Add($col2)
            $row.Margin = [System.Windows.Thickness]::new(0,2,0,2)

            $lbl = [System.Windows.Controls.TextBlock]::new()
            $lbl.Text = $k; $lbl.FontWeight = "SemiBold"; $lbl.Foreground = [Windows.Media.Brushes]::DimGray
            $lbl.FontSize = 12; [System.Windows.Controls.Grid]::SetColumn($lbl, 0)

            $val = [System.Windows.Controls.TextBlock]::new()
            $val.Text = "$($props[$k])"; $val.FontSize = 12; $val.TextWrapping = "Wrap"
            $val.Foreground = [Windows.Media.Brushes]::Black
            [System.Windows.Controls.Grid]::SetColumn($val, 1)

            $row.Children.Add($lbl) | Out-Null; $row.Children.Add($val) | Out-Null
            $rPropsContent.Children.Add($row) | Out-Null
        }
        $rPropsBox.Visibility = "Visible"

        # Members
        $rMbrList.Items.Clear()
        $rMbrHeader.Text = "Members ($($result.Members.Count))"
        foreach ($m in $result.Members) {
            $dn  = if ($m.AdditionalProperties.displayName) { $m.AdditionalProperties.displayName } else { $m.Id }
            $upn = if ($m.AdditionalProperties.userPrincipalName) { " ($($m.AdditionalProperties.userPrincipalName))" } else { "" }
            $rMbrList.Items.Add("$dn$upn") | Out-Null
        }
        $rMembersBox.Visibility = "Visible"
    })

    $window.ShowDialog() | Out-Null
}
