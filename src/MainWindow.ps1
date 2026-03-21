# MainWindow.ps1 - Migraze v2.0 main application window (single-window navigation)

function Show-MainWindow {

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Migraze v2.0 - Cloud Management Platform"
    Width="1060" Height="700"
    MinWidth="900" MinHeight="560"
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

      <!-- ─────────────────── VIEW: DISTRIBUTION GROUPS ─────────────────── -->
      <Grid x:Name="ViewDG" Visibility="Collapsed">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="200"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Sidebar -->
        <Border Grid.Column="0" Background="#0F2B50">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="68"/>
              <RowDefinition Height="*"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <Border Grid.Row="0" Background="#081D36" Padding="16,0">
              <StackPanel VerticalAlignment="Center">
                <TextBlock Text="Distribution Groups" Foreground="White" FontSize="14" FontWeight="Bold" TextWrapping="Wrap"/>
                <TextBlock Text="Exchange Online" Foreground="#7AAFD4" FontSize="11"/>
              </StackPanel>
            </Border>

            <StackPanel Grid.Row="1" Margin="8,14,8,0">
              <Button x:Name="NavCreate"        Content="&#x2795;   Create Distribution Group"  Style="{StaticResource NavBtn}"/>
              <Button x:Name="NavUpdate"        Content="&#x270F;&#xFE0F;   Update DG Properties"       Style="{StaticResource NavBtn}" Margin="0,3,0,0"/>
              <Button x:Name="NavAddMembers"    Content="&#x1F464;+  Add Members"                Style="{StaticResource NavBtn}" Margin="0,3,0,0"/>
              <Button x:Name="NavRemoveMembers" Content="&#x1F464;&#x2013;  Remove Members"             Style="{StaticResource NavBtn}" Margin="0,3,0,0"/>
              <Button x:Name="NavReadProps"     Content="&#x1F4CB;   Read Current Properties"    Style="{StaticResource NavBtn}" Margin="0,3,0,0"/>
              <Button x:Name="NavDiscover"      Content="&#x1F50D;   Discover All DGs"           Style="{StaticResource NavBtn}" Margin="0,3,0,0"/>
            </StackPanel>

            <Border Grid.Row="2" BorderBrush="#1A3A5C" BorderThickness="0,1,0,0">
              <StackPanel Margin="8,8,8,10">
                <Button x:Name="DGBtnBack" Content="&#x2190; Back to M365"
                        Style="{StaticResource NavBtn}"
                        Foreground="#FF8A80" FontSize="12"/>
                <TextBlock Text="Graph PowerShell SDK" Foreground="#4A7FAE"
                           FontSize="10" Margin="6,6,0,0"/>
              </StackPanel>
            </Border>
          </Grid>
        </Border>

        <!-- Content area -->
        <Grid Grid.Column="1">

          <!-- Panel: Create DG -->
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

          <!-- Panel: Update DG Properties -->
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

          <!-- Panel: Add Members -->
          <ScrollViewer x:Name="PanelAddMembers" Visibility="Collapsed"
                        VerticalScrollBarVisibility="Auto" Padding="30,26,30,20">
            <StackPanel>
              <TextBlock Text="Add Members" FontSize="22" FontWeight="Bold" Foreground="#0F2B50"/>
              <TextBlock Text="Add a user to a distribution group."
                         FontSize="12" Foreground="#667788" Margin="0,4,0,22"/>

              <TextBlock Text="Step 1 &#x2013; Select Distribution Group" Style="{StaticResource Lbl}" Margin="0,0,0,4"/>
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <TextBox x:Name="AM_GrpSearch" Style="{StaticResource TB}"/>
                <Button x:Name="BtnAMGrpSearch" Grid.Column="1" Content="Search"
                        Style="{StaticResource SmallBtn}" Margin="8,0,0,0"/>
              </Grid>
              <ListBox x:Name="AM_GrpList" Style="{StaticResource LB}" Height="100" Margin="0,6,0,0"/>

              <TextBlock Text="Step 2 &#x2013; Search User to Add" Style="{StaticResource Lbl}" Margin="0,16,0,4"/>
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

          <!-- Panel: Remove Members -->
          <ScrollViewer x:Name="PanelRemoveMembers" Visibility="Collapsed"
                        VerticalScrollBarVisibility="Auto" Padding="30,26,30,20">
            <StackPanel>
              <TextBlock Text="Remove Members" FontSize="22" FontWeight="Bold" Foreground="#0F2B50"/>
              <TextBlock Text="Select a group, load its members, then remove one or more."
                         FontSize="12" Foreground="#667788" Margin="0,4,0,22"/>

              <TextBlock Text="Step 1 &#x2013; Select Distribution Group" Style="{StaticResource Lbl}" Margin="0,0,0,4"/>
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <TextBox x:Name="RM_GrpSearch" Style="{StaticResource TB}"/>
                <Button x:Name="BtnRMGrpSearch" Grid.Column="1" Content="Search"
                        Style="{StaticResource SmallBtn}" Margin="8,0,0,0"/>
              </Grid>
              <ListBox x:Name="RM_GrpList" Style="{StaticResource LB}" Height="100" Margin="0,6,0,0"/>

              <TextBlock Text="Step 2 &#x2013; Load and Select Members to Remove" Style="{StaticResource Lbl}" Margin="0,16,0,4"/>
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

          <!-- Panel: Read Properties -->
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

          <!-- Panel: Discover All DGs -->
          <Grid x:Name="PanelDiscover" Visibility="Collapsed" Margin="30,26,30,20">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <StackPanel Grid.Row="0">
              <TextBlock Text="Discover All Distribution Groups" FontSize="22" FontWeight="Bold" Foreground="#0F2B50"/>
              <TextBlock Text="Retrieve and export all mail-enabled distribution groups in your Microsoft 365 tenant."
                         FontSize="12" Foreground="#667788" Margin="0,4,0,18"/>
              <StackPanel Orientation="Horizontal">
                <Button x:Name="BtnDiscoverAll" Content="  &#x1F50D;  Discover All  " Style="{StaticResource ActionBtn}"/>
                <Button x:Name="BtnExportCSV"   Content="  &#x1F4BE;  Export to CSV  " Style="{StaticResource SmallBtn}"
                        Margin="10,0,0,0" IsEnabled="False"/>
              </StackPanel>
              <TextBlock x:Name="Disc_Status" Visibility="Collapsed" FontSize="13" Margin="0,10,0,0"/>
            </StackPanel>

            <Border Grid.Row="2" Margin="0,14,0,0" Background="White" CornerRadius="7"
                    BorderBrush="#D0DCE8" BorderThickness="1">
              <ListView x:Name="DG_ResultList" BorderThickness="0" FontSize="12">
                <ListView.View>
                  <GridView>
                    <GridViewColumn Header="Display Name"    Width="210" DisplayMemberBinding="{Binding DisplayName}"/>
                    <GridViewColumn Header="Email Address"   Width="210" DisplayMemberBinding="{Binding Mail}"/>
                    <GridViewColumn Header="Mail Nickname"   Width="140" DisplayMemberBinding="{Binding MailNickname}"/>
                    <GridViewColumn Header="Description"     Width="200" DisplayMemberBinding="{Binding Description}"/>
                    <GridViewColumn Header="Security Enabled" Width="110" DisplayMemberBinding="{Binding SecurityEnabled}"/>
                  </GridView>
                </ListView.View>
              </ListView>
            </Border>
          </Grid>

        </Grid>
      </Grid>

      <!-- ─────────────────── VIEW: SHARED MAILBOX ─────────────────── -->
      <Grid x:Name="ViewSM" Visibility="Collapsed">
        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
          <TextBlock Text="&#x1F4EC;" FontSize="64" HorizontalAlignment="Center" Margin="0,0,0,16"/>
          <TextBlock Text="Shared Mailbox Management" FontSize="22" FontWeight="Bold"
                     Foreground="#0F2B50" HorizontalAlignment="Center"/>
          <TextBlock Text="This feature is coming in a future version of Migraze."
                     FontSize="14" Foreground="#667788" HorizontalAlignment="Center" Margin="0,10,0,4"/>
          <TextBlock Text="Planned: Create, configure, delegate access, and manage shared mailboxes."
                     FontSize="12" Foreground="#889AAA" HorizontalAlignment="Center"
                     TextAlignment="Center" Margin="40,6,40,0"/>
        </StackPanel>
      </Grid>

      <!-- ─────────────────── VIEW: USER MAILBOX ─────────────────── -->
      <Grid x:Name="ViewUM" Visibility="Collapsed">
        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
          <TextBlock Text="&#x1F464;" FontSize="64" HorizontalAlignment="Center" Margin="0,0,0,16"/>
          <TextBlock Text="User Mailbox Management" FontSize="22" FontWeight="Bold"
                     Foreground="#0F2B50" HorizontalAlignment="Center"/>
          <TextBlock Text="This feature is coming in a future version of Migraze."
                     FontSize="14" Foreground="#667788" HorizontalAlignment="Center" Margin="0,10,0,4"/>
          <TextBlock Text="Planned: Manage mailbox settings, forwarding, quotas, permissions, and more."
                     FontSize="12" Foreground="#889AAA" HorizontalAlignment="Center"
                     TextAlignment="Center" Margin="40,6,40,0"/>
        </StackPanel>
      </Grid>

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
    $script:viewHome   = $window.FindName("ViewHome")
    $script:viewM365   = $window.FindName("ViewM365")
    $script:viewDG     = $window.FindName("ViewDG")
    $script:viewSM     = $window.FindName("ViewSM")
    $script:viewUM     = $window.FindName("ViewUM")
    $script:btnBack    = $window.FindName("BtnBack")
    $script:headerSub  = $window.FindName("HeaderSubtitle")
    $script:connBadge  = $window.FindName("ConnStatusBadge")
    $script:connDot    = $window.FindName("ConnDot")
    $script:connText   = $window.FindName("ConnStatusText")
    $script:tenantInfo = $window.FindName("TenantInfoText")
    $script:btnConnect = $window.FindName("BtnConnect")
    $script:btnConnText= $window.FindName("BtnConnectText")
    $script:footerText = $window.FindName("FooterText")
    $script:btnDG      = $window.FindName("BtnDG")
    $script:btnSM      = $window.FindName("BtnSM")
    $script:btnUM      = $window.FindName("BtnUM")

    $script:allViews = @($script:viewHome, $script:viewM365, $script:viewDG, $script:viewSM, $script:viewUM)

    $script:CurrentView = "Home"

    # ── Navigation helpers ─────────────────────────────────────────────────────
    function script:Show-HomeView {
        $script:allViews | ForEach-Object { $_.Visibility = "Collapsed" }
        $script:viewHome.Visibility  = "Visible"
        $script:btnBack.Visibility   = "Collapsed"
        $script:connBadge.Visibility = "Collapsed"
        $script:headerSub.Text       = "Cloud Management Platform"
        $script:footerText.Text      = "Migraze v2.0  |  Ready"
        $script:CurrentView   = "Home"
    }

    function script:Show-M365View {
        $script:allViews | ForEach-Object { $_.Visibility = "Collapsed" }
        $script:viewM365.Visibility  = "Visible"
        $script:btnBack.Visibility   = "Visible"
        $script:connBadge.Visibility = "Visible"
        $script:headerSub.Text       = "Microsoft 365"
        $script:CurrentView   = "M365"
        Update-M365ConnStatus
        Write-MigrazeLog "Microsoft 365 management opened." "Info"
    }

    function script:Show-DGView {
        $script:allViews | ForEach-Object { $_.Visibility = "Collapsed" }
        $script:viewDG.Visibility    = "Visible"
        $script:btnBack.Visibility   = "Visible"
        $script:connBadge.Visibility = "Visible"
        $script:headerSub.Text       = "Distribution Groups"
        $script:CurrentView   = "DG"
        Update-M365ConnStatus
        Write-MigrazeLog "Distribution Groups management opened." "Info"
    }

    function script:Show-SMView {
        $script:allViews | ForEach-Object { $_.Visibility = "Collapsed" }
        $script:viewSM.Visibility    = "Visible"
        $script:btnBack.Visibility   = "Visible"
        $script:connBadge.Visibility = "Visible"
        $script:headerSub.Text       = "Shared Mailbox"
        $script:CurrentView   = "SM"
        Update-M365ConnStatus
        Write-MigrazeLog "Shared Mailbox management opened." "Info"
    }

    function script:Show-UMView {
        $script:allViews | ForEach-Object { $_.Visibility = "Collapsed" }
        $script:viewUM.Visibility    = "Visible"
        $script:btnBack.Visibility   = "Visible"
        $script:connBadge.Visibility = "Visible"
        $script:headerSub.Text       = "User Mailbox"
        $script:CurrentView   = "UM"
        Update-M365ConnStatus
        Write-MigrazeLog "User Mailbox management opened." "Info"
    }

    function script:Update-M365ConnStatus {
        $connected = Get-MigrazeConnectionStatus
        if ($connected.Connected) {
            $script:connDot.Fill         = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#00C853"))
            $script:connText.Text        = "Connected"
            $script:connText.Foreground  = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#AAFFCC"))
            $script:connBadge.Background = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0C2A10"))
            $script:btnConnText.Text     = "Disconnect"
            $script:btnConnect.Background= [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#C62828"))
            $infoMsg = if ($connected.Account) { "Connected as: $($connected.Account)" } elseif ($script:GraphAccount) { "Connected as: $($script:GraphAccount)" } else { "Connected to Microsoft 365." }
            $script:tenantInfo.Text      = $infoMsg
            $acctLabel = if ($connected.Account) { $connected.Account } elseif ($script:GraphAccount) { $script:GraphAccount } else { "M365" }
            $script:footerText.Text      = "Migraze v2.0  |  Microsoft 365  |  $acctLabel"
            $script:btnDG.IsEnabled = $true; $script:btnSM.IsEnabled = $true; $script:btnUM.IsEnabled = $true
        } else {
            $script:connDot.Fill         = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#FF5252"))
            $script:connText.Text        = "Not Connected"
            $script:connText.Foreground  = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#FFAAAA"))
            $script:connBadge.Background = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#2A1010"))
            $script:btnConnText.Text     = "Connect to M365 Tenant"
            $script:btnConnect.Background= [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0078D4"))
            $script:tenantInfo.Text      = "Connect to your Microsoft 365 tenant to manage Exchange Online objects."
            $script:footerText.Text      = "Migraze v2.0  |  Microsoft 365  |  Not Connected"
            $script:btnDG.IsEnabled = $false; $script:btnSM.IsEnabled = $false; $script:btnUM.IsEnabled = $false
        }
    }

    # ── Event handlers ─────────────────────────────────────────────────────────
    $window.FindName("BtnM365").Add_Click({ Show-M365View })

    $script:btnBack.Add_Click({
        switch ($script:CurrentView) {
            "M365" { Show-HomeView }
            "DG"   { Show-M365View }
            "SM"   { Show-M365View }
            "UM"   { Show-M365View }
        }
    })

    $script:btnConnect.Add_Click({
        if ((Get-MigrazeConnectionStatus).Connected) {
            Write-MigrazeLog "Disconnecting from Microsoft 365..." "Action"
            Disconnect-MigrazeGraph
        } else {
            Write-MigrazeLog "Opening browser for Microsoft 365 login..." "Action"
            $ok = Connect-MigrazeGraph
            if ($ok) { Write-MigrazeLog "Successfully connected to Microsoft 365." "Success" }
        }
        Update-M365ConnStatus
    })

    $script:btnDG.Add_Click({ Show-DGView })
    $script:btnSM.Add_Click({ Show-SMView })
    $script:btnUM.Add_Click({ Show-UMView })

    # ── Initialize DG view event handlers ─────────────────────────────────────
    Initialize-DGView $window

    # ── Start ──────────────────────────────────────────────────────────────────
    Write-MigrazeLog "Migraze v2.0 started." "Action"
    Write-MigrazeLog "Select an environment to get started." "Info"

    $window.ShowDialog() | Out-Null
}
