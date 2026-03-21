# DistributionGroups.ps1 - Distribution Group management (embedded in MainWindow)

function Initialize-DGView {
    param([System.Windows.Window]$window)

    # ── Resolve controls ──────────────────────────────────────────────────────
    $navCreate        = $window.FindName("NavCreate")
    $navUpdate        = $window.FindName("NavUpdate")
    $navAddMembers    = $window.FindName("NavAddMembers")
    $navRemoveMembers = $window.FindName("NavRemoveMembers")
    $navReadProps     = $window.FindName("NavReadProps")
    $navDiscover      = $window.FindName("NavDiscover")

    $pCreate        = $window.FindName("PanelCreate")
    $pUpdate        = $window.FindName("PanelUpdate")
    $pAddMembers    = $window.FindName("PanelAddMembers")
    $pRemoveMembers = $window.FindName("PanelRemoveMembers")
    $pReadProps     = $window.FindName("PanelReadProps")
    $pDiscover      = $window.FindName("PanelDiscover")

    $allPanels  = @($pCreate, $pUpdate, $pAddMembers, $pRemoveMembers, $pReadProps, $pDiscover)
    $allNavBtns = @($navCreate, $navUpdate, $navAddMembers, $navRemoveMembers, $navReadProps, $navDiscover)

    # ── Nav panel switching ───────────────────────────────────────────────────
    $activeNavColor   = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0078D4"))
    $inactiveNavColor = [Windows.Media.Brushes]::Transparent

    function Switch-DGPanel {
        param([int]$Idx)
        for ($i = 0; $i -lt $allPanels.Count; $i++) {
            $allPanels[$i].Visibility  = if ($i -eq $Idx) { "Visible" } else { "Collapsed" }
            $allNavBtns[$i].Background = if ($i -eq $Idx) { $activeNavColor } else { $inactiveNavColor }
            $allNavBtns[$i].FontWeight = if ($i -eq $Idx) { "SemiBold" } else { "Normal" }
        }
    }
    Switch-DGPanel 0

    $navCreate.Add_Click({        Switch-DGPanel 0 })
    $navUpdate.Add_Click({        Switch-DGPanel 1 })
    $navAddMembers.Add_Click({    Switch-DGPanel 2 })
    $navRemoveMembers.Add_Click({ Switch-DGPanel 3 })
    $navReadProps.Add_Click({     Switch-DGPanel 4 })
    $navDiscover.Add_Click({      Switch-DGPanel 5 })

    $window.FindName("DGBtnBack").Add_Click({ Show-M365View })

    # ── Helper: set status text with colour ───────────────────────────────────
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

    # ── Helper: populate a ListBox with group objects ─────────────────────────
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

        Set-StatusText $cStatus "Creating distribution group..." "info"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $btnCreate.IsEnabled = $false

        $result = New-DGGroup -DisplayName $dn -MailNickname $mn -Description $desc -SecurityEnabled $sec

        $window.Cursor = $null
        $btnCreate.IsEnabled = $true

        if ($result.Success) {
            Set-StatusText $cStatus "OK  Distribution group '$dn' created successfully!" "success"
            $cDisplayName.Clear(); $cMailNickname.Clear(); $cDescription.Clear()
            $cSecurityEnabled.IsChecked = $false
        } else {
            Set-StatusText $cStatus "Error: $($result.Error)" "error"
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
            Set-StatusText $uStatus "Properties updated successfully." "success"
        } else {
            Set-StatusText $uStatus "Error: $($result.Error)" "error"
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
        $result = Search-MigrazeUsers -Query $q
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
            Set-StatusText $amStatus "$($usr.DisplayName) added to $($grp.DisplayName) successfully." "success"
        } else {
            Set-StatusText $amStatus "Error: $($result.Error)" "error"
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
            Set-StatusText $rmStatus "Failed to load members: $($result.Error)" "error"
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
            Set-StatusText $rmStatus "$($selected.Count) member(s) removed successfully." "success"
            $btnRMLoadMbrs.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
        } else {
            Set-StatusText $rmStatus "Some removals failed: $($errors -join ', ')" "error"
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

        $rPropsContent.Children.Clear()
        $props = [ordered]@{
            "Display Name"    = $g.DisplayName
            "Email Address"   = $g.Mail
            "Mail Nickname"   = $g.MailNickname
            "Description"     = if ($g.Description) { $g.Description } else { "(none)" }
            "Mail Enabled"    = $g.MailEnabled
            "Security Enabled"= $g.SecurityEnabled
            "Group ID"        = $g.Id
            "Created"         = if ($g.CreatedDateTime) { $g.CreatedDateTime.ToString("yyyy-MM-dd HH:mm UTC") } else { "---" }
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

        $rMbrList.Items.Clear()
        $rMbrHeader.Text = "Members ($($result.Members.Count))"
        foreach ($m in $result.Members) {
            $dn  = if ($m.AdditionalProperties.displayName) { $m.AdditionalProperties.displayName } else { $m.Id }
            $upn = if ($m.AdditionalProperties.userPrincipalName) { " ($($m.AdditionalProperties.userPrincipalName))" } else { "" }
            $rMbrList.Items.Add("$dn$upn") | Out-Null
        }
        $rMembersBox.Visibility = "Visible"
    })

    # ─────────────────────────────────────────────────────
    # ── DISCOVER ALL DGs ──
    # ─────────────────────────────────────────────────────
    $discStatus    = $window.FindName("Disc_Status")
    $dgResultList  = $window.FindName("DG_ResultList")
    $btnDiscoverAll= $window.FindName("BtnDiscoverAll")
    $btnExportCSV  = $window.FindName("BtnExportCSV")
    $script:DiscoveredDGs = @()

    $btnDiscoverAll.Add_Click({
        Set-StatusText $discStatus "Discovering all distribution groups... please wait." "info"
        $window.Cursor     = [System.Windows.Input.Cursors]::Wait
        $btnDiscoverAll.IsEnabled = $false
        $btnExportCSV.IsEnabled   = $false
        $dgResultList.Items.Clear()

        $result = Get-AllDGsForDiscovery
        $window.Cursor     = $null
        $btnDiscoverAll.IsEnabled = $true

        if ($result.Success) {
            $script:DiscoveredDGs = $result.Groups
            foreach ($g in $result.Groups) {
                $item = [PSCustomObject]@{
                    DisplayName     = $g.DisplayName
                    Mail            = if ($g.Mail) { $g.Mail } else { "" }
                    MailNickname    = if ($g.MailNickname) { $g.MailNickname } else { "" }
                    Description     = if ($g.Description) { $g.Description } else { "" }
                    SecurityEnabled = $g.SecurityEnabled
                }
                $dgResultList.Items.Add($item) | Out-Null
            }
            Set-StatusText $discStatus "Found $($result.Groups.Count) distribution group(s)." "success"
            $btnExportCSV.IsEnabled = ($result.Groups.Count -gt 0)
        } else {
            Set-StatusText $discStatus "Discovery failed: $($result.Error)" "error"
        }
    })

    $btnExportCSV.Add_Click({
        if ($script:DiscoveredDGs.Count -eq 0) { return }
        $dlg = [Microsoft.Win32.SaveFileDialog]::new()
        $dlg.Title      = "Export Distribution Groups to CSV"
        $dlg.Filter     = "CSV Files (*.csv)|*.csv"
        $dlg.FileName   = "DistributionGroups_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($dlg.ShowDialog() -eq $true) {
            try {
                $script:DiscoveredDGs | Select-Object DisplayName, Mail, MailNickname, Description, SecurityEnabled |
                    Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
                Write-MigrazeLog "Exported $($script:DiscoveredDGs.Count) DGs to $($dlg.FileName)" "Success"
                Set-StatusText $discStatus "Exported to $($dlg.FileName)" "success"
            } catch {
                Set-StatusText $discStatus "Export failed: $($_.Exception.Message)" "error"
            }
        }
    })
}
