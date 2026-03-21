# DistributionGroups.ps1 - Distribution Group management (embedded in MainWindow)

function Initialize-DGView {
    param([System.Windows.Window]$window)

    # ── Resolve controls ──────────────────────────────────────────────────────
    $navCreate        = $script:mainWindow.FindName("NavCreate")
    $navUpdate        = $script:mainWindow.FindName("NavUpdate")
    $navAddMembers    = $script:mainWindow.FindName("NavAddMembers")
    $navRemoveMembers = $script:mainWindow.FindName("NavRemoveMembers")
    $navReadProps     = $script:mainWindow.FindName("NavReadProps")
    $navDiscover      = $script:mainWindow.FindName("NavDiscover")

    $pCreate        = $script:mainWindow.FindName("PanelCreate")
    $pUpdate        = $script:mainWindow.FindName("PanelUpdate")
    $pAddMembers    = $script:mainWindow.FindName("PanelAddMembers")
    $pRemoveMembers = $script:mainWindow.FindName("PanelRemoveMembers")
    $pReadProps     = $script:mainWindow.FindName("PanelReadProps")
    $pDiscover      = $script:mainWindow.FindName("PanelDiscover")

    # ── Nav panel switching ───────────────────────────────────────────────────
    $script:dgActiveNavColor   = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0078D4"))
    $script:dgInactiveNavColor = [Windows.Media.Brushes]::Transparent
    $script:dgAllPanels  = @($pCreate, $pUpdate, $pAddMembers, $pRemoveMembers, $pReadProps, $pDiscover)
    $script:dgAllNavBtns = @($navCreate, $navUpdate, $navAddMembers, $navRemoveMembers, $navReadProps, $navDiscover)

    function script:Switch-DGPanel {
        param([int]$Idx)
        for ($i = 0; $i -lt $script:dgAllPanels.Count; $i++) {
            $script:dgAllPanels[$i].Visibility  = if ($i -eq $Idx) { "Visible" } else { "Collapsed" }
            $script:dgAllNavBtns[$i].Background = if ($i -eq $Idx) { $script:dgActiveNavColor } else { $script:dgInactiveNavColor }
            $script:dgAllNavBtns[$i].FontWeight = if ($i -eq $Idx) { "SemiBold" } else { "Normal" }
        }
    }
    Switch-DGPanel 0

    $navCreate.Add_Click({        Switch-DGPanel 0 })
    $navUpdate.Add_Click({        Switch-DGPanel 1 })
    $navAddMembers.Add_Click({    Switch-DGPanel 2 })
    $navRemoveMembers.Add_Click({ Switch-DGPanel 3 })
    $navReadProps.Add_Click({     Switch-DGPanel 4 })
    $navDiscover.Add_Click({      Switch-DGPanel 5 })

    $script:mainWindow.FindName("DGBtnBack").Add_Click({ Show-M365View })

    # ── Helper: set status text with colour ───────────────────────────────────
    function script:Set-DGStatusText {
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
    function script:Fill-DGGroupList {
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
    $script:cDisplayName    = $script:mainWindow.FindName("C_DisplayName")
    $script:cMailNickname   = $script:mainWindow.FindName("C_MailNickname")
    $script:cDescription    = $script:mainWindow.FindName("C_Description")
    $script:cSecurityEnabled= $script:mainWindow.FindName("C_SecurityEnabled")
    $script:cStatus         = $script:mainWindow.FindName("C_Status")
    $script:btnCreate       = $script:mainWindow.FindName("BtnCreate")

    $script:btnCreate.Add_Click({
        $dn  = $script:cDisplayName.Text.Trim()
        $mn  = $script:cMailNickname.Text.Trim()
        $desc= $script:cDescription.Text.Trim()
        $sec = $script:cSecurityEnabled.IsChecked -eq $true

        if (-not $dn) { Set-DGStatusText $script:cStatus "Display Name is required." "error"; return }
        if (-not $mn) { Set-DGStatusText $script:cStatus "Email Alias (MailNickname) is required." "error"; return }
        if ($mn -match '\s|[^a-zA-Z0-9._\-]') {
            Set-DGStatusText $script:cStatus "MailNickname must not contain spaces or special characters." "error"; return
        }

        Set-DGStatusText $script:cStatus "Creating distribution group..." "info"
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnCreate.IsEnabled = $false

        $result = New-DGGroup -DisplayName $dn -MailNickname $mn -Description $desc -SecurityEnabled $sec

        $script:mainWindow.Cursor = $null
        $script:btnCreate.IsEnabled = $true

        if ($result.Success) {
            Set-DGStatusText $script:cStatus "OK  Distribution group '$dn' created successfully!" "success"
            $script:cDisplayName.Clear(); $script:cMailNickname.Clear(); $script:cDescription.Clear()
            $script:cSecurityEnabled.IsChecked = $false
        } else {
            Set-DGStatusText $script:cStatus "Error: $($result.Error)" "error"
        }
    })

    # ─────────────────────────────────────────────────────
    # ── UPDATE DG PROPERTIES ──
    # ─────────────────────────────────────────────────────
    $script:uSearch       = $script:mainWindow.FindName("U_Search")
    $script:uDGList       = $script:mainWindow.FindName("U_DGList")
    $script:uFieldsPanel  = $script:mainWindow.FindName("U_FieldsPanel")
    $script:uDisplayName  = $script:mainWindow.FindName("U_DisplayName")
    $script:uDescription  = $script:mainWindow.FindName("U_Description")
    $script:uStatus       = $script:mainWindow.FindName("U_Status")
    $script:btnUSearch    = $script:mainWindow.FindName("BtnUSearch")
    $script:btnUSave      = $script:mainWindow.FindName("BtnUSave")

    $script:btnUSearch.Add_Click({
        $q = $script:uSearch.Text.Trim()
        if (-not $q) { return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGList -SearchQuery $q
        $script:mainWindow.Cursor = $null
        if ($result.Success) {
            Fill-DGGroupList $script:uDGList $result.Groups
            if ($result.Groups.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No distribution groups found for '$q'.", "No Results",
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
            }
        } else {
            [System.Windows.MessageBox]::Show("Search failed:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    $script:uDGList.Add_SelectionChanged({
        if ($script:uDGList.SelectedItem) {
            $script:uDisplayName.Text  = $script:uDGList.SelectedItem.DisplayName
            $script:uFieldsPanel.Visibility = "Visible"
            $script:uStatus.Visibility = "Collapsed"
        }
    })

    $script:btnUSave.Add_Click({
        $selected = $script:uDGList.SelectedItem
        if (-not $selected) { Set-DGStatusText $script:uStatus "Please select a group first." "error"; return }
        $newDN   = $script:uDisplayName.Text.Trim()
        $newDesc = $script:uDescription.Text.Trim()
        if (-not $newDN) { Set-DGStatusText $script:uStatus "Display Name cannot be empty." "error"; return }

        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnUSave.IsEnabled = $false
        $result = Update-DGGroup -GroupId $selected.Id -DisplayName $newDN -Description $newDesc
        $script:mainWindow.Cursor = $null
        $script:btnUSave.IsEnabled = $true

        if ($result.Success) {
            Set-DGStatusText $script:uStatus "Properties updated successfully." "success"
        } else {
            Set-DGStatusText $script:uStatus "Error: $($result.Error)" "error"
        }
    })

    # ─────────────────────────────────────────────────────
    # ── ADD MEMBERS ──
    # ─────────────────────────────────────────────────────
    $script:amGrpSearch  = $script:mainWindow.FindName("AM_GrpSearch")
    $script:amGrpList    = $script:mainWindow.FindName("AM_GrpList")
    $script:amUsrSearch  = $script:mainWindow.FindName("AM_UsrSearch")
    $script:amUsrList    = $script:mainWindow.FindName("AM_UsrList")
    $script:amStatus     = $script:mainWindow.FindName("AM_Status")
    $script:btnAMGrpSrc  = $script:mainWindow.FindName("BtnAMGrpSearch")
    $script:btnAMUsrSrc  = $script:mainWindow.FindName("BtnAMUsrSearch")
    $script:btnAMAdd     = $script:mainWindow.FindName("BtnAMAdd")

    $script:btnAMGrpSrc.Add_Click({
        $q = $script:amGrpSearch.Text.Trim()
        if (-not $q) { return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGList -SearchQuery $q
        $script:mainWindow.Cursor = $null
        if ($result.Success) { Fill-DGGroupList $script:amGrpList $result.Groups }
        else {
            [System.Windows.MessageBox]::Show("Search failed:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    $script:btnAMUsrSrc.Add_Click({
        $q = $script:amUsrSearch.Text.Trim()
        if (-not $q) { return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Search-MigrazeUsers -Query $q
        $script:mainWindow.Cursor = $null
        $script:amUsrList.Items.Clear()
        if ($result.Success) {
            foreach ($u in $result.Users) {
                $item = [PSCustomObject]@{
                    Id          = $u.Id
                    DisplayName = $u.DisplayName
                    UPN         = $u.UserPrincipalName
                    ToString    = "$($u.DisplayName)  ($($u.UserPrincipalName))"
                }
                $script:amUsrList.Items.Add($item) | Out-Null
            }
            $script:amUsrList.DisplayMemberPath = "ToString"
        } else {
            [System.Windows.MessageBox]::Show("User search failed:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    $script:amGrpList.Add_SelectionChanged({ $script:btnAMAdd.IsEnabled = ($script:amGrpList.SelectedItem -and $script:amUsrList.SelectedItem) })
    $script:amUsrList.Add_SelectionChanged({ $script:btnAMAdd.IsEnabled = ($script:amGrpList.SelectedItem -and $script:amUsrList.SelectedItem) })

    $script:btnAMAdd.Add_Click({
        $grp = $script:amGrpList.SelectedItem
        $usr = $script:amUsrList.SelectedItem
        if (-not $grp -or -not $usr) { Set-DGStatusText $script:amStatus "Select both a group and a user." "error"; return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnAMAdd.IsEnabled = $false
        $result = Add-DGMember -GroupId $grp.Id -UserId $usr.UPN
        $script:mainWindow.Cursor = $null
        $script:btnAMAdd.IsEnabled = $true
        if ($result.Success) {
            Set-DGStatusText $script:amStatus "$($usr.DisplayName) added to $($grp.DisplayName) successfully." "success"
        } else {
            Set-DGStatusText $script:amStatus "Error: $($result.Error)" "error"
        }
    })

    # ─────────────────────────────────────────────────────
    # ── REMOVE MEMBERS ──
    # ─────────────────────────────────────────────────────
    $script:rmGrpSearch    = $script:mainWindow.FindName("RM_GrpSearch")
    $script:rmGrpList      = $script:mainWindow.FindName("RM_GrpList")
    $script:rmMbrList      = $script:mainWindow.FindName("RM_MbrList")
    $script:rmStatus       = $script:mainWindow.FindName("RM_Status")
    $script:btnRMGrpSrc    = $script:mainWindow.FindName("BtnRMGrpSearch")
    $script:btnRMLoadMbrs  = $script:mainWindow.FindName("BtnRMLoadMembers")
    $script:btnRMRemove    = $script:mainWindow.FindName("BtnRMRemove")

    $script:btnRMGrpSrc.Add_Click({
        $q = $script:rmGrpSearch.Text.Trim()
        if (-not $q) { return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGList -SearchQuery $q
        $script:mainWindow.Cursor = $null
        if ($result.Success) {
            Fill-DGGroupList $script:rmGrpList $result.Groups
            $script:btnRMLoadMbrs.IsEnabled = ($script:rmGrpList.Items.Count -gt 0)
        } else {
            [System.Windows.MessageBox]::Show("Search failed:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    $script:rmGrpList.Add_SelectionChanged({ $script:btnRMLoadMbrs.IsEnabled = ($null -ne $script:rmGrpList.SelectedItem) })

    $script:btnRMLoadMbrs.Add_Click({
        $grp = $script:rmGrpList.SelectedItem
        if (-not $grp) { return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGProperties -GroupId $grp.Id
        $script:mainWindow.Cursor = $null
        $script:rmMbrList.Items.Clear()
        if ($result.Success) {
            foreach ($m in $result.Members) {
                $item = [PSCustomObject]@{
                    Id          = $m.Id
                    DisplayName = $m.DisplayName
                    ToString    = $m.ToString
                }
                $script:rmMbrList.Items.Add($item) | Out-Null
            }
            $script:rmMbrList.DisplayMemberPath = "ToString"
            Set-DGStatusText $script:rmStatus "$($result.Members.Count) member(s) loaded." "info"
        } else {
            Set-DGStatusText $script:rmStatus "Failed to load members: $($result.Error)" "error"
        }
    })

    $script:rmMbrList.Add_SelectionChanged({ $script:btnRMRemove.IsEnabled = ($script:rmMbrList.SelectedItems.Count -gt 0) })

    $script:btnRMRemove.Add_Click({
        $grp = $script:rmGrpList.SelectedItem
        if (-not $grp) { Set-DGStatusText $script:rmStatus "No group selected." "error"; return }
        $selected = @($script:rmMbrList.SelectedItems)
        if ($selected.Count -eq 0) { Set-DGStatusText $script:rmStatus "No member selected." "error"; return }

        $confirm = [System.Windows.MessageBox]::Show(
            "Remove $($selected.Count) member(s) from '$($grp.DisplayName)'?",
            "Confirm Removal",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnRMRemove.IsEnabled = $false
        $errors = @()
        foreach ($mbr in $selected) {
            $res = Remove-DGMember -GroupId $grp.Id -MemberId $mbr.Id
            if (-not $res.Success) { $errors += $mbr.DisplayName }
        }
        $script:mainWindow.Cursor = $null
        $script:btnRMRemove.IsEnabled = $true

        if ($errors.Count -eq 0) {
            Set-DGStatusText $script:rmStatus "$($selected.Count) member(s) removed successfully." "success"
            $script:btnRMLoadMbrs.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
        } else {
            Set-DGStatusText $script:rmStatus "Some removals failed: $($errors -join ', ')" "error"
        }
    })

    # ─────────────────────────────────────────────────────
    # ── READ PROPERTIES ──
    # ─────────────────────────────────────────────────────
    $script:rSearch      = $script:mainWindow.FindName("R_Search")
    $script:rDGList      = $script:mainWindow.FindName("R_DGList")
    $script:rPropsBox    = $script:mainWindow.FindName("R_PropsBox")
    $script:rPropsContent= $script:mainWindow.FindName("R_PropsContent")
    $script:rMembersBox  = $script:mainWindow.FindName("R_MembersBox")
    $script:rMbrList     = $script:mainWindow.FindName("R_MbrList")
    $script:rMbrHeader   = $script:mainWindow.FindName("R_MemberHeader")
    $script:btnRSearch   = $script:mainWindow.FindName("BtnRSearch")
    $script:btnRLoad     = $script:mainWindow.FindName("BtnRLoad")

    $script:btnRSearch.Add_Click({
        $q = $script:rSearch.Text.Trim()
        if (-not $q) { return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGList -SearchQuery $q
        $script:mainWindow.Cursor = $null
        if ($result.Success) {
            Fill-DGGroupList $script:rDGList $result.Groups
            $script:btnRLoad.IsEnabled = ($script:rDGList.Items.Count -gt 0)
        } else {
            [System.Windows.MessageBox]::Show("Search failed:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    $script:rDGList.Add_SelectionChanged({ $script:btnRLoad.IsEnabled = ($null -ne $script:rDGList.SelectedItem) })

    $script:btnRLoad.Add_Click({
        $grp = $script:rDGList.SelectedItem
        if (-not $grp) { return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGProperties -GroupId $grp.Id
        $script:mainWindow.Cursor = $null
        if (-not $result.Success) {
            [System.Windows.MessageBox]::Show("Failed to load properties:`n$($result.Error)", "Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            return
        }
        $g = $result.Group

        $script:rPropsContent.Children.Clear()
        $props = [ordered]@{
            "Display Name"    = $g.DisplayName
            "Email Address"   = $g.Mail
            "Alias"           = $g.Alias
            "Description"     = if ($g.Description) { $g.Description } else { "(none)" }
            "Security Enabled"= $g.SecurityEnabled
            "Group ID"        = $g.Id
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
            $script:rPropsContent.Children.Add($row) | Out-Null
        }
        $script:rPropsBox.Visibility = "Visible"

        $script:rMbrList.Items.Clear()
        $script:rMbrHeader.Text = "Members ($($result.Members.Count))"
        foreach ($m in $result.Members) {
            $script:rMbrList.Items.Add($m.ToString) | Out-Null
        }
        $script:rMembersBox.Visibility = "Visible"
    })

    # ─────────────────────────────────────────────────────
    # ── DISCOVER ALL DGs ──
    # ─────────────────────────────────────────────────────
    $script:discStatus    = $script:mainWindow.FindName("Disc_Status")
    $script:dgResultList  = $script:mainWindow.FindName("DG_ResultList")
    $script:btnDiscoverAll= $script:mainWindow.FindName("BtnDiscoverAll")
    $script:btnExportCSV  = $script:mainWindow.FindName("BtnExportCSV")
    $script:DiscoveredDGs = @()

    $script:btnDiscoverAll.Add_Click({
        Set-DGStatusText $script:discStatus "Discovering all distribution groups... please wait." "info"
        $script:mainWindow.Cursor     = [System.Windows.Input.Cursors]::Wait
        $script:btnDiscoverAll.IsEnabled = $false
        $script:btnExportCSV.IsEnabled   = $false
        $script:dgResultList.Items.Clear()

        $result = Get-AllDGsForDiscovery
        $script:mainWindow.Cursor     = $null
        $script:btnDiscoverAll.IsEnabled = $true

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
                $script:dgResultList.Items.Add($item) | Out-Null
            }
            Set-DGStatusText $script:discStatus "Found $($result.Groups.Count) distribution group(s)." "success"
            $script:btnExportCSV.IsEnabled = ($result.Groups.Count -gt 0)
        } else {
            Set-DGStatusText $script:discStatus "Discovery failed: $($result.Error)" "error"
        }
    })

    $script:btnExportCSV.Add_Click({
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
                Set-DGStatusText $script:discStatus "Exported to $($dlg.FileName)" "success"
            } catch {
                Set-DGStatusText $script:discStatus "Export failed: $($_.Exception.Message)" "error"
            }
        }
    })
}
