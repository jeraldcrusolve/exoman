# DistributionGroups.ps1 - Distribution Group management (card-based layout)

function Initialize-DGView {
    param([System.Windows.Window]$window)

    # ── Resolve top-level DG panels ──────────────────────────────────────────
    $script:panelDGHome   = $script:mainWindow.FindName("PanelDGHome")
    $script:panelDGSingle = $script:mainWindow.FindName("PanelDGSingle")
    $script:panelDGBulk   = $script:mainWindow.FindName("PanelDGBulk")

    # ── Resolve Single sub-panels ────────────────────────────────────────────
    $script:panelSingleCreate = $script:mainWindow.FindName("PanelSingleCreate")
    $script:panelSingleSearch = $script:mainWindow.FindName("PanelSingleSearch")
    $script:panelSingleUpdate = $script:mainWindow.FindName("PanelSingleUpdate")

    # ── Resolve Bulk sub-panels ──────────────────────────────────────────────
    $script:panelBulkDiscover = $script:mainWindow.FindName("PanelBulkDiscover")
    $script:panelBulkCreate   = $script:mainWindow.FindName("PanelBulkCreate")
    $script:panelBulkUpdate   = $script:mainWindow.FindName("PanelBulkUpdate")

    # ── Resolve sidebar nav buttons ──────────────────────────────────────────
    $script:navSingleCreate = $script:mainWindow.FindName("NavSingleCreate")
    $script:navSingleSearch = $script:mainWindow.FindName("NavSingleSearch")
    $script:navSingleUpdate = $script:mainWindow.FindName("NavSingleUpdate")
    $script:navBulkDiscover = $script:mainWindow.FindName("NavBulkDiscover")
    $script:navBulkCreate   = $script:mainWindow.FindName("NavBulkCreate")
    $script:navBulkUpdate   = $script:mainWindow.FindName("NavBulkUpdate")

    # ── Colour helpers ───────────────────────────────────────────────────────
    $script:navActive   = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0078D4"))
    $script:navInactive = [Windows.Media.Brushes]::Transparent

    # ═══════════════════════════════════════════════════════════════════════
    # Navigation helpers
    # ═══════════════════════════════════════════════════════════════════════

    function script:Show-DGHomePanel {
        $script:panelDGHome.Visibility   = "Visible"
        $script:panelDGSingle.Visibility = "Collapsed"
        $script:panelDGBulk.Visibility   = "Collapsed"
    }

    function script:Show-DGSinglePanel {
        $script:panelDGHome.Visibility   = "Collapsed"
        $script:panelDGSingle.Visibility = "Visible"
        $script:panelDGBulk.Visibility   = "Collapsed"
        Switch-SinglePanel 0
    }

    function script:Show-DGBulkPanel {
        $script:panelDGHome.Visibility   = "Collapsed"
        $script:panelDGSingle.Visibility = "Collapsed"
        $script:panelDGBulk.Visibility   = "Visible"
        Switch-BulkPanel 0
    }

    function script:Switch-SinglePanel {
        param([int]$Idx)
        $panels = @($script:panelSingleCreate, $script:panelSingleSearch, $script:panelSingleUpdate)
        $navs   = @($script:navSingleCreate,   $script:navSingleSearch,   $script:navSingleUpdate)
        for ($i = 0; $i -lt $panels.Count; $i++) {
            $panels[$i].Visibility  = if ($i -eq $Idx) { "Visible" } else { "Collapsed" }
            $navs[$i].Background    = if ($i -eq $Idx) { $script:navActive } else { $script:navInactive }
            $navs[$i].FontWeight    = if ($i -eq $Idx) { "SemiBold" } else { "Normal" }
        }
    }

    function script:Switch-BulkPanel {
        param([int]$Idx)
        $panels = @($script:panelBulkDiscover, $script:panelBulkCreate, $script:panelBulkUpdate)
        $navs   = @($script:navBulkDiscover,   $script:navBulkCreate,   $script:navBulkUpdate)
        for ($i = 0; $i -lt $panels.Count; $i++) {
            $panels[$i].Visibility  = if ($i -eq $Idx) { "Visible" } else { "Collapsed" }
            $navs[$i].Background    = if ($i -eq $Idx) { $script:navActive } else { $script:navInactive }
            $navs[$i].FontWeight    = if ($i -eq $Idx) { "SemiBold" } else { "Normal" }
        }
    }

    # ── Helper: status text ──────────────────────────────────────────────────
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

    # ── Helper: populate a ListBox with DG group objects ─────────────────────
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

    # ── Start on home panel ───────────────────────────────────────────────────
    Show-DGHomePanel

    # ═══════════════════════════════════════════════════════════════════════
    # Top-level card / back navigation
    # ═══════════════════════════════════════════════════════════════════════
    $cardSingleDG  = $script:mainWindow.FindName("CardSingleDG")
    $cardBulkDG    = $script:mainWindow.FindName("CardBulkDG")
    $dgBtnBack     = $script:mainWindow.FindName("DGBtnBack")
    $btnSingleBack = $script:mainWindow.FindName("BtnSingleBackHome")
    $btnBulkBack   = $script:mainWindow.FindName("BtnBulkBackHome")

    $cardSingleDG.Add_MouseLeftButtonUp({  Show-DGSinglePanel })
    $cardBulkDG.Add_MouseLeftButtonUp({    Show-DGBulkPanel   })
    $dgBtnBack.Add_Click({                 Show-M365View       })
    $btnSingleBack.Add_Click({             Show-DGHomePanel    })
    $btnBulkBack.Add_Click({               Show-DGHomePanel    })

    # ── Single sidebar nav ───────────────────────────────────────────────────
    $script:navSingleCreate.Add_Click({ Switch-SinglePanel 0 })
    $script:navSingleSearch.Add_Click({ Switch-SinglePanel 1 })
    $script:navSingleUpdate.Add_Click({ Switch-SinglePanel 2 })

    # ── Bulk sidebar nav ─────────────────────────────────────────────────────
    $script:navBulkDiscover.Add_Click({ Switch-BulkPanel 0 })
    $script:navBulkCreate.Add_Click({   Switch-BulkPanel 1 })
    $script:navBulkUpdate.Add_Click({   Switch-BulkPanel 2 })

    # ═══════════════════════════════════════════════════════════════════════
    # SINGLE – CREATE DG
    # ═══════════════════════════════════════════════════════════════════════
    $script:cDisplayName     = $script:mainWindow.FindName("C_DisplayName")
    $script:cMailNickname    = $script:mainWindow.FindName("C_MailNickname")
    $script:cSecurityEnabled = $script:mainWindow.FindName("C_SecurityEnabled")
    $script:cStatus          = $script:mainWindow.FindName("C_Status")
    $script:btnCreate        = $script:mainWindow.FindName("BtnCreate")

    $script:btnCreate.Add_Click({
        $dn  = $script:cDisplayName.Text.Trim()
        $mn  = $script:cMailNickname.Text.Trim()
        $sec = $script:cSecurityEnabled.IsChecked -eq $true

        if (-not $dn) { Set-DGStatusText $script:cStatus "Display Name is required." "error"; return }
        if (-not $mn) { Set-DGStatusText $script:cStatus "Email Alias is required." "error"; return }
        if ($mn -match '\s|[^a-zA-Z0-9._\-]') {
            Set-DGStatusText $script:cStatus "Alias must not contain spaces or special characters." "error"; return
        }

        Set-DGStatusText $script:cStatus "Creating distribution group..." "info"
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnCreate.IsEnabled = $false

        $result = New-DGGroup -DisplayName $dn -MailNickname $mn -SecurityEnabled $sec

        $script:mainWindow.Cursor = $null
        $script:btnCreate.IsEnabled = $true

        if ($result.Success) {
            Set-DGStatusText $script:cStatus "Distribution group '$dn' created successfully!" "success"
            $script:cDisplayName.Clear()
            $script:cMailNickname.Clear()
            $script:cSecurityEnabled.IsChecked = $false
            Write-MigrazeLog "Created DG '$dn' ($mn)." "Success"
        } else {
            Set-DGStatusText $script:cStatus "Error: $($result.Error)" "error"
            Write-MigrazeLog "Failed to create DG '$dn': $($result.Error)" "Error"
        }
    })

    # ═══════════════════════════════════════════════════════════════════════
    # SINGLE – SEARCH DG  (with inline member management)
    # ═══════════════════════════════════════════════════════════════════════
    $script:sSearch        = $script:mainWindow.FindName("S_Search")
    $script:sDGList        = $script:mainWindow.FindName("S_DGList")
    $script:sPropsBox      = $script:mainWindow.FindName("S_PropsBox")
    $script:sPropsContent  = $script:mainWindow.FindName("S_PropsContent")
    $script:sMembersBox    = $script:mainWindow.FindName("S_MembersBox")
    $script:sMemberHeader  = $script:mainWindow.FindName("S_MemberHeader")
    $script:sMbrList       = $script:mainWindow.FindName("S_MbrList")
    $script:sBtnRemoveMbr  = $script:mainWindow.FindName("BtnSRemoveMember")
    $script:sAddUsrSearch  = $script:mainWindow.FindName("S_AddUsrSearch")
    $script:sBtnAddUsrSrc  = $script:mainWindow.FindName("BtnSAddUsrSearch")
    $script:sAddUsrList    = $script:mainWindow.FindName("S_AddUsrList")
    $script:sBtnAddMember  = $script:mainWindow.FindName("BtnSAddMember")
    $script:sStatus        = $script:mainWindow.FindName("S_Status")
    $script:btnSSearch     = $script:mainWindow.FindName("BtnSSearch")

    # Search DGs
    $script:btnSSearch.Add_Click({
        $q = $script:sSearch.Text.Trim()
        if (-not $q) { return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGList -SearchQuery $q
        $script:mainWindow.Cursor = $null
        if ($result.Success) {
            Fill-DGGroupList $script:sDGList $result.Groups
            $script:sPropsBox.Visibility    = "Collapsed"
            $script:sMembersBox.Visibility  = "Collapsed"
            $script:sStatus.Visibility      = "Collapsed"
            if ($result.Groups.Count -eq 0) {
                Set-DGStatusText $script:sStatus "No distribution groups found for '$q'." "info"
            }
        } else {
            Set-DGStatusText $script:sStatus "Search failed: $($result.Error)" "error"
        }
    })

    # Load props + members when a DG is selected
    $script:sDGList.Add_SelectionChanged({
        $grp = $script:sDGList.SelectedItem
        if (-not $grp) { return }

        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGProperties -GroupId $grp.Id
        $script:mainWindow.Cursor = $null

        if (-not $result.Success) {
            Set-DGStatusText $script:sStatus "Failed to load properties: $($result.Error)" "error"
            return
        }
        $g = $result.Group

        # Build property rows
        $script:sPropsContent.Children.Clear()
        $props = [ordered]@{
            "Display Name"     = $g.DisplayName
            "Email Address"    = $g.Mail
            "Alias"            = $g.Alias
            "Description"      = if ($g.Description) { $g.Description } else { "(none)" }
            "Security Enabled" = $g.SecurityEnabled
            "Group ID"         = $g.Id
        }
        foreach ($k in $props.Keys) {
            $row  = [System.Windows.Controls.Grid]::new()
            $col1 = [System.Windows.Controls.ColumnDefinition]::new(); $col1.Width = [System.Windows.GridLength]::new(160)
            $col2 = [System.Windows.Controls.ColumnDefinition]::new(); $col2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $row.ColumnDefinitions.Add($col1); $row.ColumnDefinitions.Add($col2)
            $row.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)

            $lbl = [System.Windows.Controls.TextBlock]::new()
            $lbl.Text = $k; $lbl.FontWeight = "SemiBold"
            $lbl.Foreground = [Windows.Media.Brushes]::DimGray; $lbl.FontSize = 12
            [System.Windows.Controls.Grid]::SetColumn($lbl, 0)

            $val = [System.Windows.Controls.TextBlock]::new()
            $val.Text = "$($props[$k])"; $val.FontSize = 12; $val.TextWrapping = "Wrap"
            $val.Foreground = [Windows.Media.Brushes]::Black
            [System.Windows.Controls.Grid]::SetColumn($val, 1)

            $row.Children.Add($lbl) | Out-Null; $row.Children.Add($val) | Out-Null
            $script:sPropsContent.Children.Add($row) | Out-Null
        }
        $script:sPropsBox.Visibility = "Visible"

        # Populate members list
        $script:sMbrList.Items.Clear()
        $script:sMemberHeader.Text = "Members ($($result.Members.Count))"
        foreach ($m in $result.Members) {
            $item = [PSCustomObject]@{
                Id          = $m.Id
                DisplayName = $m.DisplayName
                ToString    = $m.ToString
            }
            $script:sMbrList.Items.Add($item) | Out-Null
        }
        $script:sMbrList.DisplayMemberPath = "ToString"
        $script:sMembersBox.Visibility    = "Visible"
        $script:sBtnRemoveMbr.IsEnabled   = $false
        $script:sStatus.Visibility        = "Collapsed"
    })

    # Enable Remove button when member is selected
    $script:sMbrList.Add_SelectionChanged({
        $script:sBtnRemoveMbr.IsEnabled = ($script:sMbrList.SelectedItems.Count -gt 0)
    })

    # Remove selected member(s)
    $script:sBtnRemoveMbr.Add_Click({
        $grp = $script:sDGList.SelectedItem
        if (-not $grp) { Set-DGStatusText $script:sStatus "No group selected." "error"; return }
        $selected = @($script:sMbrList.SelectedItems)
        if ($selected.Count -eq 0) { Set-DGStatusText $script:sStatus "No member selected." "error"; return }

        $confirm = [System.Windows.MessageBox]::Show(
            "Remove $($selected.Count) member(s) from '$($grp.DisplayName)'?",
            "Confirm Removal",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:sBtnRemoveMbr.IsEnabled = $false
        $errors = @()
        foreach ($mbr in $selected) {
            $res = Remove-DGMember -GroupId $grp.Id -MemberId $mbr.Id
            if (-not $res.Success) { $errors += $mbr.DisplayName }
        }
        $script:mainWindow.Cursor = $null

        if ($errors.Count -eq 0) {
            Set-DGStatusText $script:sStatus "$($selected.Count) member(s) removed." "success"
            Write-MigrazeLog "Removed $($selected.Count) member(s) from '$($grp.DisplayName)'." "Success"
        } else {
            Set-DGStatusText $script:sStatus "Some removals failed: $($errors -join ', ')" "error"
        }
        # Reload member list
        $script:sDGList.RaiseEvent([System.Windows.Controls.SelectionChangedEventArgs]::new(
            [System.Windows.Controls.Primitives.Selector]::SelectionChangedEvent,
            [System.Collections.Generic.List[Object]]::new(),
            [System.Collections.Generic.List[Object]]::new()
        ))
    })

    # Search user to add as member
    $script:sBtnAddUsrSrc.Add_Click({
        $q = $script:sAddUsrSearch.Text.Trim()
        if (-not $q) { return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Search-MigrazeUsers -Query $q
        $script:mainWindow.Cursor = $null
        $script:sAddUsrList.Items.Clear()
        if ($result.Success) {
            foreach ($u in $result.Users) {
                $item = [PSCustomObject]@{
                    Id          = $u.Id
                    DisplayName = $u.DisplayName
                    UPN         = $u.UserPrincipalName
                    ToString    = "$($u.DisplayName)  ($($u.UserPrincipalName))"
                }
                $script:sAddUsrList.Items.Add($item) | Out-Null
            }
            $script:sAddUsrList.DisplayMemberPath = "ToString"
        } else {
            Set-DGStatusText $script:sStatus "User search failed: $($result.Error)" "error"
        }
    })

    # Enable Add button when user is selected
    $script:sAddUsrList.Add_SelectionChanged({
        $script:sBtnAddMember.IsEnabled = ($null -ne $script:sAddUsrList.SelectedItem)
    })

    # Add selected user as member
    $script:sBtnAddMember.Add_Click({
        $grp = $script:sDGList.SelectedItem
        $usr = $script:sAddUsrList.SelectedItem
        if (-not $grp) { Set-DGStatusText $script:sStatus "No group selected." "error"; return }
        if (-not $usr) { Set-DGStatusText $script:sStatus "No user selected." "error"; return }

        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:sBtnAddMember.IsEnabled = $false
        $result = Add-DGMember -GroupId $grp.Id -UserId $usr.UPN
        $script:mainWindow.Cursor = $null
        $script:sBtnAddMember.IsEnabled = $true

        if ($result.Success) {
            Set-DGStatusText $script:sStatus "$($usr.DisplayName) added to '$($grp.DisplayName)'." "success"
            Write-MigrazeLog "Added '$($usr.DisplayName)' to '$($grp.DisplayName)'." "Success"
            # Reload member list
            $script:sDGList.RaiseEvent([System.Windows.Controls.SelectionChangedEventArgs]::new(
                [System.Windows.Controls.Primitives.Selector]::SelectionChangedEvent,
                [System.Collections.Generic.List[Object]]::new(),
                [System.Collections.Generic.List[Object]]::new()
            ))
        } else {
            Set-DGStatusText $script:sStatus "Failed to add member: $($result.Error)" "error"
            Write-MigrazeLog "Failed to add '$($usr.DisplayName)' to '$($grp.DisplayName)': $($result.Error)" "Error"
        }
    })

    # ═══════════════════════════════════════════════════════════════════════
    # SINGLE – UPDATE DG SETTINGS
    # ═══════════════════════════════════════════════════════════════════════
    $script:upSearch      = $script:mainWindow.FindName("UP_Search")
    $script:upDGList      = $script:mainWindow.FindName("UP_DGList")
    $script:upFieldsPanel = $script:mainWindow.FindName("UP_FieldsPanel")
    $script:upDisplayName = $script:mainWindow.FindName("UP_DisplayName")
    $script:upStatus      = $script:mainWindow.FindName("UP_Status")
    $script:btnUPSearch   = $script:mainWindow.FindName("BtnUPSearch")
    $script:btnUPSave     = $script:mainWindow.FindName("BtnUPSave")

    $script:btnUPSearch.Add_Click({
        $q = $script:upSearch.Text.Trim()
        if (-not $q) { return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $result = Get-DGList -SearchQuery $q
        $script:mainWindow.Cursor = $null
        if ($result.Success) {
            Fill-DGGroupList $script:upDGList $result.Groups
            $script:upFieldsPanel.Visibility = "Collapsed"
            $script:upStatus.Visibility      = "Collapsed"
            if ($result.Groups.Count -eq 0) {
                Set-DGStatusText $script:upStatus "No distribution groups found for '$q'." "info"
            }
        } else {
            Set-DGStatusText $script:upStatus "Search failed: $($result.Error)" "error"
        }
    })

    $script:upDGList.Add_SelectionChanged({
        if ($script:upDGList.SelectedItem) {
            $script:upDisplayName.Text       = $script:upDGList.SelectedItem.DisplayName
            $script:upFieldsPanel.Visibility = "Visible"
            $script:upStatus.Visibility      = "Collapsed"
        }
    })

    $script:btnUPSave.Add_Click({
        $selected = $script:upDGList.SelectedItem
        if (-not $selected) { Set-DGStatusText $script:upStatus "Please select a group first." "error"; return }
        $newDN = $script:upDisplayName.Text.Trim()
        if (-not $newDN) { Set-DGStatusText $script:upStatus "Display Name cannot be empty." "error"; return }

        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnUPSave.IsEnabled = $false
        $result = Update-DGGroup -GroupId $selected.Id -DisplayName $newDN
        $script:mainWindow.Cursor = $null
        $script:btnUPSave.IsEnabled = $true

        if ($result.Success) {
            Set-DGStatusText $script:upStatus "Display name updated to '$newDN'." "success"
            Write-MigrazeLog "Updated DG '$($selected.DisplayName)' -> '$newDN'." "Success"
            # Update the list item so it reflects the change
            $selected.DisplayName = $newDN
        } else {
            Set-DGStatusText $script:upStatus "Error: $($result.Error)" "error"
            Write-MigrazeLog "Failed to update DG '$($selected.DisplayName)': $($result.Error)" "Error"
        }
    })

    # ═══════════════════════════════════════════════════════════════════════
    # BULK – DISCOVER ALL DGs
    # ═══════════════════════════════════════════════════════════════════════
    $script:discStatus     = $script:mainWindow.FindName("Disc_Status")
    $script:dgResultList   = $script:mainWindow.FindName("DG_ResultList")
    $script:btnDiscoverAll = $script:mainWindow.FindName("BtnDiscoverAll")
    $script:btnExportCSV   = $script:mainWindow.FindName("BtnExportCSV")
    $script:DiscoveredDGs  = @()

    $script:btnDiscoverAll.Add_Click({
        Set-DGStatusText $script:discStatus "Discovering all distribution groups... please wait." "info"
        $script:mainWindow.Cursor        = [System.Windows.Input.Cursors]::Wait
        $script:btnDiscoverAll.IsEnabled = $false
        $script:btnExportCSV.IsEnabled   = $false
        $script:dgResultList.Items.Clear()

        $result = Get-AllDGsForDiscovery
        $script:mainWindow.Cursor        = $null
        $script:btnDiscoverAll.IsEnabled = $true

        if ($result.Success) {
            $script:DiscoveredDGs = $result.Groups
            foreach ($g in $result.Groups) {
                $item = [PSCustomObject]@{
                    DisplayName     = $g.DisplayName
                    Mail            = if ($g.Mail)         { $g.Mail }         else { "" }
                    MailNickname    = if ($g.MailNickname) { $g.MailNickname } else { "" }
                    SecurityEnabled = $g.SecurityEnabled
                }
                $script:dgResultList.Items.Add($item) | Out-Null
            }
            Set-DGStatusText $script:discStatus "Found $($result.Groups.Count) distribution group(s)." "success"
            $script:btnExportCSV.IsEnabled = ($result.Groups.Count -gt 0)
            Write-MigrazeLog "Discovered $($result.Groups.Count) DG(s)." "Success"
        } else {
            Set-DGStatusText $script:discStatus "Discovery failed: $($result.Error)" "error"
            Write-MigrazeLog "DG discovery failed: $($result.Error)" "Error"
        }
    })

    $script:btnExportCSV.Add_Click({
        if ($script:DiscoveredDGs.Count -eq 0) { return }
        $dlg           = [Microsoft.Win32.SaveFileDialog]::new()
        $dlg.Title     = "Export Distribution Groups"
        $dlg.Filter    = "CSV Files (*.csv)|*.csv"
        $dlg.FileName  = "DistributionGroups_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($dlg.ShowDialog() -eq $true) {
            try {
                $script:DiscoveredDGs | Select-Object DisplayName,
                    @{N='Email';E={$_.Mail}},
                    @{N='Alias';E={$_.MailNickname}},
                    SecurityEnabled |
                    Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
                Set-DGStatusText $script:discStatus "Exported to: $($dlg.FileName)" "success"
                Write-MigrazeLog "Exported $($script:DiscoveredDGs.Count) DG(s) to $($dlg.FileName)." "Success"
            } catch {
                Set-DGStatusText $script:discStatus "Export failed: $_" "error"
            }
        }
    })

    # ═══════════════════════════════════════════════════════════════════════
    # BULK – BULK CREATE DGs
    # ═══════════════════════════════════════════════════════════════════════
    $script:bcCsvText    = $script:mainWindow.FindName("BC_CsvText")
    $script:bcStatus     = $script:mainWindow.FindName("BC_Status")
    $script:bcResultList = $script:mainWindow.FindName("BC_ResultList")
    $script:btnBCBrowse  = $script:mainWindow.FindName("BtnBCBrowse")
    $script:btnBCCreate  = $script:mainWindow.FindName("BtnBCCreate")

    $script:btnBCBrowse.Add_Click({
        $dlg = [Microsoft.Win32.OpenFileDialog]::new()
        $dlg.Title  = "Import CSV File"
        $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        if ($dlg.ShowDialog() -eq $true) {
            try {
                $script:bcCsvText.Text = [System.IO.File]::ReadAllText($dlg.FileName)
            } catch {
                Set-DGStatusText $script:bcStatus "Error loading file: $_" "error"
            }
        }
    })

    $script:btnBCCreate.Add_Click({
        $lines = $script:bcCsvText.Text -split "`n" | Where-Object { $_.Trim() -ne "" }
        if ($lines.Count -eq 0) { Set-DGStatusText $script:bcStatus "No data to process." "error"; return }

        $script:mainWindow.Cursor      = [System.Windows.Input.Cursors]::Wait
        $script:btnBCCreate.IsEnabled  = $false
        $script:bcResultList.Items.Clear()
        Set-DGStatusText $script:bcStatus "Processing $($lines.Count) line(s)..." "info"

        $ok = 0; $fail = 0
        foreach ($line in $lines) {
            $parts = $line.Trim() -split ",", 3
            $dn    = if ($parts.Count -gt 0) { $parts[0].Trim() } else { "" }
            $alias = if ($parts.Count -gt 1) { $parts[1].Trim() } else { "" }
            $type  = if ($parts.Count -gt 2) { $parts[2].Trim() } else { "Distribution" }
            $sec   = ($type -ieq "Security")

            if (-not $dn -or -not $alias) {
                $script:bcResultList.Items.Add("SKIP (missing fields): $line") | Out-Null
                $fail++; continue
            }
            $result = New-DGGroup -DisplayName $dn -MailNickname $alias -SecurityEnabled $sec
            if ($result.Success) {
                $script:bcResultList.Items.Add("OK:   $dn  ($alias)") | Out-Null
                $ok++
            } else {
                $script:bcResultList.Items.Add("FAIL: $dn  -- $($result.Error)") | Out-Null
                $fail++
            }
        }
        $script:mainWindow.Cursor     = $null
        $script:btnBCCreate.IsEnabled = $true
        $statusType = if ($fail -eq 0) { "success" } else { "error" }
        Set-DGStatusText $script:bcStatus "Done: $ok created, $fail failed." $statusType
        Write-MigrazeLog "Bulk create: $ok OK, $fail failed." $(if ($fail -eq 0) { "Success" } else { "Error" })
    })

    # ═══════════════════════════════════════════════════════════════════════
    # BULK – BULK UPDATE DG SETTINGS
    # ═══════════════════════════════════════════════════════════════════════
    $script:buCsvText    = $script:mainWindow.FindName("BU_CsvText")
    $script:buStatus     = $script:mainWindow.FindName("BU_Status")
    $script:buResultList = $script:mainWindow.FindName("BU_ResultList")
    $script:btnBUBrowse  = $script:mainWindow.FindName("BtnBUBrowse")
    $script:btnBUApply   = $script:mainWindow.FindName("BtnBUApply")

    $script:btnBUBrowse.Add_Click({
        $dlg = [Microsoft.Win32.OpenFileDialog]::new()
        $dlg.Title  = "Import CSV File"
        $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        if ($dlg.ShowDialog() -eq $true) {
            try {
                $script:buCsvText.Text = [System.IO.File]::ReadAllText($dlg.FileName)
            } catch {
                Set-DGStatusText $script:buStatus "Error loading file: $_" "error"
            }
        }
    })

    $script:btnBUApply.Add_Click({
        $lines = $script:buCsvText.Text -split "`n" | Where-Object { $_.Trim() -ne "" }
        if ($lines.Count -eq 0) { Set-DGStatusText $script:buStatus "No data to process." "error"; return }

        $script:mainWindow.Cursor     = [System.Windows.Input.Cursors]::Wait
        $script:btnBUApply.IsEnabled  = $false
        $script:buResultList.Items.Clear()
        Set-DGStatusText $script:buStatus "Processing $($lines.Count) line(s)..." "info"

        $ok = 0; $fail = 0
        foreach ($line in $lines) {
            # Format: Identity, NewDisplayName, AddMember(UPN), RemoveMember(UPN)
            $parts     = $line.Trim() -split ",", 4
            $identity  = if ($parts.Count -gt 0) { $parts[0].Trim() } else { "" }
            $newDN     = if ($parts.Count -gt 1) { $parts[1].Trim() } else { "" }
            $addMbr    = if ($parts.Count -gt 2) { $parts[2].Trim() } else { "" }
            $removeMbr = if ($parts.Count -gt 3) { $parts[3].Trim() } else { "" }

            if (-not $identity) {
                $script:buResultList.Items.Add("SKIP (no identity): $line") | Out-Null
                $fail++; continue
            }

            $lineOk = $true
            if ($newDN) {
                $res = Update-DGGroup -GroupId $identity -DisplayName $newDN
                if (-not $res.Success) {
                    $script:buResultList.Items.Add("FAIL rename '$identity': $($res.Error)") | Out-Null
                    $lineOk = $false
                }
            }
            if ($addMbr) {
                $res = Add-DGMember -GroupId $identity -UserId $addMbr
                if (-not $res.Success) {
                    $script:buResultList.Items.Add("FAIL add member '$addMbr' to '$identity': $($res.Error)") | Out-Null
                    $lineOk = $false
                }
            }
            if ($removeMbr) {
                $res = Remove-DGMember -GroupId $identity -MemberId $removeMbr
                if (-not $res.Success) {
                    $script:buResultList.Items.Add("FAIL remove member '$removeMbr' from '$identity': $($res.Error)") | Out-Null
                    $lineOk = $false
                }
            }

            if ($lineOk) {
                $script:buResultList.Items.Add("OK:   $identity") | Out-Null
                $ok++
            } else {
                $fail++
            }
        }
        $script:mainWindow.Cursor    = $null
        $script:btnBUApply.IsEnabled = $true
        $statusType = if ($fail -eq 0) { "success" } else { "error" }
        Set-DGStatusText $script:buStatus "Done: $ok updated, $fail failed." $statusType
        Write-MigrazeLog "Bulk update: $ok OK, $fail failed." $(if ($fail -eq 0) { "Success" } else { "Error" })
    })
}