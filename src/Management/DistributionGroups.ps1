# DistributionGroups.ps1 - Distribution Group management (redesigned)

function Initialize-DGView {
    param([System.Windows.Window]$window)

    $script:DiscoveredDGs = @()

    # ── Control references ──────────────────────────────────────────────────────

    # Home / layer panels
    $script:panelDGHome   = $script:mainWindow.FindName("PanelDGHome")
    $script:panelDGSingle = $script:mainWindow.FindName("PanelDGSingle")
    $script:panelDGBulk   = $script:mainWindow.FindName("PanelDGBulk")
    $script:cardSingleDG  = $script:mainWindow.FindName("CardSingleDG")
    $script:cardBulkDG    = $script:mainWindow.FindName("CardBulkDG")

    # Single sub-nav + panels
    $script:navSingleCreate = $script:mainWindow.FindName("NavSingleCreate")
    $script:navSingleSearch = $script:mainWindow.FindName("NavSingleSearch")
    $script:navSingleUpdate = $script:mainWindow.FindName("NavSingleUpdate")
    $script:pSingleCreate   = $script:mainWindow.FindName("PanelSingleCreate")
    $script:pSingleSearch   = $script:mainWindow.FindName("PanelSingleSearch")
    $script:pSingleUpdate   = $script:mainWindow.FindName("PanelSingleUpdate")

    # Single Create controls
    $script:cDisplayName     = $script:mainWindow.FindName("C_DisplayName")
    $script:cMailNickname    = $script:mainWindow.FindName("C_MailNickname")
    $script:cSecurityEnabled = $script:mainWindow.FindName("C_SecurityEnabled")
    $script:cStatus          = $script:mainWindow.FindName("C_Status")
    $script:btnCreate        = $script:mainWindow.FindName("BtnCreate")

    # Single Search controls
    $script:sSearch          = $script:mainWindow.FindName("S_Search")
    $script:btnSSearch       = $script:mainWindow.FindName("BtnSSearch")
    $script:sDGList          = $script:mainWindow.FindName("S_DGList")
    $script:sPropsBox        = $script:mainWindow.FindName("S_PropsBox")
    $script:sPropsContent    = $script:mainWindow.FindName("S_PropsContent")
    $script:sMembersBox      = $script:mainWindow.FindName("S_MembersBox")
    $script:sMemberHeader    = $script:mainWindow.FindName("S_MemberHeader")
    $script:sMbrList         = $script:mainWindow.FindName("S_MbrList")
    $script:sAddUsrSearch    = $script:mainWindow.FindName("S_AddUsrSearch")
    $script:btnSAddUsrSearch = $script:mainWindow.FindName("BtnSAddUsrSearch")
    $script:sAddUsrList      = $script:mainWindow.FindName("S_AddUsrList")
    $script:btnSAddMember    = $script:mainWindow.FindName("BtnSAddMember")
    $script:btnSRemoveMember = $script:mainWindow.FindName("BtnSRemoveMember")
    $script:sStatus          = $script:mainWindow.FindName("S_Status")

    # Single Update controls
    $script:upSearch      = $script:mainWindow.FindName("UP_Search")
    $script:btnUPSearch   = $script:mainWindow.FindName("BtnUPSearch")
    $script:upDGList      = $script:mainWindow.FindName("UP_DGList")
    $script:upFieldsPanel = $script:mainWindow.FindName("UP_FieldsPanel")
    $script:upDisplayName = $script:mainWindow.FindName("UP_DisplayName")
    $script:upStatus      = $script:mainWindow.FindName("UP_Status")
    $script:btnUPSave     = $script:mainWindow.FindName("BtnUPSave")

    # Bulk sub-nav + panels
    $script:navBulkDiscover = $script:mainWindow.FindName("NavBulkDiscover")
    $script:navBulkCreate   = $script:mainWindow.FindName("NavBulkCreate")
    $script:navBulkUpdate   = $script:mainWindow.FindName("NavBulkUpdate")
    $script:pBulkDiscover   = $script:mainWindow.FindName("PanelBulkDiscover")
    $script:pBulkCreate     = $script:mainWindow.FindName("PanelBulkCreate")
    $script:pBulkUpdate     = $script:mainWindow.FindName("PanelBulkUpdate")

    # Bulk Discover controls
    $script:btnDiscoverAll = $script:mainWindow.FindName("BtnDiscoverAll")
    $script:btnExportCSV   = $script:mainWindow.FindName("BtnExportCSV")
    $script:discStatus     = $script:mainWindow.FindName("Disc_Status")
    $script:dgResultList   = $script:mainWindow.FindName("DG_ResultList")

    # Bulk Create controls
    $script:bcCsvText    = $script:mainWindow.FindName("BC_CsvText")
    $script:btnBCBrowse  = $script:mainWindow.FindName("BtnBCBrowse")
    $script:btnBCCreate  = $script:mainWindow.FindName("BtnBCCreate")
    $script:bcStatus     = $script:mainWindow.FindName("BC_Status")
    $script:bcResultList = $script:mainWindow.FindName("BC_ResultList")

    # Bulk Update controls
    $script:buCsvText    = $script:mainWindow.FindName("BU_CsvText")
    $script:btnBUBrowse  = $script:mainWindow.FindName("BtnBUBrowse")
    $script:btnBUApply   = $script:mainWindow.FindName("BtnBUApply")
    $script:buStatus     = $script:mainWindow.FindName("BU_Status")
    $script:buResultList = $script:mainWindow.FindName("BU_ResultList")

    # ── Navigation helpers ──────────────────────────────────────────────────────

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
        $panels = @($script:pSingleCreate, $script:pSingleSearch, $script:pSingleUpdate)
        $btns   = @($script:navSingleCreate, $script:navSingleSearch, $script:navSingleUpdate)
        for ($i = 0; $i -lt $panels.Count; $i++) {
            $panels[$i].Visibility = if ($i -eq $Idx) { "Visible" } else { "Collapsed" }
            $btns[$i].Background   = if ($i -eq $Idx) { [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0078D4")) } else { [Windows.Media.Brushes]::Transparent }
            $btns[$i].FontWeight   = if ($i -eq $Idx) { "SemiBold" } else { "Normal" }
        }
    }

    function script:Switch-BulkPanel {
        param([int]$Idx)
        $panels = @($script:pBulkDiscover, $script:pBulkCreate, $script:pBulkUpdate)
        $btns   = @($script:navBulkDiscover, $script:navBulkCreate, $script:navBulkUpdate)
        for ($i = 0; $i -lt $panels.Count; $i++) {
            $panels[$i].Visibility = if ($i -eq $Idx) { "Visible" } else { "Collapsed" }
            $btns[$i].Background   = if ($i -eq $Idx) { [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0078D4")) } else { [Windows.Media.Brushes]::Transparent }
            $btns[$i].FontWeight   = if ($i -eq $Idx) { "SemiBold" } else { "Normal" }
        }
    }

    function script:Set-DGStatus {
        param($ctrl, [string]$Msg, [string]$Type = "info")
        $ctrl.Text = $Msg
        $ctrl.Foreground = switch ($Type) {
            "success" { [Windows.Media.Brushes]::DarkGreen }
            "error"   { [Windows.Media.Brushes]::Crimson }
            default   { [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#0078D4")) }
        }
        $ctrl.Visibility = "Visible"
    }

    function script:Fill-DGList {
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

    function script:Refresh-SearchMembers {
        $grp = $script:sDGList.SelectedItem
        if ($null -eq $grp) { return }
        try {
            $res = Get-DGProperties -GroupId $grp.Id
            if (-not $res.Success) { return }
            $members = $res.Members
            $script:sMbrList.Items.Clear()
            foreach ($m in $members) {
                $mItem = [PSCustomObject]@{
                    Id          = $m.Id
                    UPN         = $m.Id
                    DisplayName = $m.DisplayName
                    ToString    = "$($m.DisplayName)  <$($m.Id)>"
                }
                $script:sMbrList.Items.Add($mItem) | Out-Null
            }
            $script:sMbrList.DisplayMemberPath = "ToString"
            $script:sMemberHeader.Text = "Members ($($members.Count))"
            $script:btnSRemoveMember.IsEnabled = $false
        } catch { }
    }

    # ── Home navigation ─────────────────────────────────────────────────────────

    $script:cardSingleDG.Add_MouseLeftButtonUp({ Show-DGSinglePanel })
    $script:cardBulkDG.Add_MouseLeftButtonUp({   Show-DGBulkPanel })
    $script:mainWindow.FindName("DGBtnBack").Add_Click({ Show-M365View })

    # ── Single sub-nav ──────────────────────────────────────────────────────────

    $script:mainWindow.FindName("BtnSingleBackHome").Add_Click({ Show-DGHomePanel })
    $script:navSingleCreate.Add_Click({ Switch-SinglePanel 0 })
    $script:navSingleSearch.Add_Click({ Switch-SinglePanel 1 })
    $script:navSingleUpdate.Add_Click({ Switch-SinglePanel 2 })

    # ── Bulk sub-nav ────────────────────────────────────────────────────────────

    $script:mainWindow.FindName("BtnBulkBackHome").Add_Click({ Show-DGHomePanel })
    $script:navBulkDiscover.Add_Click({ Switch-BulkPanel 0 })
    $script:navBulkCreate.Add_Click({   Switch-BulkPanel 1 })
    $script:navBulkUpdate.Add_Click({   Switch-BulkPanel 2 })

    # ── SINGLE CREATE ───────────────────────────────────────────────────────────

    $script:btnCreate.Add_Click({
        $dn  = $script:cDisplayName.Text.Trim()
        $mn  = $script:cMailNickname.Text.Trim()
        $sec = $script:cSecurityEnabled.IsChecked -eq $true

        if (-not $dn) { Set-DGStatus $script:cStatus "Display Name is required." "error"; return }
        if (-not $mn) { Set-DGStatus $script:cStatus "Email Alias is required." "error"; return }
        if ($mn -match '\s|[^a-zA-Z0-9._\-]') {
            Set-DGStatus $script:cStatus "MailNickname must not contain spaces or special characters." "error"
            return
        }

        Set-DGStatus $script:cStatus "Creating distribution group..." "info"
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnCreate.IsEnabled = $false
        try {
            $result = New-DGGroup -DisplayName $dn -MailNickname $mn -SecurityEnabled $sec
            if ($result.Success) {
                Set-DGStatus $script:cStatus "Distribution group '$dn' created successfully." "success"
                $script:cDisplayName.Text          = ""
                $script:cMailNickname.Text         = ""
                $script:cSecurityEnabled.IsChecked = $false
                Write-MigrazeLog "Created DG: $dn ($mn)" "Success"
            } else {
                Set-DGStatus $script:cStatus "Error: $($result.Error)" "error"
                Write-MigrazeLog "Failed to create DG '$dn': $($result.Error)" "Error"
            }
        } catch {
            Set-DGStatus $script:cStatus "Unexpected error: $_" "error"
        } finally {
            $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
            $script:btnCreate.IsEnabled = $true
        }
    })

    # ── SINGLE SEARCH ───────────────────────────────────────────────────────────

    $script:btnSSearch.Add_Click({
        $q = $script:sSearch.Text.Trim()
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $res = Get-DGList -SearchQuery $q
            if ($res.Success) {
                Fill-DGList $script:sDGList $res.Groups
                $script:sPropsBox.Visibility   = "Collapsed"
                $script:sMembersBox.Visibility = "Collapsed"
                if ($res.Groups.Count -eq 0) {
                    Set-DGStatus $script:sStatus "No distribution groups found." "info"
                } else {
                    $script:sStatus.Visibility = "Collapsed"
                }
            } else {
                Set-DGStatus $script:sStatus "Error: $($res.Error)" "error"
            }
        } catch {
            Set-DGStatus $script:sStatus "Error searching: $_" "error"
        } finally {
            $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })

    $script:sDGList.Add_SelectionChanged({
        $grp = $script:sDGList.SelectedItem
        if ($null -eq $grp) { return }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $res = Get-DGProperties -GroupId $grp.Id
            if (-not $res.Success) {
                Set-DGStatus $script:sStatus "Error loading properties: $($res.Error)" "error"
                return
            }
            $props = $res.Group

            # Populate properties card
            $script:sPropsContent.Children.Clear()
            $propRows = @(
                @{ L = "Display Name";     V = $props.DisplayName },
                @{ L = "Email";            V = $props.Mail },
                @{ L = "Alias";            V = $props.Alias },
                @{ L = "Security Enabled"; V = $props.SecurityEnabled }
            )
            foreach ($kv in $propRows) {
                $row = New-Object System.Windows.Controls.StackPanel
                $row.Orientation = "Horizontal"
                $row.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)

                $lbl = New-Object System.Windows.Controls.TextBlock
                $lbl.Text       = "$($kv.L):  "
                $lbl.FontWeight = "SemiBold"
                $lbl.FontSize   = 12
                $lbl.Foreground = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#2C3E50"))
                $lbl.Width      = 130

                $val = New-Object System.Windows.Controls.TextBlock
                $val.Text       = "$($kv.V)"
                $val.FontSize   = 12
                $val.Foreground = [Windows.Media.SolidColorBrush]([Windows.Media.ColorConverter]::ConvertFromString("#334455"))

                $row.Children.Add($lbl) | Out-Null
                $row.Children.Add($val) | Out-Null
                $script:sPropsContent.Children.Add($row) | Out-Null
            }
            $script:sPropsBox.Visibility = "Visible"

            # Load members
            $members = $res.Members
            $script:sMbrList.Items.Clear()
            foreach ($m in $members) {
                $mItem = [PSCustomObject]@{
                    Id          = $m.Id
                    UPN         = $m.Id
                    DisplayName = $m.DisplayName
                    ToString    = "$($m.DisplayName)  <$($m.Id)>"
                }
                $script:sMbrList.Items.Add($mItem) | Out-Null
            }
            $script:sMbrList.DisplayMemberPath = "ToString"
            $script:sMemberHeader.Text = "Members ($($members.Count))"
            $script:sMembersBox.Visibility = "Visible"
            $script:btnSRemoveMember.IsEnabled = $false
        } catch {
            Set-DGStatus $script:sStatus "Error loading properties: $_" "error"
        } finally {
            $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })

    $script:btnSAddUsrSearch.Add_Click({
        $q = $script:sAddUsrSearch.Text.Trim()
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $res = Search-MigrazeUsers -Query $q
            $script:sAddUsrList.Items.Clear()
            if ($res.Success) {
                foreach ($u in $res.Users) {
                    $uItem = [PSCustomObject]@{
                        Id          = $u.Id
                        UPN         = $u.Id
                        DisplayName = $u.DisplayName
                        ToString    = "$($u.DisplayName)  <$($u.Id)>"
                    }
                    $script:sAddUsrList.Items.Add($uItem) | Out-Null
                }
                $script:sAddUsrList.DisplayMemberPath = "ToString"
            } else {
                Set-DGStatus $script:sStatus "Error: $($res.Error)" "error"
            }
        } catch {
            Set-DGStatus $script:sStatus "Error searching users: $_" "error"
        } finally {
            $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })

    $script:sAddUsrList.Add_SelectionChanged({
        $script:btnSAddMember.IsEnabled = (
            $null -ne $script:sDGList.SelectedItem -and
            $null -ne $script:sAddUsrList.SelectedItem
        )
    })

    $script:btnSAddMember.Add_Click({
        $grp = $script:sDGList.SelectedItem
        $usr = $script:sAddUsrList.SelectedItem
        if ($null -eq $grp -or $null -eq $usr) { return }

        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnSAddMember.IsEnabled = $false
        try {
            $result = Add-DGMember -GroupId $grp.Id -UserId $usr.UPN
            if ($result.Success) {
                Set-DGStatus $script:sStatus "Member '$($usr.DisplayName)' added successfully." "success"
                Write-MigrazeLog "Added '$($usr.UPN)' to DG '$($grp.DisplayName)'." "Success"
                Refresh-SearchMembers
            } else {
                Set-DGStatus $script:sStatus "Error: $($result.Error)" "error"
            }
        } catch {
            Set-DGStatus $script:sStatus "Error adding member: $_" "error"
        } finally {
            $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
            $script:btnSAddMember.IsEnabled = $true
        }
    })

    $script:sMbrList.Add_SelectionChanged({
        $script:btnSRemoveMember.IsEnabled = ($script:sMbrList.SelectedItems.Count -gt 0)
    })

    $script:btnSRemoveMember.Add_Click({
        $grp = $script:sDGList.SelectedItem
        if ($null -eq $grp) { return }
        $selected = @($script:sMbrList.SelectedItems)
        if ($selected.Count -eq 0) { return }

        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnSRemoveMember.IsEnabled = $false
        $removed = 0; $failed = 0
        foreach ($m in $selected) {
            try {
                $result = Remove-DGMember -GroupId $grp.Id -MemberId $m.UPN
                if ($result.Success) { $removed++ } else { $failed++ }
            } catch { $failed++ }
        }
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
        $msg  = if ($failed -eq 0) { "Removed $removed member(s) successfully." } else { "Removed $removed, failed $failed." }
        $type = if ($failed -eq 0) { "success" } else { "error" }
        Set-DGStatus $script:sStatus $msg $type
        Write-MigrazeLog "Remove members from '$($grp.DisplayName)': $msg" $(if ($failed -eq 0) { "Success" } else { "Error" })
        Refresh-SearchMembers
    })

    # ── SINGLE UPDATE ───────────────────────────────────────────────────────────

    $script:btnUPSearch.Add_Click({
        $q = $script:upSearch.Text.Trim()
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $res = Get-DGList -SearchQuery $q
            if ($res.Success) {
                Fill-DGList $script:upDGList $res.Groups
                $script:upFieldsPanel.Visibility = "Collapsed"
            } else {
                Set-DGStatus $script:upStatus "Error: $($res.Error)" "error"
            }
        } catch {
            Set-DGStatus $script:upStatus "Error: $_" "error"
        } finally {
            $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })

    $script:upDGList.Add_SelectionChanged({
        $grp = $script:upDGList.SelectedItem
        if ($null -eq $grp) { return }
        $script:upDisplayName.Text       = $grp.DisplayName
        $script:upFieldsPanel.Visibility = "Visible"
        $script:upStatus.Visibility      = "Collapsed"
    })

    $script:btnUPSave.Add_Click({
        $grp = $script:upDGList.SelectedItem
        if ($null -eq $grp) { return }
        $newDN = $script:upDisplayName.Text.Trim()
        if (-not $newDN) { Set-DGStatus $script:upStatus "Display Name cannot be empty." "error"; return }

        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnUPSave.IsEnabled = $false
        try {
            $result = Update-DGGroup -GroupId $grp.Id -DisplayName $newDN
            if ($result.Success) {
                Set-DGStatus $script:upStatus "Display name updated to '$newDN'." "success"
                Write-MigrazeLog "Updated DG '$($grp.DisplayName)' -> '$newDN'." "Success"
            } else {
                Set-DGStatus $script:upStatus "Error: $($result.Error)" "error"
            }
        } catch {
            Set-DGStatus $script:upStatus "Unexpected error: $_" "error"
        } finally {
            $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
            $script:btnUPSave.IsEnabled = $true
        }
    })

    # ── BULK DISCOVER ───────────────────────────────────────────────────────────

    $script:btnDiscoverAll.Add_Click({
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnDiscoverAll.IsEnabled = $false
        Set-DGStatus $script:discStatus "Discovering all distribution groups..." "info"
        try {
            $res = Get-AllDGsForDiscovery
            if ($res.Success) {
                $script:DiscoveredDGs = $res.Groups
                $script:dgResultList.Items.Clear()
                foreach ($g in $res.Groups) { $script:dgResultList.Items.Add($g) | Out-Null }
                Set-DGStatus $script:discStatus "Found $($res.Groups.Count) distribution group(s)." "success"
                $script:btnExportCSV.IsEnabled = ($res.Groups.Count -gt 0)
                Write-MigrazeLog "Discovered $($res.Groups.Count) DG(s)." "Success"
            } else {
                Set-DGStatus $script:discStatus "Error: $($res.Error)" "error"
            }
        } catch {
            Set-DGStatus $script:discStatus "Error: $_" "error"
        } finally {
            $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
            $script:btnDiscoverAll.IsEnabled = $true
        }
    })

    $script:btnExportCSV.Add_Click({
        $dlg           = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Filter    = "CSV Files (*.csv)|*.csv"
        $dlg.FileName  = "DistributionGroups_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($dlg.ShowDialog() -eq $true) {
            try {
                $script:DiscoveredDGs | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
                Set-DGStatus $script:discStatus "Exported to: $($dlg.FileName)" "success"
                Write-MigrazeLog "Exported DGs to $($dlg.FileName)." "Success"
            } catch {
                Set-DGStatus $script:discStatus "Export failed: $_" "error"
            }
        }
    })

    # ── BULK CREATE ─────────────────────────────────────────────────────────────

    $script:btnBCBrowse.Add_Click({
        $dlg        = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        if ($dlg.ShowDialog() -eq $true) {
            try {
                $script:bcCsvText.Text = [System.IO.File]::ReadAllText($dlg.FileName)
            } catch {
                Set-DGStatus $script:bcStatus "Error loading file: $_" "error"
            }
        }
    })

    $script:btnBCCreate.Add_Click({
        $lines = $script:bcCsvText.Text -split "`n" | Where-Object { $_.Trim() -ne "" }
        if ($lines.Count -eq 0) { Set-DGStatus $script:bcStatus "No data to process." "error"; return }

        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnBCCreate.IsEnabled = $false
        $script:bcResultList.Items.Clear()
        Set-DGStatus $script:bcStatus "Processing $($lines.Count) line(s)..." "info"

        $ok = 0; $err = 0
        foreach ($line in $lines) {
            $parts = $line.Trim() -split ",", 3
            $dn    = if ($parts.Count -gt 0) { $parts[0].Trim() } else { "" }
            $alias = if ($parts.Count -gt 1) { $parts[1].Trim() } else { "" }
            $type  = if ($parts.Count -gt 2) { $parts[2].Trim() } else { "Distribution" }
            $sec   = ($type -ieq "Security")

            if (-not $dn -or -not $alias) {
                $script:bcResultList.Items.Add("SKIP (missing fields): $line") | Out-Null
                $err++
                continue
            }
            try {
                $result = New-DGGroup -DisplayName $dn -MailNickname $alias -SecurityEnabled $sec
                if ($result.Success) {
                    $script:bcResultList.Items.Add("OK: $dn ($alias)") | Out-Null
                    $ok++
                } else {
                    $script:bcResultList.Items.Add("FAIL: $dn - $($result.Error)") | Out-Null
                    $err++
                }
            } catch {
                $script:bcResultList.Items.Add("ERROR: $dn - $_") | Out-Null
                $err++
            }
        }
        $type2 = if ($err -eq 0) { "success" } else { "error" }
        Set-DGStatus $script:bcStatus "Done: $ok created, $err failed." $type2
        Write-MigrazeLog "Bulk create: $ok OK, $err failed." $(if ($err -eq 0) { "Success" } else { "Error" })
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
        $script:btnBCCreate.IsEnabled = $true
    })

    # ── BULK UPDATE ─────────────────────────────────────────────────────────────

    $script:btnBUBrowse.Add_Click({
        $dlg        = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        if ($dlg.ShowDialog() -eq $true) {
            try {
                $script:buCsvText.Text = [System.IO.File]::ReadAllText($dlg.FileName)
            } catch {
                Set-DGStatus $script:buStatus "Error loading file: $_" "error"
            }
        }
    })

    $script:btnBUApply.Add_Click({
        $lines = $script:buCsvText.Text -split "`n" | Where-Object { $_.Trim() -ne "" }
        if ($lines.Count -eq 0) { Set-DGStatus $script:buStatus "No data to process." "error"; return }

        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
        $script:btnBUApply.IsEnabled = $false
        $script:buResultList.Items.Clear()
        Set-DGStatus $script:buStatus "Processing $($lines.Count) line(s)..." "info"

        $ok = 0; $err = 0
        foreach ($line in $lines) {
            $parts     = $line.Trim() -split ",", 4
            $identity  = if ($parts.Count -gt 0) { $parts[0].Trim() } else { "" }
            $newDN     = if ($parts.Count -gt 1) { $parts[1].Trim() } else { "" }
            $addMbr    = if ($parts.Count -gt 2) { $parts[2].Trim() } else { "" }
            $removeMbr = if ($parts.Count -gt 3) { $parts[3].Trim() } else { "" }

            if (-not $identity) {
                $script:buResultList.Items.Add("SKIP (no identity): $line") | Out-Null
                $err++
                continue
            }
            try {
                $grpRes = Get-DGList -SearchQuery $identity
                if (-not $grpRes.Success) { throw "Lookup failed: $($grpRes.Error)" }
                $grp = $grpRes.Groups | Where-Object { $_.Mail -ieq $identity -or $_.DisplayName -ieq $identity } | Select-Object -First 1
                if ($null -eq $grp) { throw "Group '$identity' not found." }

                if ($newDN) {
                    $res = Update-DGGroup -GroupId $grp.Id -DisplayName $newDN
                    if (-not $res.Success) { throw "Update failed: $($res.Error)" }
                }
                if ($addMbr) {
                    $res = Add-DGMember -GroupId $grp.Id -UserId $addMbr
                    if (-not $res.Success) { throw "Add member failed: $($res.Error)" }
                }
                if ($removeMbr) {
                    $res = Remove-DGMember -GroupId $grp.Id -MemberId $removeMbr
                    if (-not $res.Success) { throw "Remove member failed: $($res.Error)" }
                }
                $script:buResultList.Items.Add("OK: $identity") | Out-Null
                $ok++
            } catch {
                $script:buResultList.Items.Add("FAIL: $identity - $_") | Out-Null
                $err++
            }
        }
        $type3 = if ($err -eq 0) { "success" } else { "error" }
        Set-DGStatus $script:buStatus "Done: $ok updated, $err failed." $type3
        Write-MigrazeLog "Bulk update: $ok OK, $err failed." $(if ($err -eq 0) { "Success" } else { "Error" })
        $script:mainWindow.Cursor = [System.Windows.Input.Cursors]::Arrow
        $script:btnBUApply.IsEnabled = $true
    })
}