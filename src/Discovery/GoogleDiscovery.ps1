# GoogleDiscovery.ps1 - Google Workspace discovery operations

function Show-SaveFileDialog {
    param([string]$Title, [string]$Filter, [string]$DefaultName)
    Add-Type -AssemblyName System.Windows.Forms
    $sfd          = [System.Windows.Forms.SaveFileDialog]::new()
    $sfd.Title    = $Title
    $sfd.Filter   = $Filter
    $sfd.FileName = $DefaultName
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $sfd.FileName }
    return $null
}

function Invoke-PagedGoogleAPI {
    param([string]$BaseUrl, [string]$CollectionKey, [string]$PageParam = "pageToken")
    $all   = @()
    $token = $null
    do {
        $url = $BaseUrl
        if ($token) { $url += "&${PageParam}=$token" }
        $resp  = Invoke-GoogleAPI -Url $url
        $items = $resp.$CollectionKey
        if ($items) { $all += $items }
        $token = $resp.nextPageToken
    } while ($token)
    return $all
}

function Get-GoogleUsers {
    param([switch]$ExportCSV)
    Invoke-WithGoogleAuth {
        try {
            Write-MigrazeLog "Discovering Google Workspace users..." "Action"
            $users = Invoke-PagedGoogleAPI -BaseUrl "https://admin.googleapis.com/admin/directory/v1/users?customer=my_customer&maxResults=500&orderBy=email" -CollectionKey "users"
            Write-MigrazeLog "Discovered $($users.Count) Google Workspace user(s)." "Success"
            if ($ExportCSV) {
                $csv = $users | Select-Object `
                    @{N="Email";E={$_.primaryEmail}},
                    @{N="FirstName";E={$_.name.givenName}},
                    @{N="LastName";E={$_.name.familyName}},
                    @{N="OrgUnit";E={$_.orgUnitPath}},
                    @{N="Suspended";E={$_.suspended}},
                    @{N="Admin";E={$_.isAdmin}},
                    @{N="LastLogin";E={$_.lastLoginTime}},
                    @{N="Created";E={$_.creationTime}}
                $p = Show-SaveFileDialog -Title "Export Google Users" -Filter "CSV files (*.csv)|*.csv" -DefaultName "GoogleUsers.csv"
                if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
            }
            return @{ Success = $true; Users = $users }
        } catch {
            Write-MigrazeLog "Google user discovery failed: $($_.Exception.Message)" "Error"
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-GoogleGroups {
    param([switch]$ExportCSV)
    Invoke-WithGoogleAuth {
        try {
            Write-MigrazeLog "Discovering Google Workspace groups..." "Action"
            $groups = Invoke-PagedGoogleAPI -BaseUrl "https://admin.googleapis.com/admin/directory/v1/groups?customer=my_customer&maxResults=200" -CollectionKey "groups"
            Write-MigrazeLog "Discovered $($groups.Count) Google group(s)." "Success"
            if ($ExportCSV) {
                $csv = $groups | Select-Object `
                    @{N="Email";E={$_.email}},
                    @{N="Name";E={$_.name}},
                    @{N="Description";E={$_.description}},
                    @{N="Members";E={$_.directMembersCount}}
                $p = Show-SaveFileDialog -Title "Export Google Groups" -Filter "CSV files (*.csv)|*.csv" -DefaultName "GoogleGroups.csv"
                if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
            }
            return @{ Success = $true; Groups = $groups }
        } catch {
            Write-MigrazeLog "Google group discovery failed: $($_.Exception.Message)" "Error"
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-GoogleDomains {
    param([switch]$ExportCSV)
    Invoke-WithGoogleAuth {
        try {
            Write-MigrazeLog "Discovering Google Workspace domains..." "Action"
            $resp = Invoke-GoogleAPI -Url "https://admin.googleapis.com/admin/directory/v1/customer/my_customer/domains"
            $domains = $resp.domains
            Write-MigrazeLog "Discovered $($domains.Count) domain(s)." "Success"
            if ($ExportCSV) {
                $csv = $domains | Select-Object domainName, isPrimary, verified, creationTime
                $p = Show-SaveFileDialog -Title "Export Domains" -Filter "CSV files (*.csv)|*.csv" -DefaultName "GoogleDomains.csv"
                if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
            }
            return @{ Success = $true; Domains = $domains }
        } catch {
            Write-MigrazeLog "Google domain discovery failed: $($_.Exception.Message)" "Error"
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-GoogleOrgUnits {
    param([switch]$ExportCSV)
    Invoke-WithGoogleAuth {
        try {
            Write-MigrazeLog "Discovering Google Workspace org units..." "Action"
            $resp = Invoke-GoogleAPI -Url "https://admin.googleapis.com/admin/directory/v1/customer/my_customer/orgunits?type=all"
            $ous  = $resp.organizationUnits
            Write-MigrazeLog "Discovered $($ous.Count) org unit(s)." "Success"
            if ($ExportCSV) {
                $csv = $ous | Select-Object name, orgUnitPath, parentOrgUnitPath, description
                $p = Show-SaveFileDialog -Title "Export Org Units" -Filter "CSV files (*.csv)|*.csv" -DefaultName "GoogleOrgUnits.csv"
                if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
            }
            return @{ Success = $true; OrgUnits = $ous }
        } catch {
            Write-MigrazeLog "Google OU discovery failed: $($_.Exception.Message)" "Error"
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-GoogleSharedDrives {
    param([switch]$ExportCSV)
    Invoke-WithGoogleAuth {
        try {
            Write-MigrazeLog "Discovering Google Shared Drives..." "Action"
            $resp   = Invoke-GoogleAPI -Url "https://www.googleapis.com/drive/v3/drives?pageSize=100"
            $drives = $resp.drives
            Write-MigrazeLog "Discovered $($drives.Count) shared drive(s)." "Success"
            if ($ExportCSV) {
                $csv = $drives | Select-Object id, name, @{N="CreatedTime";E={$_.createdTime}}
                $p = Show-SaveFileDialog -Title "Export Shared Drives" -Filter "CSV files (*.csv)|*.csv" -DefaultName "GoogleSharedDrives.csv"
                if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
            }
            return @{ Success = $true; Drives = $drives }
        } catch {
            Write-MigrazeLog "Google Shared Drives discovery failed: $($_.Exception.Message)" "Error"
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    }
}

function Get-GoogleCollabMailboxes {
    param([switch]$ExportCSV)
    Invoke-WithGoogleAuth {
        try {
            Write-MigrazeLog "Discovering Google collaboration mailboxes (Groups with external email)..." "Action"
            $groups = Invoke-PagedGoogleAPI -BaseUrl "https://admin.googleapis.com/admin/directory/v1/groups?customer=my_customer&maxResults=200" -CollectionKey "groups"
            Write-MigrazeLog "Discovered $($groups.Count) collaboration group(s)." "Success"
            if ($ExportCSV) {
                $csv = $groups | Select-Object @{N="Email";E={$_.email}}, @{N="Name";E={$_.name}}, @{N="Members";E={$_.directMembersCount}}, @{N="Description";E={$_.description}}
                $p = Show-SaveFileDialog -Title "Export Collab Mailboxes" -Filter "CSV files (*.csv)|*.csv" -DefaultName "GoogleCollabMailboxes.csv"
                if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
            }
            return @{ Success = $true; Groups = $groups }
        } catch {
            Write-MigrazeLog "Google collab mailbox discovery failed: $($_.Exception.Message)" "Error"
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    }
}