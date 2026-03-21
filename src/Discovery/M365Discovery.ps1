# M365Discovery.ps1 - Source M365 tenant discovery operations

function Get-M365Users {
    param([switch]$ExportCSV)
    try {
        Write-MigrazeLog "Discovering M365 source users..." "Action"
        Import-GraphModules
        $users = Get-MgUser -All `
            -Property "Id,DisplayName,UserPrincipalName,Mail,AccountEnabled,JobTitle,Department,CreatedDateTime" `
            -ErrorAction Stop
        Write-MigrazeLog "Discovered $($users.Count) M365 user(s)." "Success"
        if ($ExportCSV) {
            $csv = $users | Select-Object DisplayName,UserPrincipalName,Mail,AccountEnabled,JobTitle,Department,CreatedDateTime
            $p = Show-SaveFileDialog -Title "Export M365 Users" -Filter "CSV files (*.csv)|*.csv" -DefaultName "M365Users.csv"
            if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
        }
        return @{ Success = $true; Users = $users }
    } catch {
        Write-MigrazeLog "M365 user discovery failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-M365SecurityGroups {
    param([switch]$ExportCSV)
    try {
        Write-MigrazeLog "Discovering M365 security groups..." "Action"
        Import-GraphModules
        $groups = Get-MgGroup -Filter "securityEnabled eq true and NOT groupTypes/any(c:c eq 'Unified')" -All `
            -Property "Id,DisplayName,Mail,Description,CreatedDateTime" -ErrorAction Stop
        Write-MigrazeLog "Discovered $($groups.Count) security group(s)." "Success"
        if ($ExportCSV) {
            $csv = $groups | Select-Object DisplayName,Mail,Description,CreatedDateTime
            $p = Show-SaveFileDialog -Title "Export Security Groups" -Filter "CSV files (*.csv)|*.csv" -DefaultName "M365SecurityGroups.csv"
            if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
        }
        return @{ Success = $true; Groups = $groups }
    } catch {
        Write-MigrazeLog "M365 security group discovery failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-M365DistributionGroupsSrc {
    param([switch]$ExportCSV)
    try {
        Write-MigrazeLog "Discovering M365 distribution groups..." "Action"
        Import-GraphModules
        $groups = Get-MgGroup -Filter "mailEnabled eq true and NOT groupTypes/any(c:c eq 'Unified')" -All `
            -Property "Id,DisplayName,Mail,MailNickname,Description,CreatedDateTime" -ErrorAction Stop
        Write-MigrazeLog "Discovered $($groups.Count) distribution group(s)." "Success"
        if ($ExportCSV) {
            $csv = $groups | Select-Object DisplayName,Mail,MailNickname,Description,CreatedDateTime
            $p = Show-SaveFileDialog -Title "Export Distribution Groups" -Filter "CSV files (*.csv)|*.csv" -DefaultName "M365DistributionGroups.csv"
            if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
        }
        return @{ Success = $true; Groups = $groups }
    } catch {
        Write-MigrazeLog "M365 DG discovery failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-M365SharedMailboxesSrc {
    param([switch]$ExportCSV)
    try {
        Write-MigrazeLog "Discovering M365 shared mailboxes (via Graph - users)..." "Action"
        Import-GraphModules
        $users = Get-MgUser -Filter "userType eq 'Member'" -All `
            -Property "Id,DisplayName,UserPrincipalName,Mail,AccountEnabled" -ErrorAction Stop
        Write-MigrazeLog "Retrieved $($users.Count) user mailbox candidate(s)." "Info"
        Write-MigrazeLog "Note: For precise shared mailbox filtering, use Exchange Online PowerShell (Get-Mailbox -RecipientTypeDetails SharedMailbox)." "Warning"
        if ($ExportCSV) {
            $csv = $users | Select-Object DisplayName,UserPrincipalName,Mail,AccountEnabled
            $p = Show-SaveFileDialog -Title "Export Mailbox Users" -Filter "CSV files (*.csv)|*.csv" -DefaultName "M365SharedMailboxes.csv"
            if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
        }
        return @{ Success = $true; Users = $users }
    } catch {
        Write-MigrazeLog "M365 shared mailbox discovery failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-M365UserMailboxesSrc {
    param([switch]$ExportCSV)
    try {
        Write-MigrazeLog "Discovering M365 user mailboxes..." "Action"
        Import-GraphModules
        $users = Get-MgUser -Filter "accountEnabled eq true" -All `
            -Property "Id,DisplayName,UserPrincipalName,Mail,JobTitle,Department,CreatedDateTime" -ErrorAction Stop
        Write-MigrazeLog "Discovered $($users.Count) enabled user mailbox(es)." "Success"
        if ($ExportCSV) {
            $csv = $users | Select-Object DisplayName,UserPrincipalName,Mail,JobTitle,Department
            $p = Show-SaveFileDialog -Title "Export User Mailboxes" -Filter "CSV files (*.csv)|*.csv" -DefaultName "M365UserMailboxes.csv"
            if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
        }
        return @{ Success = $true; Users = $users }
    } catch {
        Write-MigrazeLog "M365 user mailbox discovery failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-M365RoomsResources {
    param([switch]$ExportCSV)
    try {
        Write-MigrazeLog "Discovering M365 rooms and resources..." "Action"
        Import-GraphModules
        $places = Get-MgPlace -All -ErrorAction Stop
        Write-MigrazeLog "Discovered $($places.Count) room/resource(s)." "Success"
        if ($ExportCSV) {
            $csv = $places | Select-Object DisplayName,@{N="Email";E={$_.AdditionalProperties.emailAddress}},@{N="Type";E={$_."@odata.type"}}
            $p = Show-SaveFileDialog -Title "Export Rooms and Resources" -Filter "CSV files (*.csv)|*.csv" -DefaultName "M365RoomsResources.csv"
            if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
        }
        return @{ Success = $true; Places = $places }
    } catch {
        Write-MigrazeLog "M365 rooms/resources discovery failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-M365Contacts {
    param([switch]$ExportCSV)
    try {
        Write-MigrazeLog "Discovering M365 org contacts..." "Action"
        Import-GraphModules
        $contacts = Get-MgContact -All -Property "Id,DisplayName,Mail,MailNickname" -ErrorAction Stop
        Write-MigrazeLog "Discovered $($contacts.Count) contact(s)." "Success"
        if ($ExportCSV) {
            $csv = $contacts | Select-Object DisplayName,Mail,MailNickname
            $p = Show-SaveFileDialog -Title "Export Contacts" -Filter "CSV files (*.csv)|*.csv" -DefaultName "M365Contacts.csv"
            if ($p) { $csv | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8; Write-MigrazeLog "Exported to: $p" "Success" }
        }
        return @{ Success = $true; Contacts = $contacts }
    } catch {
        Write-MigrazeLog "M365 contact discovery failed: $($_.Exception.Message)" "Error"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}