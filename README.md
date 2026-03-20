# ExoMan v1.0 – Exchange Online Management Tool

A Windows GUI tool for administrators to perform post-migration tasks on Exchange Online using the **Microsoft Graph PowerShell SDK**.

---

## Features

| Module | Status | Operations |
|--------|--------|------------|
| **Distribution Groups** | ✅ Ready | Create DG · Update Properties · Add Members · Remove Members · Read Properties |
| **Shared Mailbox** | 🔜 Coming soon | — |
| **User Mailbox** | 🔜 Coming soon | — |

---

## Prerequisites

- Windows 10 / 11 or Windows Server 2019+
- PowerShell 5.1 or PowerShell 7+
- Internet access (for Microsoft 365 login and module installation)
- The following PowerShell modules *(ExoMan will offer to install them automatically)*:
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.Groups`
  - `Microsoft.Graph.Users`

---

## Quick Start

### Option 1 – Double-click launcher
```
Launch-ExoMan.bat
```

### Option 2 – PowerShell
```powershell
powershell -STA -ExecutionPolicy Bypass -File ExoMan.ps1
```

---

## Authentication

ExoMan uses **interactive Microsoft 365 login via your default browser**.  
Click **"Connect to Exchange Online"** on the main screen — your browser will open the Microsoft identity platform login page. Sign in with your admin account.

Required Graph API scopes:
- `User.Read`, `User.ReadBasic.All`
- `Group.Read.All`, `Group.ReadWrite.All`
- `GroupMember.Read.All`, `GroupMember.ReadWrite.All`
- `Directory.Read.All`

---

## Distribution Group Operations

| Operation | Description |
|-----------|-------------|
| **Create Distribution Group** | Creates a new mail-enabled group (optionally security-enabled) |
| **Update DG Properties** | Search for a DG and update Display Name or Description |
| **Add Members** | Search for a DG and a user, then add the user as a member |
| **Remove Members** | Load current members of a DG and remove one or more |
| **Read Current Properties** | View all properties and full member list of any DG |

---

## File Structure

```
exoman/
├── ExoMan.ps1               # Entry point
├── Launch-ExoMan.bat        # Easy double-click launcher
├── src/
│   ├── GraphHelper.ps1      # Microsoft Graph auth + DG operations
│   ├── MainWindow.ps1       # Main application window
│   ├── DistributionGroups.ps1  # Distribution Groups management window
│   ├── SharedMailbox.ps1    # Shared Mailbox (coming soon)
│   └── UserMailbox.ps1      # User Mailbox (coming soon)
└── README.md
```

---

## Technology

- **UI**: Windows Presentation Foundation (WPF) via PowerShell
- **API**: Microsoft Graph PowerShell SDK
- **Auth**: Interactive OAuth 2.0 browser login (`Connect-MgGraph`)

---

## Version History

| Version | Notes |
|---------|-------|
| 1.0 | Initial release – Distribution Group management |
