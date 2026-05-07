# MSAccess-CLI-Agent

> **Automate Microsoft Access from plain English inside VS Code — no MCP server, no Python, no extra processes.**

![Platform: Windows](https://img.shields.io/badge/platform-Windows-blue?logo=windows)
![PowerShell: 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![VS Code](https://img.shields.io/badge/VS%20Code-GitHub%20Copilot%20Chat-blueviolet?logo=visual-studio-code)
![Functions: 91](https://img.shields.io/badge/functions-91-brightgreen)
![Module Version](https://img.shields.io/badge/version-1.0.0-orange)
![License: MIT](https://img.shields.io/badge/license-MIT-green)

## What is this?

**MSAccess-CLI-Agent** is a VS Code agent (powered by GitHub Copilot Chat) that lets you talk to Microsoft Access in plain language. You describe what you want and the agent translates it into PowerShell commands that manipulate your `.accdb` / `.mdb` database live via COM — no manual VBA editing required.

```text
You:   "Create a Customers table with ID (AutoNumber PK), Name (Text 100), and Email (Text 255)"
Agent: → New-AccessTable → confirms success
```

The **AccessPOSH** module (included) is a PowerShell port of [unmateria/MCP-Access](https://github.com/unmateria/MCP-Access), expanded to **91 public functions** covering databases, tables, forms, controls, VBA/VBE, SQL, reports, imports, security, and UI automation.

## How it works

```
VS Code Copilot Chat (agent mode)
        │
        ▼
  access-dev agent (.md instructions)
        │  describes which PowerShell command to run
        ▼
  AccessPOSH module  (imported in the VS Code terminal)
        │  COM calls via DAO / Access Object Model
        ▼
  Microsoft Access (.accdb / .mdb)
```

- **No separate server** — the module runs directly in the VS Code integrated terminal.
- **No Python / Node** — pure PowerShell 5.1+ on Windows.
- **Full COM access** — everything you can do from VBA, you can do from the agent.
- **-WhatIf / -Confirm** — all state-changing functions support PowerShell's standard risk-mitigation flags.
- **Pester tests** — 17 test files cover every public command group.

## Prerequisites

| Requirement | Details |
|---|---|
| **OS** | Windows 10 / 11 (COM automation is Windows-only) |
| **Microsoft Access** | Access 2016, 2019, 2021, or Microsoft 365 (desktop) |
| **PowerShell** | 5.1 (Windows PowerShell) **or** PowerShell 7+ |
| **VS Code** | Latest stable, with the **GitHub Copilot Chat** extension |
| **Copilot** | An active GitHub Copilot subscription |

## Setup

### 1 — Clone the repo

```powershell
git clone https://github.com/jcolozzi/MSAccess-CLI-Agent.git
```

### 2 — Install the agent instructions

Choose **one** of the following:

**Option A — User-level (available in every workspace)**

Copy both `.md` files from the repo root to:
```
C:\Users\%USERNAME%\AppData\Roaming\Code\User\prompts\
```

**Option B — Workspace-level (scoped to this project)**

Copy both `.md` files into a `.github\agents\` folder in your workspace root. VS Code automatically detects any `.md` files in that folder as custom agents.

> **Note:** VS Code detects any `.md` files in the `.github/agents/` folder of your workspace as custom agents.

### 3 — Update the module path inside the agent files

Open each `.md` agent file and replace the placeholder path with the actual path to `AccessPOSH.psd1` on your machine:

```
# Before
Import-Module "C:\path\to\AccessPOSH\AccessPOSH.psd1"

# After (example)
Import-Module "C:\Projects\MSAccess-agent\AccessPOSH\AccessPOSH.psd1"
```

### 4 — Select the agent and start prompting

In VS Code Copilot Chat, click the agent picker and choose **access-dev**. Open (or have the agent open) a `.accdb` file, then start describing what you want.

## Usage examples

| Prompt | Functions called |
|---|---|
| "List all tables in MyDB.accdb" | `Get-AccessObject` |
| "Add an EmailAddress field (Text 255) to the Customers table" | `Edit-AccessTable` |
| "Create a one-to-many relationship between Customers and Orders" | `New-AccessRelationship` |
| "Run a query that deletes records older than 90 days" | `Invoke-AccessSQL` |
| "Show me the VBA in Module1" | `Get-AccessCode` |
| "Add error handling to the SaveRecord procedure" | `Update-AccessVbeProc` |
| "Export the Monthly Sales report to PDF" | `Export-AccessReport` |
| "Take a screenshot of the open form" | `Get-AccessScreenshot` |
| "Import data from Customers.csv into the Customers table" | `Import-AccessFromCSV` |
| "What would happen if I ran New-AccessTable? (dry run)" | `New-AccessTable -WhatIf` |

## Project structure

```
MSAccess-agent/
├── AccessPOSH/             # PowerShell module (the engine)
│   ├── AccessPOSH.psd1     # Module manifest (v1.0.0, PS 5.1+, Desktop + Core)
│   ├── AccessPOSH.psm1     # Module loader
│   ├── Public/             # 17 files — one per command category
│   │   ├── DatabaseOps.ps1
│   │   ├── TableOps.ps1
│   │   ├── FormReportOps.ps1
│   │   ├── VbeOps.ps1
│   │   └── ...
│   └── Private/            # Internal helpers (COM session, error formatting, etc.)
├── Tests/                  # Pester test suite — 17 test files
│   ├── DatabaseOps.Tests.ps1
│   ├── VbeOps.Tests.ps1
│   └── ...
├── access-dev.md           # Agent instructions (the Copilot Chat prompt)
└── README.md
```

## Running the tests

```powershell
# From the repo root
Invoke-Pester .\Tests\ -Output Detailed
```

> Requires [Pester](https://github.com/pester/Pester) 5.x: `Install-Module Pester -MinimumVersion 5.0 -Force`

## Function reference

<details>
<summary><strong>View all 91 public functions</strong></summary>

| Category | Functions |
|---|---|
| **Database** | `New-AccessDatabase`, `Close-AccessDatabase`, `Repair-AccessDatabase`, `Invoke-AccessDecompile` |
| **Objects** | `Get-AccessObject`, `Get-AccessCode`, `Set-AccessCode`, `Remove-AccessObject`, `Export-AccessStructure` |
| **SQL** | `Invoke-AccessSQL`, `Invoke-AccessSQLBatch` |
| **Tables** | `Get-AccessTableInfo`, `New-AccessTable`, `Edit-AccessTable` |
| **Fields** | `Get-AccessFieldProperty`, `Set-AccessFieldProperty` |
| **Indexes** | `Get-AccessIndex`, `Set-AccessIndex` |
| **VBE** | `Get-AccessVbeLine`, `Get-AccessVbeProc`, `Get-AccessVbeModuleInfo`, `Set-AccessVbeLine`, `Set-AccessVbeProc`, `Update-AccessVbeProc`, `Add-AccessVbeCode`, `Import-AccessVbaFile`, `Test-AccessVbaFileEncoding` |
| **Search** | `Find-AccessVbeText`, `Search-AccessVbe`, `Search-AccessQuery`, `Find-AccessUsage` |
| **VBA Execution** | `Invoke-AccessMacro`, `Invoke-AccessVba`, `Invoke-AccessEval`, `Test-AccessVbaCompile` |
| **Forms** | `New-AccessForm`, `Get-AccessFormProperty`, `Set-AccessFormProperty` |
| **Controls** | `Get-AccessControl`, `Get-AccessControlDetail`, `New-AccessControl`, `Remove-AccessControl`, `Set-AccessControlProperty`, `Set-AccessControlBatch` |
| **Linked Tables** | `Get-AccessLinkedTable`, `Set-AccessLinkedTable` |
| **Relationships** | `Get-AccessRelationship`, `New-AccessRelationship`, `Remove-AccessRelationship` |
| **References** | `Get-AccessReference`, `Set-AccessReference` |
| **Queries** | `Set-AccessQuery`, `Search-AccessQuery` |
| **Properties** | `Get-AccessDatabaseProperty`, `Set-AccessDatabaseProperty`, `Get-AccessStartupOption` |
| **Export** | `Export-AccessReport`, `Copy-AccessData`, `Export-AccessToExcel`, `Export-AccessFilteredReport` |
| **Import** | `Import-AccessFromExcel`, `Import-AccessFromCSV`, `Import-AccessFromXML`, `Import-AccessFromDatabase` |
| **Security** | `Test-AccessDatabasePassword`, `Set-AccessDatabasePassword`, `Remove-AccessDatabasePassword`, `Get-AccessDatabaseEncryption` |
| **Reports** | `New-AccessReport`, `Get-AccessGroupLevel`, `Set-AccessGroupLevel`, `Remove-AccessGroupLevel` |
| **SubDataSheets** | `Get-AccessSubDataSheet`, `Set-AccessSubDataSheet` |
| **Navigation Pane** | `Show-AccessNavigationPane`, `Hide-AccessNavigationPane`, `Set-AccessNavigationPaneLock` |
| **Ribbon** | `Get-AccessRibbon`, `Set-AccessRibbon`, `Remove-AccessRibbon` |
| **Themes** | `Get-AccessTheme`, `Set-AccessTheme`, `Get-AccessThemeList` |
| **TempVars** | `Get-AccessTempVar`, `Set-AccessTempVar`, `Remove-AccessTempVar` |
| **UI Automation** | `Get-AccessScreenshot`, `Send-AccessClick`, `Send-AccessKeyboard` |
| **Print** | `Send-AccessReportToPrinter`, `Export-AccessFilteredReport` |
| **Application** | `Get-AccessApplicationInfo`, `Test-AccessRuntime`, `Get-AccessFileInfo` |
| **Tips** | `Get-AccessTip` |

</details>

All state-changing functions support `-WhatIf` and `-Confirm` via PowerShell's standard `ShouldProcess` mechanism.

## Contributing

Pull requests are welcome. For significant changes, open an issue first to discuss what you would like to change. Please include or update Pester tests for any new or modified functions.

## Credits

- Original MCP server: [unmateria/MCP-Access](https://github.com/unmateria/MCP-Access)
- PowerShell port and VS Code agent integration: Access-POSH

## License

[MIT](LICENSE) © 2026 Access-POSH
