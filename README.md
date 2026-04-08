# MSAccess-agent

**Access-POSH.ps1** is a port of the Python unmateria/MCP-Access server (54+ tools) to native PowerShell.

No MCP server needed — A custom access-dev agent calls functions directly via terminal.

The agent will use the **Access-POSH.ps1** PowerShell script to interact with Access databases via COM automation. 

Setup:
1. Clone or download the repo
2. Put the two .md files in C:\Users\\%USERNAME%\AppData\Roaming\Code\User\prompts folder 
3. Replace the path in the .md files to the location of the **Access-POSH.ps1** script on your computer
4. Select access-dev as the agent

## Available Functions (54 public)

| Category | Functions |
|----------|-----------|
| **Database** | `New-AccessDatabase`, `Close-AccessDatabase`, `Repair-AccessDatabase` |
| **Objects** | `Get-AccessObject`, `Get-AccessCode`, `Set-AccessCode`, `Remove-AccessObject`, `Export-AccessStructure` |
| **SQL** | `Invoke-AccessSQL`, `Invoke-AccessSQLBatch` |
| **Tables** | `Get-AccessTableInfo`, `New-AccessTable`, `Edit-AccessTable` |
| **VBE** | `Get-AccessVbeLine`, `Get-AccessVbeProc`, `Get-AccessVbeModuleInfo`, `Set-AccessVbeLine`, `Set-AccessVbeProc`, `Add-AccessVbeCode` |
| **Search** | `Find-AccessVbeText`, `Search-AccessVbe`, `Search-AccessQuery`, `Find-AccessUsage` |
| **VBA Exec** | `Invoke-AccessMacro`, `Invoke-AccessVba`, `Invoke-AccessEval`, `Test-AccessVbaCompile` |
| **Forms** | `New-AccessForm`, `Get-AccessFormProperty`, `Set-AccessFormProperty` |
| **Controls** | `Get-AccessControl`, `Get-AccessControlDetail`, `New-AccessControl`, `Remove-AccessControl`, `Set-AccessControlProperty`, `Set-AccessControlBatch` |
| **Fields** | `Get-AccessFieldProperty`, `Set-AccessFieldProperty` |
| **Linked Tables** | `Get-AccessLinkedTable`, `Set-AccessLinkedTable` |
| **Relationships** | `Get-AccessRelationship`, `New-AccessRelationship`, `Remove-AccessRelationship` |
| **References** | `Get-AccessReference`, `Set-AccessReference` |
| **Queries** | `Set-AccessQuery` |
| **Indexes** | `Get-AccessIndex`, `Set-AccessIndex` |
| **Properties** | `Get-AccessDatabaseProperty`, `Set-AccessDatabaseProperty`, `Get-AccessStartupOption` |
| **Export** | `Export-AccessReport`, `Copy-AccessData` |
| **UI** | `Get-AccessScreenshot`, `Send-AccessClick`, `Send-AccessKeyboard` |
| **Tips** | `Get-AccessTip` |
