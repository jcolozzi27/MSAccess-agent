# MSAccess-agent

Access-POSH.ps1 is a PowerShell port of [unmateria/MCP-Access](https://github.com/unmateria/MCP-Access) project. It provides over 50 PowerShell functions for automating Microsoft Access via COM—things like creating databases, running SQL, editing tables and forms, working with VBA code, and more.

The custom access-dev agent in VS Code runs PowerShell in a terminal and calls the Access functions in Access-POSH.ps1 directly. 

You still get all the “tools” (the 54 functions) without an MCP server. 

Setup:
1. Clone or download the repo
2. Put the two .md files in C:\Users\\%USERNAME%\AppData\Roaming\Code\User\prompts folder (user level access) or create .gituhb\agents folder in your project folder and save the two .md files in the agents folder (workspace level access)
>_**Note:** VS Code detects any .md files in the .github/agents folder of your workspace as custom agents_
4. Replace the path in the .md files to the location of the **Access-POSH.ps1** script on your computer
5. Select access-dev as the agent

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
