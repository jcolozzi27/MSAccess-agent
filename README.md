# MSAccess-agent

Access-POSH.ps1 is a PowerShell port of [unmateria/MCP-Access](https://github.com/unmateria/MCP-Access) that exposes 50+ commands for automating Microsoft Access (creating databases, running SQL, editing forms, VBA, etc.) using COM. Instead of running a separate MCP server, a custom access-dev agent in VS Code calls these PowerShell functions directly in a terminal.

Setup:
1. Clone or download the repo
2. Put the two .md files in C:\Users\\%USERNAME%\AppData\Roaming\Code\User\prompts folder (user level access) or create a .github\agents folder in your project folder and save the two .md files in the agent folder 
>_**Note:** VS Code detects any .md files in the .github/agents folder of your workspace as custom agents_
3. Replace the path in the .md files to the location of the **AccessPOSH.psd1** module on your computer
4. Select access-dev from the agent picker before prompting

## Available Functions (91 public)

| Category | Functions |
|----------|-----------|
| **Database** | `New-AccessDatabase`, `Close-AccessDatabase`, `Repair-AccessDatabase`, `Invoke-AccessDecompile` |
| **Objects** | `Get-AccessObject`, `Get-AccessCode`, `Set-AccessCode`, `Remove-AccessObject`, `Export-AccessStructure` |
| **SQL** | `Invoke-AccessSQL`, `Invoke-AccessSQLBatch` |
| **Tables** | `Get-AccessTableInfo`, `New-AccessTable`, `Edit-AccessTable` |
| **VBE** | `Get-AccessVbeLine`, `Get-AccessVbeProc`, `Get-AccessVbeModuleInfo`, `Set-AccessVbeLine`, `Set-AccessVbeProc`, `Update-AccessVbeProc`, `Add-AccessVbeCode` |
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
| **TempVars** | `Get-AccessTempVar`, `Set-AccessTempVar`, `Remove-AccessTempVar` |
| **Import** | `Import-AccessFromExcel`, `Import-AccessFromCSV`, `Import-AccessFromXML`, `Import-AccessFromDatabase`, `Export-AccessToExcel` |
| **Security** | `Test-AccessDatabasePassword`, `Set-AccessDatabasePassword`, `Remove-AccessDatabasePassword`, `Get-AccessDatabaseEncryption` |
| **Reports** | `New-AccessReport`, `Get-AccessGroupLevel`, `Set-AccessGroupLevel`, `Remove-AccessGroupLevel` |
| **SubDataSheets** | `Get-AccessSubDataSheet`, `Set-AccessSubDataSheet` |
| **Navigation Pane** | `Show-AccessNavigationPane`, `Hide-AccessNavigationPane`, `Set-AccessNavigationPaneLock` |
| **Ribbon** | `Get-AccessRibbon`, `Set-AccessRibbon`, `Remove-AccessRibbon` |
| **Application** | `Get-AccessApplicationInfo`, `Test-AccessRuntime`, `Get-AccessFileInfo` |
| **Themes** | `Get-AccessTheme`, `Set-AccessTheme`, `Get-AccessThemeList` |
| **Print** | `Export-AccessFilteredReport`, `Send-AccessReportToPrinter` |
