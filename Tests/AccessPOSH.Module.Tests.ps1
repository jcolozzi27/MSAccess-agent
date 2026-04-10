# Tests/AccessPOSH.Module.Tests.ps1
# Pester 5+ tests — module loading, function exports, file structure

# PSScriptAnalyzer doesn't understand Pester's BeforeAll/It scoping
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
}

Describe 'AccessPOSH Module' {

    Context 'Module manifest' {
        It 'Manifest file exists' {
            Test-Path $modulePath | Should -BeTrue
        }

        It 'Manifest is valid' {
            { Test-ModuleManifest -Path $modulePath -ErrorAction Stop } | Should -Not -Throw
        }

        It 'Manifest has correct RootModule' {
            $manifest = Test-ModuleManifest -Path $modulePath
            $manifest.RootModule | Should -Be 'AccessPOSH.psm1'
        }
    }

    Context 'Module loads' {
        BeforeAll {
            # Remove if already loaded
            Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
        }

        AfterAll {
            Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        It 'Module is loaded' {
            Get-Module AccessPOSH | Should -Not -BeNullOrEmpty
        }

        It 'Module version is 1.0.0' {
            (Get-Module AccessPOSH).Version.ToString() | Should -Be '1.0.0'
        }

        It 'Exports exactly 91 public functions' {
            $exported = (Get-Module AccessPOSH).ExportedFunctions.Keys
            $exported.Count | Should -Be 91
        }
    }

    Context 'Expected function exports' {
        BeforeAll {
            Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
            $script:exported = (Get-Module AccessPOSH).ExportedFunctions.Keys
        }

        AfterAll {
            Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        It "Exports <_>" -ForEach @(
            # DatabaseOps (11)
            'Close-AccessDatabase', 'New-AccessDatabase', 'Repair-AccessDatabase',
            'Invoke-AccessDecompile', 'Get-AccessObject', 'Get-AccessCode',
            'Set-AccessCode', 'Remove-AccessObject', 'Export-AccessStructure',
            'Invoke-AccessSQL', 'Invoke-AccessSQLBatch',
            # TableOps (7)
            'Get-AccessTableInfo', 'New-AccessTable', 'Edit-AccessTable',
            'Get-AccessFieldProperty', 'Set-AccessFieldProperty',
            'Get-AccessIndex', 'Set-AccessIndex',
            # VbeOps (15)
            'Get-AccessVbeLine', 'Get-AccessVbeProc', 'Get-AccessVbeModuleInfo',
            'Set-AccessVbeLine', 'Set-AccessVbeProc', 'Update-AccessVbeProc',
            'Add-AccessVbeCode', 'Find-AccessVbeText', 'Search-AccessVbe',
            'Search-AccessQuery', 'Find-AccessUsage', 'Invoke-AccessMacro',
            'Invoke-AccessVba', 'Invoke-AccessEval', 'Test-AccessVbaCompile',
            # FormReportOps (9)
            'New-AccessForm', 'Get-AccessFormProperty', 'Set-AccessFormProperty',
            'Get-AccessControl', 'Get-AccessControlDetail', 'New-AccessControl',
            'Remove-AccessControl', 'Set-AccessControlProperty', 'Set-AccessControlBatch',
            # MetadataOps (12)
            'Get-AccessLinkedTable', 'Set-AccessLinkedTable', 'Get-AccessRelationship',
            'New-AccessRelationship', 'Remove-AccessRelationship', 'Get-AccessReference',
            'Set-AccessReference', 'Set-AccessQuery', 'Get-AccessStartupOption',
            'Get-AccessDatabaseProperty', 'Set-AccessDatabaseProperty', 'Get-AccessTip',
            # ExportTransfer (2)
            'Export-AccessReport', 'Copy-AccessData',
            # UIAutomation (3)
            'Get-AccessScreenshot', 'Send-AccessClick', 'Send-AccessKeyboard',
            # TempVarOps (3)
            'Get-AccessTempVar', 'Set-AccessTempVar', 'Remove-AccessTempVar',
            # ImportOps (5)
            'Import-AccessFromExcel', 'Import-AccessFromCSV', 'Import-AccessFromXML',
            'Import-AccessFromDatabase', 'Export-AccessToExcel',
            # SecurityOps (4)
            'Test-AccessDatabasePassword', 'Set-AccessDatabasePassword',
            'Remove-AccessDatabasePassword', 'Get-AccessDatabaseEncryption',
            # ReportOps (4)
            'New-AccessReport', 'Get-AccessGroupLevel', 'Set-AccessGroupLevel',
            'Remove-AccessGroupLevel',
            # SubDataSheetOps (2)
            'Get-AccessSubDataSheet', 'Set-AccessSubDataSheet',
            # NavigationPaneOps (3)
            'Show-AccessNavigationPane', 'Hide-AccessNavigationPane', 'Set-AccessNavigationPaneLock',
            # RibbonOps (3)
            'Get-AccessRibbon', 'Set-AccessRibbon', 'Remove-AccessRibbon',
            # ApplicationOps (3)
            'Get-AccessApplicationInfo', 'Test-AccessRuntime', 'Get-AccessFileInfo',
            # ThemeOps (3)
            'Get-AccessTheme', 'Set-AccessTheme', 'Get-AccessThemeList',
            # PrintOps (2)
            'Export-AccessFilteredReport', 'Send-AccessReportToPrinter'
        ) {
            $script:exported | Should -Contain $_
        }
    }

    Context 'No private functions leaked' {
        BeforeAll {
            Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
            $script:exported = (Get-Module AccessPOSH).ExportedFunctions.Keys
        }

        AfterAll {
            Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        It "Does NOT export private function <_>" -ForEach @(
            'Test-AccessAlive', 'Get-AccessHwnd', 'Set-AccessVisibleBestEffort',
            'Clear-AccessCaches', 'Connect-AccessDB', 'ConvertTo-SafeValue',
            'ConvertTo-CoercedProp', 'Format-AccessOutput', 'Read-TempFile',
            'Write-TempFile', 'Remove-BinarySections', 'Get-BinaryBlocks',
            'Restore-BinarySections', 'Split-CodeBehind', 'Set-FieldProperty',
            'Invoke-VbaAfterImport', 'Test-TextMatch', 'Get-CodeModule',
            'Get-AllModuleCode', 'Test-WsNormalizedMatch', 'Get-ClosestMatchContext',
            'Open-InDesignView', 'Get-DesignObject', 'Save-AndCloseDesign',
            'ConvertFrom-ControlBlock', 'Get-ParsedControls'
        ) {
            $script:exported | Should -Not -Contain $_
        }
    }

    Context 'File structure' {
        It 'Has Private folder' {
            Test-Path (Join-Path $PSScriptRoot '..\AccessPOSH\Private') | Should -BeTrue
        }

        It 'Has Public folder' {
            Test-Path (Join-Path $PSScriptRoot '..\AccessPOSH\Public') | Should -BeTrue
        }

        It 'Has 5 Private .ps1 files' {
            $files = Get-ChildItem (Join-Path $PSScriptRoot '..\AccessPOSH\Private\*.ps1')
            $files.Count | Should -Be 5
        }

        It 'Has 17 Public .ps1 files' {
            $files = Get-ChildItem (Join-Path $PSScriptRoot '..\AccessPOSH\Public\*.ps1')
            $files.Count | Should -Be 17
        }
    }

    Context 'C# Add-Type classes' {
        BeforeAll {
            Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
        }

        AfterAll {
            Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        It 'AccessPoshNative type is loaded' {
            ([System.Management.Automation.PSTypeName]'AccessPoshNative').Type | Should -Not -BeNullOrEmpty
        }

        It 'AccessPoshUI type is loaded' {
            ([System.Management.Automation.PSTypeName]'AccessPoshUI').Type | Should -Not -BeNullOrEmpty
        }
    }
}
