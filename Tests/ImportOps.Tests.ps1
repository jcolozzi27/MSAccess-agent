# Tests/ImportOps.Tests.ps1
# Parameter-validation tests for ImportOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Import-AccessFromExcel' {
    It 'Has CmdletBinding' {
        (Get-Command Import-AccessFromExcel).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Import-AccessFromExcel
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -ExcelPath is omitted' {
        { Import-AccessFromExcel -DbPath 'x:\fake.accdb' -TableName 'T' } | Should -Throw '*-ExcelPath is required*'
    }
    It 'Throws when -TableName is omitted' {
        { Import-AccessFromExcel -DbPath 'x:\fake.accdb' -ExcelPath 'x:\fake.xlsx' } | Should -Throw '*-TableName is required*'
    }
    It 'Has SpreadsheetType with ValidateSet' {
        $p = (Get-Command Import-AccessFromExcel).Parameters['SpreadsheetType']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'xlsx'
        $vs[0].ValidValues | Should -Contain 'xls'
    }
    It 'Has HasFieldNames switch' {
        $p = (Get-Command Import-AccessFromExcel).Parameters['HasFieldNames']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Import-AccessFromExcel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Import-AccessFromCSV' {
    It 'Has CmdletBinding' {
        (Get-Command Import-AccessFromCSV).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Import-AccessFromCSV
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -FilePath is omitted' {
        { Import-AccessFromCSV -DbPath 'x:\fake.accdb' -TableName 'T' } | Should -Throw '*-FilePath is required*'
    }
    It 'Throws when -TableName is omitted' {
        { Import-AccessFromCSV -DbPath 'x:\fake.accdb' -FilePath 'x:\fake.csv' } | Should -Throw '*-TableName is required*'
    }
    It 'Has SpecificationName parameter (optional)' {
        (Get-Command Import-AccessFromCSV).Parameters['SpecificationName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Import-AccessFromCSV).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Import-AccessFromXML' {
    It 'Has CmdletBinding' {
        (Get-Command Import-AccessFromXML).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Import-AccessFromXML
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -XmlPath is omitted' {
        { Import-AccessFromXML -DbPath 'x:\fake.accdb' } | Should -Throw '*-XmlPath is required*'
    }
    It 'Has ImportOptions with ValidateSet' {
        $p = (Get-Command Import-AccessFromXML).Parameters['ImportOptions']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'structureanddata'
    }
    It 'Has AsJson switch' {
        (Get-Command Import-AccessFromXML).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Import-AccessFromDatabase' {
    It 'Has CmdletBinding' {
        (Get-Command Import-AccessFromDatabase).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Import-AccessFromDatabase
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -SourceDbPath is omitted' {
        { Import-AccessFromDatabase -DbPath 'x:\fake.accdb' -SourceObject 'T' } | Should -Throw '*-SourceDbPath is required*'
    }
    It 'Throws when -SourceObject is omitted' {
        { Import-AccessFromDatabase -DbPath 'x:\fake.accdb' -SourceDbPath 'x:\source.accdb' } | Should -Throw '*-SourceObject is required*'
    }
    It 'Has ObjectType with ValidateSet' {
        $p = (Get-Command Import-AccessFromDatabase).Parameters['ObjectType']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'table'
        $vs[0].ValidValues | Should -Contain 'query'
    }
    It 'Has AsJson switch' {
        (Get-Command Import-AccessFromDatabase).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Export-AccessToExcel' {
    It 'Has CmdletBinding' {
        (Get-Command Export-AccessToExcel).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Export-AccessToExcel
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -ObjectName is omitted' {
        { Export-AccessToExcel -DbPath 'x:\fake.accdb' -ExcelPath 'x:\out.xlsx' } | Should -Throw '*-ObjectName is required*'
    }
    It 'Throws when -ExcelPath is omitted' {
        { Export-AccessToExcel -DbPath 'x:\fake.accdb' -ObjectName 'T' } | Should -Throw '*-ExcelPath is required*'
    }
    It 'Has SpreadsheetType with ValidateSet' {
        $p = (Get-Command Export-AccessToExcel).Parameters['SpreadsheetType']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'xlsx'
    }
    It 'Has AsJson switch' {
        (Get-Command Export-AccessToExcel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
