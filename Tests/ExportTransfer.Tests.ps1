# Tests/ExportTransfer.Tests.ps1
# Parameter-validation tests for ExportTransfer functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Export-AccessReport' {
    It 'Has CmdletBinding' {
        (Get-Command Export-AccessReport).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Export-AccessReport).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has OutputPath parameter' {
        (Get-Command Export-AccessReport).Parameters['OutputPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Copy-AccessData' {
    It 'Has CmdletBinding' {
        (Get-Command Copy-AccessData).CmdletBinding | Should -BeTrue
    }
    It 'Has Action parameter' {
        (Get-Command Copy-AccessData).Parameters['Action'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FilePath parameter' {
        (Get-Command Copy-AccessData).Parameters['FilePath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Export-AccessReport — Parameter Validation' {
    It 'throws when -ObjectName is omitted' {
        { Export-AccessReport -DbPath 'x:\fake.accdb' } | Should -Throw '*-ObjectName is required*'
    }
}

Describe 'Copy-AccessData — Parameter Validation' {
    It 'throws when -Action is omitted' {
        { Copy-AccessData -DbPath 'x:\fake.accdb' } | Should -Throw '*-Action is required*'
    }
    It 'throws when -FilePath is omitted' {
        { Copy-AccessData -DbPath 'x:\fake.accdb' -Action 'import' } | Should -Throw '*-FilePath is required*'
    }
    It 'throws when -TableName is omitted' {
        { Copy-AccessData -DbPath 'x:\fake.accdb' -Action 'import' -FilePath 'x:\fake.xlsx' } | Should -Throw '*-TableName is required*'
    }
}

Describe 'Export-AccessSource' {
    It 'Has CmdletBinding' {
        (Get-Command Export-AccessSource).CmdletBinding | Should -BeTrue
    }
    It 'Has OutputFolder parameter' {
        (Get-Command Export-AccessSource).Parameters['OutputFolder'] | Should -Not -BeNullOrEmpty
    }
    It 'Has DbPath parameter' {
        (Get-Command Export-AccessSource).Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ObjectType parameter' {
        (Get-Command Export-AccessSource).Parameters['ObjectType'] | Should -Not -BeNullOrEmpty
    }
    It 'Has IncludeTableData parameter' {
        (Get-Command Export-AccessSource).Parameters['IncludeTableData'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ClearFirst parameter' {
        (Get-Command Export-AccessSource).Parameters['ClearFirst'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Export-AccessSource).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Export-AccessSource — Parameter Validation' {
    It 'throws when -OutputFolder is omitted' {
        { Export-AccessSource -DbPath 'x:\fake.accdb' } | Should -Throw '*-OutputFolder is required*'
    }
}

Describe 'Import-AccessSource' {
    It 'Has CmdletBinding' {
        (Get-Command Import-AccessSource).CmdletBinding | Should -BeTrue
    }
    It 'Has InputFolder parameter' {
        (Get-Command Import-AccessSource).Parameters['InputFolder'] | Should -Not -BeNullOrEmpty
    }
    It 'Has DbPath parameter' {
        (Get-Command Import-AccessSource).Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ObjectType parameter' {
        (Get-Command Import-AccessSource).Parameters['ObjectType'] | Should -Not -BeNullOrEmpty
    }
    It 'Has OverwriteExisting parameter' {
        (Get-Command Import-AccessSource).Parameters['OverwriteExisting'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Import-AccessSource).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Import-AccessSource — Parameter Validation' {
    It 'throws when -InputFolder is omitted' {
        { Import-AccessSource -DbPath 'x:\fake.accdb' } | Should -Throw '*-InputFolder is required*'
    }
}
