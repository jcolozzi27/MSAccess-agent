# Tests/DatabaseOps.Tests.ps1
# Parameter-validation tests for DatabaseOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Close-AccessDatabase' {
    It 'Has CmdletBinding' {
        (Get-Command Close-AccessDatabase).CmdletBinding | Should -BeTrue
    }
}

Describe 'New-AccessDatabase' {
    It 'Has CmdletBinding' {
        (Get-Command New-AccessDatabase).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter' {
        (Get-Command New-AccessDatabase).Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Repair-AccessDatabase' {
    It 'Has CmdletBinding' {
        (Get-Command Repair-AccessDatabase).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter' {
        (Get-Command Repair-AccessDatabase).Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Invoke-AccessDecompile' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessDecompile).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter' {
        (Get-Command Invoke-AccessDecompile).Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessObject' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessObject).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectType parameter' {
        (Get-Command Get-AccessObject).Parameters['ObjectType'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessCode' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessCode).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command Get-AccessCode).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-AccessCode' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessCode).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command Set-AccessCode).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Code parameter' {
        (Get-Command Set-AccessCode).Parameters['Code'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-AccessObject' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessObject).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command Remove-AccessObject).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ObjectType parameter' {
        (Get-Command Remove-AccessObject).Parameters['ObjectType'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Export-AccessStructure' {
    It 'Has CmdletBinding' {
        (Get-Command Export-AccessStructure).CmdletBinding | Should -BeTrue
    }
    It 'Has OutputPath parameter' {
        (Get-Command Export-AccessStructure).Parameters['OutputPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Invoke-AccessSQL' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessSQL).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -SQL is omitted' {
        { Invoke-AccessSQL -DbPath 'x:\fake.accdb' } | Should -Throw '*-SQL is required*'
    }
}

Describe 'Invoke-AccessSQLBatch' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessSQLBatch).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -Statements is omitted' {
        { Invoke-AccessSQLBatch -DbPath 'x:\fake.accdb' } | Should -Throw '*-Statements is required*'
    }
}
