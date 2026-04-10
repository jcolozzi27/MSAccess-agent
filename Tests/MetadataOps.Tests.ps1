# Tests/MetadataOps.Tests.ps1
# Parameter-validation tests for MetadataOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessLinkedTable' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessLinkedTable).CmdletBinding | Should -BeTrue
    }
}

Describe 'Set-AccessLinkedTable' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessLinkedTable).CmdletBinding | Should -BeTrue
    }
    It 'Has TableName parameter' {
        (Get-Command Set-AccessLinkedTable).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has NewConnect parameter' {
        (Get-Command Set-AccessLinkedTable).Parameters['NewConnect'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessRelationship' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessRelationship).CmdletBinding | Should -BeTrue
    }
}

Describe 'New-AccessRelationship' {
    It 'Has CmdletBinding' {
        (Get-Command New-AccessRelationship).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command New-AccessRelationship).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Table parameter' {
        (Get-Command New-AccessRelationship).Parameters['Table'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ForeignTable parameter' {
        (Get-Command New-AccessRelationship).Parameters['ForeignTable'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Fields parameter' {
        (Get-Command New-AccessRelationship).Parameters['Fields'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-AccessRelationship' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessRelationship).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -Name is omitted' {
        { Remove-AccessRelationship -DbPath 'x:\fake.accdb' } | Should -Throw '*-Name is required*'
    }
}

Describe 'Get-AccessReference' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessReference).CmdletBinding | Should -BeTrue
    }
}

Describe 'Set-AccessReference' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessReference).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -Action is omitted' {
        { Set-AccessReference -DbPath 'x:\fake.accdb' } | Should -Throw '*-Action is required*'
    }
}

Describe 'Set-AccessQuery' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessQuery).CmdletBinding | Should -BeTrue
    }
    It 'Has QueryName parameter' {
        (Get-Command Set-AccessQuery).Parameters['QueryName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Sql parameter' {
        (Get-Command Set-AccessQuery).Parameters['Sql'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessStartupOption' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessStartupOption).CmdletBinding | Should -BeTrue
    }
}

Describe 'Get-AccessDatabaseProperty' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessDatabaseProperty).CmdletBinding | Should -BeTrue
    }
}

Describe 'Set-AccessDatabaseProperty' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessDatabaseProperty).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command Set-AccessDatabaseProperty).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Value parameter' {
        (Get-Command Set-AccessDatabaseProperty).Parameters['Value'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessTip' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessTip).CmdletBinding | Should -BeTrue
    }
}
