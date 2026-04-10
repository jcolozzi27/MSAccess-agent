# Tests/FormReportOps.Tests.ps1
# Parameter-validation tests for FormReportOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'New-AccessForm' {
    It 'Has CmdletBinding' {
        (Get-Command New-AccessForm).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -FormName is omitted' {
        { New-AccessForm -DbPath 'x:\fake.accdb' } | Should -Throw '*-FormName is required*'
    }
}

Describe 'Get-AccessFormProperty' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessFormProperty).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -ObjectName is omitted' {
        { Get-AccessFormProperty -DbPath 'x:\fake.accdb' } | Should -Throw '*-ObjectName is required*'
    }
}

Describe 'Set-AccessFormProperty' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessFormProperty).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Set-AccessFormProperty).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Properties parameter' {
        (Get-Command Set-AccessFormProperty).Parameters['Properties'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessControl' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessControl).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory ObjectName parameter' {
        (Get-Command Get-AccessControl).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessControlDetail' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessControlDetail).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory ObjectName parameter' {
        (Get-Command Get-AccessControlDetail).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory ControlName parameter' {
        (Get-Command Get-AccessControlDetail).Parameters['ControlName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'New-AccessControl' {
    It 'Has CmdletBinding' {
        (Get-Command New-AccessControl).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command New-AccessControl).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ControlType parameter' {
        (Get-Command New-AccessControl).Parameters['ControlType'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-AccessControl' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessControl).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Remove-AccessControl).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ControlName parameter' {
        (Get-Command Remove-AccessControl).Parameters['ControlName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-AccessControlProperty' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessControlProperty).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Set-AccessControlProperty).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ControlName parameter' {
        (Get-Command Set-AccessControlProperty).Parameters['ControlName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Properties parameter' {
        (Get-Command Set-AccessControlProperty).Parameters['Properties'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-AccessControlBatch' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessControlBatch).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Set-AccessControlBatch).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Controls parameter' {
        (Get-Command Set-AccessControlBatch).Parameters['Controls'] | Should -Not -BeNullOrEmpty
    }
}
