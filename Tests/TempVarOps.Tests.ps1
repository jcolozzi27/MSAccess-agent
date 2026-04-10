# Tests/TempVarOps.Tests.ps1
# Parameter-validation tests for TempVarOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessTempVar' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessTempVar).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Get-AccessTempVar
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Name parameter (optional)' {
        (Get-Command Get-AccessTempVar).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Get-AccessTempVar).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessTempVar' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessTempVar).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Set-AccessTempVar
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -Name is omitted' {
        { Set-AccessTempVar -DbPath 'x:\fake.accdb' } | Should -Throw '*-Name is required*'
    }
    It 'Throws when -Value is omitted' {
        { Set-AccessTempVar -DbPath 'x:\fake.accdb' -Name 'x' } | Should -Throw '*-Value is required*'
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Set-AccessTempVar).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}

Describe 'Remove-AccessTempVar' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessTempVar).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Remove-AccessTempVar
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Name parameter (optional)' {
        (Get-Command Remove-AccessTempVar).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Remove-AccessTempVar).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}
