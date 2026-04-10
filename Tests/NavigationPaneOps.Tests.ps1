# Tests/NavigationPaneOps.Tests.ps1
# Parameter-validation tests for NavigationPaneOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Show-AccessNavigationPane' {
    It 'Has CmdletBinding' {
        (Get-Command Show-AccessNavigationPane).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Show-AccessNavigationPane
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Show-AccessNavigationPane).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}

Describe 'Hide-AccessNavigationPane' {
    It 'Has CmdletBinding' {
        (Get-Command Hide-AccessNavigationPane).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Hide-AccessNavigationPane
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Hide-AccessNavigationPane).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessNavigationPaneLock' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessNavigationPaneLock).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Set-AccessNavigationPaneLock
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -Locked is omitted' {
        { Set-AccessNavigationPaneLock -DbPath 'x:\fake.accdb' } | Should -Throw '*-Locked is required*'
    }
    It 'Has AsJson switch' {
        (Get-Command Set-AccessNavigationPaneLock).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
