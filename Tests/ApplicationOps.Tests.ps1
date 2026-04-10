# Tests/ApplicationOps.Tests.ps1
# Parameter-validation tests for ApplicationOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessApplicationInfo' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessApplicationInfo).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Get-AccessApplicationInfo
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Get-AccessApplicationInfo).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}

Describe 'Test-AccessRuntime' {
    It 'Has CmdletBinding' {
        (Get-Command Test-AccessRuntime).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Test-AccessRuntime
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Test-AccessRuntime).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Get-AccessFileInfo' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessFileInfo).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Get-AccessFileInfo
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Get-AccessFileInfo).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}
