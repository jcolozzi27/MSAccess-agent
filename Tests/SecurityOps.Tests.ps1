# Tests/SecurityOps.Tests.ps1
# Parameter-validation tests for SecurityOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Test-AccessDatabasePassword' {
    It 'Has CmdletBinding' {
        (Get-Command Test-AccessDatabasePassword).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Test-AccessDatabasePassword
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Test-AccessDatabasePassword).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessDatabasePassword' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessDatabasePassword).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Set-AccessDatabasePassword
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -NewPassword is omitted' {
        { Set-AccessDatabasePassword -DbPath 'x:\fake.accdb' } | Should -Throw '*-NewPassword is required*'
    }
    It 'Has OldPassword parameter (optional)' {
        (Get-Command Set-AccessDatabasePassword).Parameters['OldPassword'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Set-AccessDatabasePassword).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Remove-AccessDatabasePassword' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessDatabasePassword).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Remove-AccessDatabasePassword
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -CurrentPassword is omitted' {
        { Remove-AccessDatabasePassword -DbPath 'x:\fake.accdb' } | Should -Throw '*-CurrentPassword is required*'
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-AccessDatabasePassword).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Get-AccessDatabaseEncryption' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessDatabaseEncryption).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Get-AccessDatabaseEncryption
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Get-AccessDatabaseEncryption).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}
