# Tests/ThemeOps.Tests.ps1
# Parameter-validation tests for ThemeOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessTheme' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessTheme).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Get-AccessTheme
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -ObjectName is omitted' {
        { Get-AccessTheme -DbPath 'x:\fake.accdb' } | Should -Throw '*-ObjectName is required*'
    }
    It 'Has ObjectType with ValidateSet' {
        $p = (Get-Command Get-AccessTheme).Parameters['ObjectType']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'form'
        $vs[0].ValidValues | Should -Contain 'report'
    }
    It 'Has AsJson switch' {
        (Get-Command Get-AccessTheme).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessTheme' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessTheme).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Set-AccessTheme
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -ObjectName is omitted' {
        { Set-AccessTheme -DbPath 'x:\fake.accdb' -ThemeName 'T' } | Should -Throw '*-ObjectName is required*'
    }
    It 'Throws when -ThemeName is omitted' {
        { Set-AccessTheme -DbPath 'x:\fake.accdb' -ObjectName 'F' } | Should -Throw '*-ThemeName is required*'
    }
    It 'Has ObjectType with ValidateSet' {
        $p = (Get-Command Set-AccessTheme).Parameters['ObjectType']
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs[0].ValidValues | Should -Contain 'form'
        $vs[0].ValidValues | Should -Contain 'report'
    }
    It 'Has AsJson switch' {
        (Get-Command Set-AccessTheme).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Get-AccessThemeList' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessThemeList).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Get-AccessThemeList
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Get-AccessThemeList).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
