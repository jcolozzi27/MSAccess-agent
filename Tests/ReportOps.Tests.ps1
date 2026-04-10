# Tests/ReportOps.Tests.ps1
# Parameter-validation tests for ReportOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'New-AccessReport' {
    It 'Has CmdletBinding' {
        (Get-Command New-AccessReport).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command New-AccessReport
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -ReportName is omitted' {
        { New-AccessReport -DbPath 'x:\fake.accdb' } | Should -Throw '*-ReportName is required*'
    }
    It 'Has RecordSource parameter (optional)' {
        (Get-Command New-AccessReport).Parameters['RecordSource'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command New-AccessReport).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Get-AccessGroupLevel' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessGroupLevel).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Get-AccessGroupLevel
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -ReportName is omitted' {
        { Get-AccessGroupLevel -DbPath 'x:\fake.accdb' } | Should -Throw '*-ReportName is required*'
    }
    It 'Has AsJson switch' {
        (Get-Command Get-AccessGroupLevel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessGroupLevel' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessGroupLevel).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Set-AccessGroupLevel
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -ReportName is omitted' {
        { Set-AccessGroupLevel -DbPath 'x:\fake.accdb' } | Should -Throw '*-ReportName is required*'
    }
    It 'Throws when -Expression is omitted' {
        { Set-AccessGroupLevel -DbPath 'x:\fake.accdb' -ReportName 'R' } | Should -Throw '*-Expression is required*'
    }
    It 'Has SortOrder with ValidateSet' {
        $p = (Get-Command Set-AccessGroupLevel).Parameters['SortOrder']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'ascending'
        $vs[0].ValidValues | Should -Contain 'descending'
    }
    It 'Has GroupHeader switch' {
        (Get-Command Set-AccessGroupLevel).Parameters['GroupHeader'].SwitchParameter | Should -BeTrue
    }
    It 'Has GroupFooter switch' {
        (Get-Command Set-AccessGroupLevel).Parameters['GroupFooter'].SwitchParameter | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Set-AccessGroupLevel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Remove-AccessGroupLevel' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessGroupLevel).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Remove-AccessGroupLevel
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -ReportName is omitted' {
        { Remove-AccessGroupLevel -DbPath 'x:\fake.accdb' } | Should -Throw '*-ReportName is required*'
    }
    It 'Throws when -LevelIndex is omitted' {
        { Remove-AccessGroupLevel -DbPath 'x:\fake.accdb' -ReportName 'R' } | Should -Throw '*-LevelIndex is required*'
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-AccessGroupLevel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
