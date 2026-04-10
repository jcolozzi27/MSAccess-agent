# Tests/SubDataSheetOps.Tests.ps1
# Parameter-validation tests for SubDataSheetOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessSubDataSheet' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessSubDataSheet).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Get-AccessSubDataSheet
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -TableName is omitted' {
        { Get-AccessSubDataSheet -DbPath 'x:\fake.accdb' } | Should -Throw '*-TableName is required*'
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Get-AccessSubDataSheet).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessSubDataSheet' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessSubDataSheet).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Set-AccessSubDataSheet
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -TableName is omitted' {
        { Set-AccessSubDataSheet -DbPath 'x:\fake.accdb' } | Should -Throw '*-TableName is required*'
    }
    It 'Has SubDataSheetName parameter' {
        (Get-Command Set-AccessSubDataSheet).Parameters['SubDataSheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has LinkChildFields parameter' {
        (Get-Command Set-AccessSubDataSheet).Parameters['LinkChildFields'] | Should -Not -BeNullOrEmpty
    }
    It 'Has LinkMasterFields parameter' {
        (Get-Command Set-AccessSubDataSheet).Parameters['LinkMasterFields'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Height parameter' {
        (Get-Command Set-AccessSubDataSheet).Parameters['Height'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Set-AccessSubDataSheet).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}
