# Tests/ExportTransfer.Tests.ps1
# Parameter-validation tests for ExportTransfer functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Export-AccessReport' {
    It 'Has CmdletBinding' {
        (Get-Command Export-AccessReport).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Export-AccessReport).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has OutputPath parameter' {
        (Get-Command Export-AccessReport).Parameters['OutputPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Copy-AccessData' {
    It 'Has CmdletBinding' {
        (Get-Command Copy-AccessData).CmdletBinding | Should -BeTrue
    }
    It 'Has Action parameter' {
        (Get-Command Copy-AccessData).Parameters['Action'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FilePath parameter' {
        (Get-Command Copy-AccessData).Parameters['FilePath'] | Should -Not -BeNullOrEmpty
    }
}
