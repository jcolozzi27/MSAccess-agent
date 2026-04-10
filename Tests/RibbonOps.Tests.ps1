# Tests/RibbonOps.Tests.ps1
# Parameter-validation tests for RibbonOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessRibbon' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessRibbon).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Get-AccessRibbon).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has RibbonName parameter (optional)' {
        (Get-Command Get-AccessRibbon).Parameters['RibbonName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Get-AccessRibbon).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessRibbon' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessRibbon).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Set-AccessRibbon).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has RibbonName parameter (mandatory)' {
        $p = (Get-Command Set-AccessRibbon).Parameters['RibbonName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has RibbonXml parameter (mandatory)' {
        $p = (Get-Command Set-AccessRibbon).Parameters['RibbonXml']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has SetAsDefault switch' {
        (Get-Command Set-AccessRibbon).Parameters['SetAsDefault'].SwitchParameter | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Set-AccessRibbon).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Remove-AccessRibbon' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessRibbon).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Remove-AccessRibbon).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has RibbonName parameter (mandatory)' {
        $p = (Get-Command Remove-AccessRibbon).Parameters['RibbonName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-AccessRibbon).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
