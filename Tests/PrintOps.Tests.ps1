# Tests/PrintOps.Tests.ps1
# Parameter-validation tests for PrintOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Export-AccessFilteredReport' {
    It 'Has CmdletBinding' {
        (Get-Command Export-AccessFilteredReport).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Export-AccessFilteredReport
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -ReportName is omitted' {
        { Export-AccessFilteredReport -DbPath 'x:\fake.accdb' } | Should -Throw '*-ReportName is required*'
    }
    It 'Has WhereCondition parameter (optional)' {
        (Get-Command Export-AccessFilteredReport).Parameters['WhereCondition'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FilterName parameter (optional)' {
        (Get-Command Export-AccessFilteredReport).Parameters['FilterName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has OutputFormat with ValidateSet' {
        $p = (Get-Command Export-AccessFilteredReport).Parameters['OutputFormat']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'pdf'
        $vs[0].ValidValues | Should -Contain 'xlsx'
    }
    It 'Has OutputPath parameter (optional)' {
        (Get-Command Export-AccessFilteredReport).Parameters['OutputPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has OpenAfterExport switch' {
        (Get-Command Export-AccessFilteredReport).Parameters['OpenAfterExport'].SwitchParameter | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Export-AccessFilteredReport).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Send-AccessReportToPrinter' {
    It 'Has CmdletBinding' {
        (Get-Command Send-AccessReportToPrinter).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (optional, session fallback)' {
        $cmd = Get-Command Send-AccessReportToPrinter
        $cmd.Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -ReportName is omitted' {
        { Send-AccessReportToPrinter -DbPath 'x:\fake.accdb' } | Should -Throw '*-ReportName is required*'
    }
    It 'Has WhereCondition parameter (optional)' {
        (Get-Command Send-AccessReportToPrinter).Parameters['WhereCondition'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Copies parameter' {
        (Get-Command Send-AccessReportToPrinter).Parameters['Copies'] | Should -Not -BeNullOrEmpty
    }
    It 'Has PrintRange with ValidateSet' {
        $p = (Get-Command Send-AccessReportToPrinter).Parameters['PrintRange']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'all'
        $vs[0].ValidValues | Should -Contain 'pages'
    }
    It 'Has FromPage parameter' {
        (Get-Command Send-AccessReportToPrinter).Parameters['FromPage'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ToPage parameter' {
        (Get-Command Send-AccessReportToPrinter).Parameters['ToPage'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Send-AccessReportToPrinter).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
