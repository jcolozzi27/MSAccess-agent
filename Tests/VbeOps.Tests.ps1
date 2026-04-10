# Tests/VbeOps.Tests.ps1
# Parameter-validation tests for VbeOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessVbeLine' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessVbeLine).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Get-AccessVbeLine).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessVbeProc' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessVbeProc).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Get-AccessVbeProc).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ProcName parameter' {
        (Get-Command Get-AccessVbeProc).Parameters['ProcName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessVbeModuleInfo' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessVbeModuleInfo).CmdletBinding | Should -BeTrue
    }
}

Describe 'Set-AccessVbeLine' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessVbeLine).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Set-AccessVbeLine).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has StartLine parameter' {
        (Get-Command Set-AccessVbeLine).Parameters['StartLine'] | Should -Not -BeNullOrEmpty
    }
    It 'Has NewCode parameter' {
        (Get-Command Set-AccessVbeLine).Parameters['NewCode'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-AccessVbeProc' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessVbeProc).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Set-AccessVbeProc).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ProcName parameter' {
        (Get-Command Set-AccessVbeProc).Parameters['ProcName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has NewCode parameter' {
        (Get-Command Set-AccessVbeProc).Parameters['NewCode'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Update-AccessVbeProc' {
    It 'Has CmdletBinding' {
        (Get-Command Update-AccessVbeProc).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Update-AccessVbeProc).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ProcName parameter' {
        (Get-Command Update-AccessVbeProc).Parameters['ProcName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Patches parameter' {
        (Get-Command Update-AccessVbeProc).Parameters['Patches'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Add-AccessVbeCode' {
    It 'Has CmdletBinding' {
        (Get-Command Add-AccessVbeCode).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Add-AccessVbeCode).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Code parameter' {
        (Get-Command Add-AccessVbeCode).Parameters['Code'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Find-AccessVbeText' {
    It 'Has CmdletBinding' {
        (Get-Command Find-AccessVbeText).CmdletBinding | Should -BeTrue
    }
    It 'Has SearchText parameter' {
        (Get-Command Find-AccessVbeText).Parameters['SearchText'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Search-AccessVbe' {
    It 'Has CmdletBinding' {
        (Get-Command Search-AccessVbe).CmdletBinding | Should -BeTrue
    }
    It 'Has SearchText parameter' {
        (Get-Command Search-AccessVbe).Parameters['SearchText'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Search-AccessQuery' {
    It 'Has CmdletBinding' {
        (Get-Command Search-AccessQuery).CmdletBinding | Should -BeTrue
    }
    It 'Has SearchText parameter' {
        (Get-Command Search-AccessQuery).Parameters['SearchText'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Find-AccessUsage' {
    It 'Has CmdletBinding' {
        (Get-Command Find-AccessUsage).CmdletBinding | Should -BeTrue
    }
    It 'Has SearchText parameter' {
        (Get-Command Find-AccessUsage).Parameters['SearchText'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Invoke-AccessMacro' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessMacro).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -MacroName is omitted' {
        { Invoke-AccessMacro -DbPath 'x:\fake.accdb' } | Should -Throw '*-MacroName is required*'
    }
}

Describe 'Invoke-AccessVba' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessVba).CmdletBinding | Should -BeTrue
    }
    It 'Has Procedure parameter' {
        (Get-Command Invoke-AccessVba).Parameters['Procedure'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Invoke-AccessEval' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessEval).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -Expression is omitted' {
        { Invoke-AccessEval -DbPath 'x:\fake.accdb' } | Should -Throw '*-Expression is required*'
    }
}

Describe 'Test-AccessVbaCompile' {
    It 'Has CmdletBinding' {
        (Get-Command Test-AccessVbaCompile).CmdletBinding | Should -BeTrue
    }
}

Describe 'Import-AccessVbaFile' {
    It 'Has CmdletBinding' {
        (Get-Command Import-AccessVbaFile).CmdletBinding | Should -BeTrue
    }
    It 'Has FilePath parameter' {
        (Get-Command Import-AccessVbaFile).Parameters['FilePath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Force parameter' {
        (Get-Command Import-AccessVbaFile).Parameters['Force'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Import-AccessVbaFile).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -FilePath is omitted' {
        { Import-AccessVbaFile -DbPath 'x:\fake.accdb' } | Should -Throw '*-FilePath is required*'
    }
    It 'Throws for non-existent file' {
        { Import-AccessVbaFile -DbPath 'x:\fake.accdb' -FilePath 'x:\nonexistent.cls' } | Should -Throw '*File not found*'
    }
    It 'Throws for unsupported extension' {
        $tmp = [System.IO.Path]::GetTempFileName()
        try {
            $txtFile = $tmp -replace '\.tmp$', '.txt'
            Rename-Item $tmp $txtFile
            { Import-AccessVbaFile -DbPath 'x:\fake.accdb' -FilePath $txtFile } | Should -Throw '*Only .bas and .cls*'
        } finally {
            Remove-Item $txtFile -Force -ErrorAction SilentlyContinue
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Test-AccessVbaFileEncoding' {
    It 'Has CmdletBinding' {
        (Get-Command Test-AccessVbaFileEncoding).CmdletBinding | Should -BeTrue
    }
    It 'Has FilePath parameter' {
        (Get-Command Test-AccessVbaFileEncoding).Parameters['FilePath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Test-AccessVbaFileEncoding).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when -FilePath is omitted' {
        { Test-AccessVbaFileEncoding } | Should -Throw '*-FilePath is required*'
    }
    It 'Detects ANSI file correctly' {
        $tmp = [System.IO.Path]::Combine($env:TEMP, "test_ansi_$([guid]::NewGuid().ToString('N')).cls")
        try {
            $header = "VERSION 1.0 CLASS`r`nBEGIN`r`n  MultiUse = -1`r`nEND`r`nAttribute VB_Name = ""TestClass""`r`n"
            [System.IO.File]::WriteAllText($tmp, $header, [System.Text.Encoding]::GetEncoding(1252))
            $result = Test-AccessVbaFileEncoding -FilePath $tmp
            $result.is_ansi | Should -BeTrue
            $result.encoding | Should -Be 'ansi'
        } finally {
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
    It 'Detects UTF-8 BOM file correctly' {
        $tmp = [System.IO.Path]::Combine($env:TEMP, "test_utf8bom_$([guid]::NewGuid().ToString('N')).cls")
        try {
            $header = "VERSION 1.0 CLASS`r`nBEGIN`r`n  MultiUse = -1`r`nEND`r`n"
            [System.IO.File]::WriteAllText($tmp, $header, [System.Text.Encoding]::UTF8)
            $result = Test-AccessVbaFileEncoding -FilePath $tmp
            $result.is_ansi | Should -BeFalse
            $result.encoding | Should -Be 'utf-8-bom'
        } finally {
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
    It 'Detects UTF-16 LE file correctly' {
        $tmp = [System.IO.Path]::Combine($env:TEMP, "test_utf16_$([guid]::NewGuid().ToString('N')).cls")
        try {
            $header = "VERSION 1.0 CLASS`r`nBEGIN`r`n  MultiUse = -1`r`nEND`r`n"
            [System.IO.File]::WriteAllText($tmp, $header, [System.Text.Encoding]::Unicode)
            $result = Test-AccessVbaFileEncoding -FilePath $tmp
            $result.is_ansi | Should -BeFalse
            $result.encoding | Should -Be 'utf-16-le'
        } finally {
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}
