<#
.SYNOPSIS
    Access-POSH — PowerShell Access Database Automation

.DESCRIPTION
    Provides full COM automation of Microsoft Access databases (.accdb/.mdb).
    Port of the Python MCP-Access server (54+ tools) to native PowerShell.
    No MCP server needed — AI agents call functions directly via terminal.

    Usage:
        . .\Access-POSH.ps1                           # dot-source
        Invoke-AccessSQL -DbPath "C:\my.accdb" -SQL "SELECT * FROM Users"
        Close-AccessDatabase                           # release COM

.NOTES
    Requires: Windows + Microsoft Access (full install, not Runtime)
    PowerShell: 5.1+ or PowerShell 7+
#>

# ═══════════════════════════════════════════════════════════════════════════
# NATIVE INTEROP (Win32 API via C# Add-Type)
# ═══════════════════════════════════════════════════════════════════════════

if (-not ([System.Management.Automation.PSTypeName]'AccessPoshNative').Type) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class AccessPoshNative
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
}
'@
}

# ═══════════════════════════════════════════════════════════════════════════
# CONSTANTS & TYPE MAPS
# ═══════════════════════════════════════════════════════════════════════════

# Access object type numbers for DoCmd operations
$script:AC_TYPE = @{
    query  = 1   # acQuery
    form   = 2   # acForm
    report = 3   # acReport
    macro  = 4   # acMacro
    module = 5   # acModule
}

# DAO field type mapping: friendly name → DAO type number
$script:FIELD_TYPE_MAP = @{
    autonumber    = 4;  autoincrement = 4   # dbLong + dbAutoIncrField attribute
    long          = 4;  integer       = 3;  short = 3;  byte = 2
    text          = 10; memo          = 12; currency = 5
    double        = 7;  single        = 6;  float = 7
    datetime      = 8;  date          = 8
    boolean       = 1;  yesno         = 1;  bit = 1
    guid          = 15; ole           = 11; bigint = 16
}

# SaveAsText control type numbers → friendly names
$script:CTRL_TYPE = @{
    100 = 'Label';           101 = 'Rectangle';      102 = 'Line'
    103 = 'Image';           104 = 'CommandButton';   105 = 'OptionButton'
    106 = 'CheckBox';        107 = 'OptionGroup';     108 = 'BoundObjectFrame'
    109 = 'TextBox';         110 = 'ListBox';         111 = 'ComboBox'
    112 = 'SubForm';         113 = 'ObjectFrame';     114 = 'PageBreak'
    118 = 'Page';            119 = 'CustomControl';   122 = 'Attachment'
    124 = 'NavigationButton'; 125 = 'NavigationControl'; 126 = 'WebBrowser'
}

# Reverse map: name → number (for CreateControl)
$script:CTRL_TYPE_BY_NAME = @{}
foreach ($kv in $script:CTRL_TYPE.GetEnumerator()) {
    $script:CTRL_TYPE_BY_NAME[$kv.Value.ToLower()] = $kv.Key
}
# Additional acControlType names not in SaveAsText
$script:CTRL_TYPE_BY_NAME['customcontrol']     = 119
$script:CTRL_TYPE_BY_NAME['webbrowser']        = 128   # acWebBrowser (native, not ActiveX)
$script:CTRL_TYPE_BY_NAME['navigationcontrol'] = 129
$script:CTRL_TYPE_BY_NAME['navigationbutton']  = 130
$script:CTRL_TYPE_BY_NAME['chart']             = 133
$script:CTRL_TYPE_BY_NAME['edgebrowser']       = 134   # acEdgeBrowser

# DAO field type number → friendly name
$script:DAO_FIELD_TYPE = @{
    1 = 'Boolean';  2 = 'Byte';      3 = 'Integer';   4 = 'Long'
    5 = 'Currency'; 6 = 'Single';    7 = 'Double';    8 = 'Date/Time'
    10 = 'Text';    11 = 'OLE Object'; 12 = 'Memo';   15 = 'GUID'
    16 = 'BigInt';  20 = 'Decimal'
}

# Relationship attribute flags
$script:REL_ATTR = @{
    1    = 'Unique'
    2    = 'DontEnforce'
    256  = 'UpdateCascade'
    4096 = 'DeleteCascade'
}

# SQL prefixes that require confirmation
$script:DESTRUCTIVE_PREFIXES = @('DELETE', 'DROP', 'TRUNCATE', 'ALTER')

# Binary sections in form/report exports (stripped on export, restored on import)
$script:BINARY_SECTIONS = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@('PrtMip', 'PrtDevMode', 'PrtDevModeW', 'PrtDevNames', 'PrtDevNamesW', 'RecSrcDt', 'GUID', 'NameMap'),
    [System.StringComparer]::Ordinal
)

# VBE component name prefixes by object type
$script:VBE_PREFIX = @{
    module = ''
    form   = 'Form_'
    report = 'Report_'
}

# Form/report section name → number mapping
$script:SECTION_MAP = @{
    detail             = 0
    header             = 1; formheader       = 1; reportheader     = 1
    footer             = 2; formfooter       = 2; reportfooter     = 2
    pageheader         = 3
    pagefooter         = 4
    grouplevel1header  = 5; group1header     = 5
    grouplevel1footer  = 6; group1footer     = 6
    grouplevel2header  = 7; group2header     = 7
    grouplevel2footer  = 8; group2footer     = 8
}

# Startup property names
$script:STARTUP_PROPS = @(
    'AppTitle', 'AppIcon', 'StartupForm', 'StartupShowDBWindow',
    'StartupShowStatusBar', 'StartupShortcutMenuBar',
    'AllowShortcutMenus', 'AllowFullMenus', 'AllowBuiltInToolbars',
    'AllowToolbarChanges', 'AllowBreakIntoCode', 'AllowSpecialKeys',
    'AllowBypassKey', 'AllowDatasheetSchema'
)

# Report output format strings
$script:OUTPUT_FORMATS = @{
    pdf  = 'PDF Format (*.pdf)'
    xlsx = 'Microsoft Excel (*.xlsx)'
    rtf  = 'Rich Text Format (*.rtf)'
    txt  = 'MS-DOS Text (*.txt)'
}

# QueryDef type number → name
$script:QUERYDEF_TYPE = @{
    0   = 'Select';      16 = 'Crosstab';    32 = 'Delete'
    48  = 'Update';      64 = 'Append';      80 = 'MakeTable'
    96  = 'DDL';         112 = 'SQLPassThrough'; 128 = 'Union'
    240 = 'Action'
}

# Control properties to search in Find-AccessUsage
$script:CONTROL_SEARCH_PROPS = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@('ControlSource', 'RecordSource', 'RowSource', 'DefaultValue', 'ValidationRule'),
    [System.StringComparer]::Ordinal
)

# Access COM constants
$script:DB_AUTO_INCR_FIELD = 16       # dbAutoIncrField attribute flag
$script:DB_ATTACH_SAVE_PWD = 131072   # dbAttachSavePWD
$script:DB_SEE_CHANGES     = 512      # dbSeeChanges — ODBC IDENTITY columns

$script:AC_DESIGN    = 1   # acDesign / acViewDesign
$script:AC_FORM      = 2   # acForm (DoCmd.Close/Save)
$script:AC_REPORT    = 3   # acReport (DoCmd.Close/Save)
$script:AC_SAVE_YES  = 1   # acSaveYes
$script:AC_SAVE_NO   = 2   # acSaveNo

$script:AC_OUTPUT_REPORT    = 3    # acOutputReport
$script:AC_IMPORT           = 0    # acImport
$script:AC_EXPORT           = 1    # acExport
$script:AC_EXPORT_DELIM     = 2    # acExportDelim (CSV)
$script:AC_SPREADSHEET_XLSX = 10   # acSpreadsheetTypeExcel12Xml
$script:AC_CMD_COMPILE      = 126  # acCmdCompileAndSaveAllModules

# ═══════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════

$script:AccessSession = @{
    App            = $null      # COM Access.Application object
    DbPath         = $null      # Currently open database path (resolved)
    VbeCodeCache   = @{}        # "type:name" → full module text
    ControlsCache  = @{}        # "type:name" → parsed control structure
    CmCache        = @{}        # "type:name" → CodeModule COM object
}

# ═══════════════════════════════════════════════════════════════════════════
# INTERNAL HELPERS — Session Management
# ═══════════════════════════════════════════════════════════════════════════

function Test-AccessAlive {
    <#
    .SYNOPSIS
        Best-effort COM liveness check (does not depend on Visible).
    #>
    if ($null -eq $script:AccessSession.App) { return $false }
    try {
        $null = Get-AccessHwnd -App $script:AccessSession.App
        return $true
    } catch {}
    try {
        $null = $script:AccessSession.App.Version
        return $true
    } catch {}
    return $false
}

function Get-AccessHwnd {
    <#
    .SYNOPSIS
        Get the Access window handle. Handles hWndAccessApp being a property or method.
    #>
    param([Parameter(Mandatory)]$App)

    $h = $App.hWndAccessApp
    if ($h -is [System.Management.Automation.PSMethod]) {
        return [long]$h.Invoke(@())
    }
    return [long]$h
}

function Set-AccessVisibleBestEffort {
    <#
    .SYNOPSIS
        Try to set Access visibility. Never fail startup if unsupported.
    #>
    param([bool]$Visible = $true)
    if ($null -eq $script:AccessSession.App) { return }
    try {
        $script:AccessSession.App.Visible = $Visible
    } catch {
        Write-Verbose "Could not set Access.Visible=$Visible (continuing): $_"
    }
}

function Clear-AccessCaches {
    <#
    .SYNOPSIS
        Clear all VBE/control/CodeModule caches.
    #>
    $script:AccessSession.VbeCodeCache  = @{}
    $script:AccessSession.ControlsCache = @{}
    $script:AccessSession.CmCache       = @{}
}

function Connect-AccessDB {
    <#
    .SYNOPSIS
        Internal: ensure Access COM is running and the requested DB is open.
        Returns the COM Application object.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DbPath
    )

    $resolved = [System.IO.Path]::GetFullPath($DbPath)

    # If we have an existing session, check liveness
    if ($null -ne $script:AccessSession.App) {
        if (-not (Test-AccessAlive)) {
            Write-Verbose 'COM session stale — auto-reconnecting...'
            # Force cleanup without calling methods on dead COM
            $script:AccessSession.App    = $null
            $script:AccessSession.DbPath = $null
            Clear-AccessCaches
        }
    }

    # Launch Access if needed
    if ($null -eq $script:AccessSession.App) {
        Write-Verbose 'Launching Access.Application...'
        try {
            $script:AccessSession.App = New-Object -ComObject 'Access.Application'
        } catch {
            throw "Failed to create Access.Application COM object. Is Microsoft Access installed? Error: $_"
        }
        Set-AccessVisibleBestEffort -Visible $true
        Write-Verbose 'Access launched OK'
    }

    # Switch database if needed
    if ($script:AccessSession.DbPath -ne $resolved) {
        if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
            throw "Database file not found: $resolved"
        }

        # Close previous database
        if ($null -ne $script:AccessSession.DbPath) {
            Write-Verbose "Closing previous DB: $($script:AccessSession.DbPath)"
            try {
                $script:AccessSession.App.CloseCurrentDatabase()
            } catch {
                Write-Verbose "Error closing previous DB: $_"
            }
        }

        # Open new database
        Write-Verbose "Opening DB: $resolved"
        try {
            $script:AccessSession.App.OpenCurrentDatabase($resolved)
        } catch {
            if ($_.Exception.Message -match 'already have the database open') {
                Write-Verbose 'DB was already open — syncing state'
            } else {
                throw
            }
        }

        $script:AccessSession.DbPath = $resolved
        Set-AccessVisibleBestEffort -Visible $true
        Clear-AccessCaches
        Write-Verbose 'DB opened OK'
    }

    return $script:AccessSession.App
}

# ═══════════════════════════════════════════════════════════════════════════
# PUBLIC — Close-AccessDatabase
# ═══════════════════════════════════════════════════════════════════════════

function Close-AccessDatabase {
    <#
    .SYNOPSIS
        Close the Access COM session and release the file lock.
    .DESCRIPTION
        Closes the current database, quits Access, releases COM objects,
        and clears all caches. Safe to call even if no session is open.
    .EXAMPLE
        Close-AccessDatabase
    #>
    [CmdletBinding()]
    param()

    if ($null -ne $script:AccessSession.App) {
        Write-Verbose 'Closing Access...'
        # Capture PID before we lose the COM object
        $accessPid = $null
        try {
            $hwnd = Get-AccessHwnd -App $script:AccessSession.App
            [uint32]$pid = 0
            $null = [AccessPoshNative]::GetWindowThreadProcessId([IntPtr]::new($hwnd), [ref]$pid)
            if ($pid -gt 0) { $accessPid = [int]$pid }
        } catch {}

        try {
            if ($null -ne $script:AccessSession.DbPath) {
                $script:AccessSession.App.CloseCurrentDatabase()
            }
        } catch {
            Write-Verbose "Error closing DB: $_"
        }
        try {
            $script:AccessSession.App.Quit()
            Write-Verbose 'Access quit OK'
        } catch {
            Write-Verbose "Error quitting Access: $_"
        }
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:AccessSession.App)
        } catch {}
        $script:AccessSession.App    = $null
        $script:AccessSession.DbPath = $null
        Clear-AccessCaches

        # Wait briefly then kill if still alive
        if ($null -ne $accessPid) {
            Start-Sleep -Milliseconds 500
            $proc = Get-Process -Id $accessPid -ErrorAction SilentlyContinue
            if ($null -ne $proc -and -not $proc.HasExited) {
                Write-Verbose "Access process $accessPid still alive after Quit — terminating"
                Stop-Process -Id $accessPid -Force -ErrorAction SilentlyContinue
            }
        }
        Write-Verbose 'Access closed OK'
    }
}

# Register cleanup on PowerShell exit
Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    if ($null -ne $script:AccessSession -and $null -ne $script:AccessSession.App) {
        try { $script:AccessSession.App.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:AccessSession.App) } catch {}
        $script:AccessSession.App    = $null
        $script:AccessSession.DbPath = $null
    }
} | Out-Null

# ═══════════════════════════════════════════════════════════════════════════
# INTERNAL HELPERS — Value Conversion
# ═══════════════════════════════════════════════════════════════════════════

function ConvertTo-SafeValue {
    <#
    .SYNOPSIS
        Convert COM values to PowerShell-safe types for JSON serialization.
    #>
    param([AllowNull()]$Value)

    if ($null -eq $Value)               { return $null }
    if ($Value -is [System.DBNull])     { return $null }
    if ($Value -is [System.DateTime])   { return $Value.ToString('o') }  # ISO 8601
    if ($Value -is [decimal])           { return [double]$Value }
    if ($Value -is [byte[]])            { return "<binary $($Value.Length) bytes>" }
    return $Value
}

function ConvertTo-CoercedProp {
    <#
    .SYNOPSIS
        Convert string property values to int/bool as needed for COM.
    #>
    param($Value)

    if ($Value -is [int] -or $Value -is [double] -or $Value -is [float] -or $Value -is [bool]) {
        return $Value
    }
    if ($Value -is [string]) {
        $low = $Value.ToLower()
        if ($low -in 'true', 'yes', '-1')  { return $true }
        if ($low -in 'false', 'no', '0')   { return $false }
        $intVal = 0
        if ([int]::TryParse($Value, [ref]$intVal))    { return $intVal }
        $dblVal = 0.0
        if ([double]::TryParse($Value, [ref]$dblVal))  { return $dblVal }
    }
    return $Value
}

# ═══════════════════════════════════════════════════════════════════════════
# INTERNAL HELPERS — Output Formatting
# ═══════════════════════════════════════════════════════════════════════════

function Format-AccessOutput {
    <#
    .SYNOPSIS
        Handle -AsJson switch: convert hashtable/PSCustomObject to JSON or return as-is.
    #>
    param(
        [Parameter(Mandatory)]$Data,
        [switch]$AsJson
    )

    if ($Data -is [hashtable]) {
        $Data = [PSCustomObject]$Data
    }
    if ($AsJson) {
        return $Data | ConvertTo-Json -Depth 10 -Compress
    }
    return $Data
}

# ═══════════════════════════════════════════════════════════════════════════
# INTERNAL HELPERS — Temp File I/O (encoding negotiation)
# ═══════════════════════════════════════════════════════════════════════════

function Read-TempFile {
    <#
    .SYNOPSIS
        Read a file exported by Access. Auto-detects encoding (UTF-16 BOM, UTF-8-sig, cp1252).
        Returns [PSCustomObject]@{ Content = [string]; Encoding = [string] }
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    # Check BOM
    $bom = [byte[]]::new(2)
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $null = $fs.Read($bom, 0, 2)
    } finally {
        $fs.Close()
    }

    # UTF-16 LE or BE BOM
    if (($bom[0] -eq 0xFF -and $bom[1] -eq 0xFE) -or ($bom[0] -eq 0xFE -and $bom[1] -eq 0xFF)) {
        $content = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Unicode)
        return [PSCustomObject]@{ Content = $content; Encoding = 'utf-16' }
    }

    # Try UTF-8 with BOM
    try {
        $utf8Bom = New-Object System.Text.UTF8Encoding($true, $true)  # throwOnInvalid
        $content = [System.IO.File]::ReadAllText($Path, $utf8Bom)
        return [PSCustomObject]@{ Content = $content; Encoding = 'utf-8-sig' }
    } catch {}

    # Try Windows-1252 (cp1252) — Access default for ANSI modules
    try {
        $cp1252 = [System.Text.Encoding]::GetEncoding(1252)
        $content = [System.IO.File]::ReadAllText($Path, $cp1252)
        return [PSCustomObject]@{ Content = $content; Encoding = 'cp1252' }
    } catch {}

    # Fallback: UTF-8 with replacement
    $content = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    return [PSCustomObject]@{ Content = $content; Encoding = 'utf-8' }
}

function Write-TempFile {
    <#
    .SYNOPSIS
        Write content for Access to read with LoadFromText.
        Default utf-16 (Access .accdb expects UTF-16LE with BOM).
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$Content,
        [string]$Encoding = 'utf-16'
    )

    switch ($Encoding) {
        'utf-16' {
            [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::Unicode)
        }
        'cp1252' {
            $cp1252 = [System.Text.Encoding]::GetEncoding(1252)
            [System.IO.File]::WriteAllText($Path, $Content, $cp1252)
        }
        default {
            [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::UTF8)
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# INTERNAL HELPERS — Binary Section Handling (Forms/Reports)
# ═══════════════════════════════════════════════════════════════════════════

function Remove-BinarySections {
    <#
    .SYNOPSIS
        Strip binary sections (PrtMip, PrtDevMode, NameMap, etc.) from a form/report export.
        Reduces size ~20x without affecting VBA or controls.
        Also removes the Checksum line (Access recalculates on import).
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Text
    )

    $lines = $Text.Split([string[]]@("`r`n", "`n"), [System.StringSplitOptions]::None)
    $result = [System.Collections.Generic.List[string]]::new($lines.Count)
    $skipDepth  = 0
    $skipIndent = ''

    foreach ($line in $lines) {
        $rstripped = $line.TrimEnd("`r", "`n")
        $stripped  = $rstripped.TrimStart()
        $indent    = $rstripped.Substring(0, $rstripped.Length - $stripped.Length)

        if ($skipDepth -gt 0) {
            # Is this the closing End at the same indent level?
            if ($stripped -eq 'End' -and $indent -eq $skipIndent) {
                $skipDepth--
            }
            continue  # skip this line (part of binary block)
        }

        # Checksum line at root level
        if ($rstripped -match '^\s*Checksum\s*=\s*') {
            continue
        }

        # Does a binary block begin here?
        if ($rstripped -match '^(\s*)(\w+)\s*=\s*Begin\s*$') {
            $blockIndent = $Matches[1]
            $blockName   = $Matches[2]
            if ($script:BINARY_SECTIONS.Contains($blockName)) {
                $skipIndent = $blockIndent
                $skipDepth  = 1
                continue
            }
        }

        $result.Add($line)
    }

    return ($result -join "`r`n")
}

function Get-BinaryBlocks {
    <#
    .SYNOPSIS
        Extract binary Begin...End blocks from the original form/report export.
        Returns a hashtable: { section_name = full_block_text }.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Text
    )

    $blocks = @{}
    $lines = $Text.Split([string[]]@("`r`n", "`n"), [System.StringSplitOptions]::None)
    $i = 0

    while ($i -lt $lines.Count) {
        $line = $lines[$i]
        $rstripped = $line.TrimEnd("`r", "`n")

        if ($rstripped -match '^(\s*)(\w+)\s*=\s*Begin\s*$') {
            $blockIndent = $Matches[1]
            $blockName   = $Matches[2]
            if ($script:BINARY_SECTIONS.Contains($blockName)) {
                $blockLines = [System.Collections.Generic.List[string]]::new()
                $blockLines.Add($line)
                $j = $i + 1
                while ($j -lt $lines.Count) {
                    $bl = $lines[$j]
                    $blr = $bl.TrimEnd("`r", "`n")
                    $bls = $blr.TrimStart()
                    $blIndent = $blr.Substring(0, $blr.Length - $bls.Length)
                    $blockLines.Add($bl)
                    if ($bls -eq 'End' -and $blIndent -eq $blockIndent) {
                        break
                    }
                    $j++
                }
                $blocks[$blockName] = ($blockLines -join "`r`n")
                $i = $j + 1
                continue
            }
        }
        $i++
    }

    return $blocks
}

function Restore-BinarySections {
    <#
    .SYNOPSIS
        Re-inject binary sections from the current Access object's export,
        before calling LoadFromText with edited code.
        If the object doesn't exist yet, returns the code unmodified.
    #>
    param(
        [Parameter(Mandatory)]$App,
        [Parameter(Mandatory)][string]$ObjectType,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$NewCode
    )

    $tmp = [System.IO.Path]::GetTempFileName()
    try {
        try {
            $App.SaveAsText($script:AC_TYPE[$ObjectType], $Name, $tmp)
        } catch {
            Write-Verbose "Restore-BinarySections: '$Name' doesn't exist yet — importing without binary sections"
            return $NewCode
        }
        $original = (Read-TempFile -Path $tmp).Content
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }

    $blocks = Get-BinaryBlocks -Text $original
    if ($blocks.Count -eq 0) { return $NewCode }

    # Inject blocks just before "End Form" / "End Report"
    $lines = $NewCode.Split([string[]]@("`r`n", "`n"), [System.StringSplitOptions]::None)
    $result = [System.Collections.Generic.List[string]]::new($lines.Count + 50)
    $inTopForm = $false
    $injected  = $false

    foreach ($line in $lines) {
        $stripped = $line.Trim()

        if ($stripped -match '^\s*Begin\s+(Form|Report)\s*$') {
            $inTopForm = $true
        }

        if ($inTopForm -and (-not $injected) -and $stripped -match '^\s*End\s+(Form|Report)\s*$') {
            foreach ($blockText in $blocks.Values) {
                $result.Add($blockText)
            }
            $injected  = $true
            $inTopForm = $false
        }

        $result.Add($line)
    }

    return ($result -join "`r`n")
}

# ═══════════════════════════════════════════════════════════════════════════
# INTERNAL HELPERS — CodeBehind Splitting (for Set-AccessCode)
# ═══════════════════════════════════════════════════════════════════════════

function Split-CodeBehind {
    <#
    .SYNOPSIS
        Separate form/report export text into (form_text, vba_code).
        If CodeBehindForm/CodeBehindReport marker exists, splits there.
        Returns [PSCustomObject]@{ FormText; VbaCode }
    #>
    param(
        [Parameter(Mandatory)][string]$Code
    )

    foreach ($marker in @('CodeBehindForm', 'CodeBehindReport')) {
        $idx = $Code.IndexOf($marker)
        if ($idx -ge 0) {
            $formPart = $Code.Substring(0, $idx).TrimEnd() + "`n"
            $remainder = $Code.Substring($idx)
            $parts = $remainder -split "`n", 2
            $vbaCode = if ($parts.Count -gt 1) { $parts[1] } else { '' }
            # Strip Attribute VB_ lines (auto-generated)
            $vbaLines = foreach ($line in $vbaCode.Split("`n")) {
                if ($line.Trim() -notmatch '^Attribute VB_') { $line }
            }
            $vbaCode = ($vbaLines -join "`n").Trim()
            return [PSCustomObject]@{ FormText = $formPart; VbaCode = $vbaCode }
        }
    }
    return [PSCustomObject]@{ FormText = $Code; VbaCode = '' }
}

function Set-FieldProperty {
    <#
    .SYNOPSIS
        Set a field-level DAO property, creating it if it doesn't exist.
    #>
    param(
        [Parameter(Mandatory)]$Db,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][string]$FieldName,
        [Parameter(Mandatory)][string]$PropertyName,
        $Value
    )

    $fld = $Db.TableDefs($TableName).Fields($FieldName)
    try {
        $fld.Properties($PropertyName).Value = $Value
    } catch {
        $prop = $fld.CreateProperty($PropertyName, 10, $Value)  # 10 = dbText
        $fld.Properties.Append($prop)
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# PHASE 2 — Core Database Operations
# ═══════════════════════════════════════════════════════════════════════════

# ── 2.1 New-AccessDatabase ────────────────────────────────────────────────

function New-AccessDatabase {
    <#
    .SYNOPSIS
        Create a new empty Access database (.accdb).
    .PARAMETER DbPath
        Full path for the new database file. Fails if already exists.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-AccessDatabase -DbPath "C:\Data\new.accdb"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [switch]$AsJson
    )

    $resolved = [System.IO.Path]::GetFullPath($DbPath)
    if (Test-Path -LiteralPath $resolved) {
        throw "File already exists: $resolved. Use Invoke-AccessSQL to modify it."
    }

    # Ensure Access is running
    if ($null -eq $script:AccessSession.App) {
        try {
            $script:AccessSession.App = New-Object -ComObject 'Access.Application'
        } catch {
            throw "Failed to create Access.Application COM object. Is Microsoft Access installed? Error: $_"
        }
        Set-AccessVisibleBestEffort -Visible $true
    }
    $app = $script:AccessSession.App

    # Close any previously open DB
    if ($null -ne $script:AccessSession.DbPath) {
        try { $app.CloseCurrentDatabase() } catch {}
        $script:AccessSession.DbPath = $null
    }

    try {
        $app.NewCurrentDatabase($resolved)
    } catch {
        throw "Error creating database: $_"
    }

    # Close and reopen to ensure CurrentDb() works reliably
    try {
        $app.CloseCurrentDatabase()
        $app.OpenCurrentDatabase($resolved)
    } catch {}

    $script:AccessSession.DbPath = $resolved
    Clear-AccessCaches

    $size = 0
    if (Test-Path -LiteralPath $resolved) {
        $size = (Get-Item -LiteralPath $resolved).Length
    }

    Format-AccessOutput -AsJson:$AsJson -Data @{
        db_path    = $resolved
        status     = 'created'
        size_bytes = $size
    }
}

# ── 2.2 Repair-AccessDatabase ────────────────────────────────────────────

function Repair-AccessDatabase {
    <#
    .SYNOPSIS
        Compact and repair the database. Closes DB, compacts to temp, atomic swap, reopens.
    .PARAMETER DbPath
        Path to the database to compact/repair.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Repair-AccessDatabase -DbPath "C:\Data\mydb.accdb"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [switch]$AsJson
    )

    $resolved = [System.IO.Path]::GetFullPath($DbPath)
    $app = Connect-AccessDB -DbPath $resolved
    $originalSize = (Get-Item -LiteralPath $resolved).Length

    # Close current database (keep Access alive)
    # Release any outstanding DAO references first
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    try {
        $app.CloseCurrentDatabase()
    } catch {
        throw "Could not close database for compact/repair: $_"
    }
    $script:AccessSession.DbPath = $null
    Clear-AccessCaches

    # Brief delay to ensure file lock is released
    Start-Sleep -Milliseconds 500

    # Temp/bak paths in same directory
    $dir  = [System.IO.Path]::GetDirectoryName($resolved)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($resolved)
    $ext  = [System.IO.Path]::GetExtension($resolved)
    $tmpPath = Join-Path $dir "${base}_compact_tmp${ext}"
    $bakPath = Join-Path $dir "${base}_compact_bak${ext}"

    try {
        foreach ($p in @($tmpPath, $bakPath)) {
            if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force }
        }

        try {
            $app.CompactRepair($resolved, $tmpPath)
        } catch {
            if (Test-Path -LiteralPath $tmpPath) { Remove-Item -LiteralPath $tmpPath -Force }
            throw "CompactRepair failed: $_"
        }

        if (-not (Test-Path -LiteralPath $tmpPath)) {
            throw 'CompactRepair did not produce output file'
        }
        $compactedSize = (Get-Item -LiteralPath $tmpPath).Length

        # Atomic swap: original → .bak, tmp → original
        Rename-Item -LiteralPath $resolved -NewName ([System.IO.Path]::GetFileName($bakPath)) -Force
        try {
            Rename-Item -LiteralPath $tmpPath -NewName ([System.IO.Path]::GetFileName($resolved)) -Force
        } catch {
            # Rollback
            Rename-Item -LiteralPath $bakPath -NewName ([System.IO.Path]::GetFileName($resolved)) -Force
            throw
        }

        Remove-Item -LiteralPath $bakPath -Force -ErrorAction SilentlyContinue

    } catch {
        # Try to reopen whatever is at the original path
        try {
            if (Test-Path -LiteralPath $resolved) {
                $app.OpenCurrentDatabase($resolved)
                $script:AccessSession.DbPath = $resolved
            }
        } catch {}
        throw
    }

    # Reopen compacted database
    try {
        $app.OpenCurrentDatabase($resolved)
        $script:AccessSession.DbPath = $resolved
    } catch {
        throw "Database compacted OK but failed to reopen: $_"
    }

    $saved = $originalSize - $compactedSize
    $pct   = if ($originalSize -gt 0) { [math]::Round($saved / $originalSize * 100, 1) } else { 0 }

    Format-AccessOutput -AsJson:$AsJson -Data @{
        original_size  = $originalSize
        compacted_size = $compactedSize
        saved_bytes    = $saved
        saved_pct      = $pct
        status         = 'compacted'
    }
}

# ── 2.2b Invoke-AccessDecompile ──────────────────────────────────────────

function Invoke-AccessDecompile {
    <#
    .SYNOPSIS
        Decompile VBA p-code, recompile, and compact the database.
        Strips orphaned p-code via MSACCESS.EXE /decompile, then compacts.
        Typical size reduction: 60-70%.
    .PARAMETER DbPath
        Path to the Access database (.accdb/.mdb).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-AccessDecompile -DbPath "C:\Data\mydb.accdb"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [switch]$AsJson
    )

    $resolved = [System.IO.Path]::GetFullPath($DbPath)
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw "Database not found: $resolved"
    }
    $originalSize = (Get-Item -LiteralPath $resolved).Length

    # ── 1. Close COM session completely so the file is unlocked ──
    if ($null -ne $script:AccessSession.App) {
        Write-Verbose 'Closing COM session for /decompile...'
        try {
            if ($null -ne $script:AccessSession.DbPath) {
                $script:AccessSession.App.CloseCurrentDatabase()
            }
        } catch {
            Write-Verbose "Error closing DB: $_"
        }
        $script:AccessSession.DbPath = $null
        Clear-AccessCaches
        try {
            $script:AccessSession.App.Quit(1)  # acQuitSaveNone
        } catch {
            Write-Verbose "Error quitting Access: $_"
        }
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:AccessSession.App)
        } catch {}
        $script:AccessSession.App = $null
    }

    # ── 2. Find MSACCESS.EXE ──
    $msaccessCandidates = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\MSACCESS.EXE"
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\MSACCESS.EXE"
    )
    $msaccess = $msaccessCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $msaccess) {
        throw 'MSACCESS.EXE not found in known Office 16 paths'
    }

    # ── 3. Launch MSACCESS /decompile with SHIFT held ──
    $VK_SHIFT        = 0x10
    $KEYEVENTF_KEYUP = 0x0002
    $shiftHeld       = $false
    try {
        [AccessPoshNative]::keybd_event($VK_SHIFT, 0, 0, [UIntPtr]::Zero)
        Start-Sleep -Milliseconds 300
        $shiftHeld = $true
        Write-Verbose 'SHIFT held for /decompile bypass'
    } catch {
        Write-Verbose "Could not simulate SHIFT — AutoExec may run: $_"
    }

    $proc = $null
    try {
        $proc = Start-Process -FilePath $msaccess `
                              -ArgumentList "`"$resolved`"", '/decompile' `
                              -PassThru
    } catch {
        if ($shiftHeld) {
            try { [AccessPoshNative]::keybd_event($VK_SHIFT, 0, $KEYEVENTF_KEYUP, [UIntPtr]::Zero) } catch {}
        }
        throw "Failed to launch MSACCESS /decompile: $_"
    }

    # Wait for Access to read key state, then release SHIFT
    Start-Sleep -Seconds 3
    if ($shiftHeld) {
        try { [AccessPoshNative]::keybd_event($VK_SHIFT, 0, $KEYEVENTF_KEYUP, [UIntPtr]::Zero) } catch {}
        Write-Verbose 'SHIFT released'
    }

    # Wait for decompile to complete, then kill Access
    Start-Sleep -Seconds 5
    if ($null -ne $proc -and -not $proc.HasExited) {
        Write-Verbose "Killing decompile Access process (PID $($proc.Id))"
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
    }

    $decompileSize = (Get-Item -LiteralPath $resolved).Length

    # ── 4. Reopen via COM and try to recompile VBA ──
    Write-Verbose 'Relaunching COM after decompile...'
    try {
        $script:AccessSession.App = New-Object -ComObject 'Access.Application'
    } catch {
        throw "Failed to relaunch Access COM after decompile: $_"
    }
    Set-AccessVisibleBestEffort -Visible $true

    try {
        $script:AccessSession.App.OpenCurrentDatabase($resolved)
    } catch {
        if ($_.Exception.Message -match 'already have the database open') {
            Write-Verbose 'DB was already open — syncing state'
        } else {
            throw "Failed to reopen DB after decompile: $_"
        }
    }
    $script:AccessSession.DbPath = $resolved

    # Best-effort recompile (not critical for compact)
    try {
        $script:AccessSession.App.RunCommand(137)  # acCmdCompileAllModules
        Write-Verbose 'VBA recompiled after decompile'
    } catch {
        Write-Verbose "VBA recompile skipped: $_"
    }

    # Close DB for compact
    try {
        $script:AccessSession.App.CloseCurrentDatabase()
    } catch {
        Write-Verbose "Error closing DB before compact: $_"
    }
    $script:AccessSession.DbPath = $null
    Clear-AccessCaches

    # ── 5. Compact & Repair ──
    $dir     = [System.IO.Path]::GetDirectoryName($resolved)
    $base    = [System.IO.Path]::GetFileNameWithoutExtension($resolved)
    $ext     = [System.IO.Path]::GetExtension($resolved)
    $tmpPath = Join-Path $dir "${base}_compact_tmp${ext}"
    $bakPath = Join-Path $dir "${base}_compact_bak${ext}"

    foreach ($p in @($tmpPath, $bakPath)) {
        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force }
    }

    try {
        $script:AccessSession.App.CompactRepair($resolved, $tmpPath)
    } catch {
        if (Test-Path -LiteralPath $tmpPath) { Remove-Item -LiteralPath $tmpPath -Force }
        throw "CompactRepair after decompile failed: $_"
    }

    if (-not (Test-Path -LiteralPath $tmpPath)) {
        throw 'CompactRepair did not produce output file'
    }
    $compactedSize = (Get-Item -LiteralPath $tmpPath).Length

    # Atomic swap
    Rename-Item -LiteralPath $resolved -NewName ([System.IO.Path]::GetFileName($bakPath)) -Force
    try {
        Rename-Item -LiteralPath $tmpPath -NewName ([System.IO.Path]::GetFileName($resolved)) -Force
    } catch {
        # Rollback
        Rename-Item -LiteralPath $bakPath -NewName ([System.IO.Path]::GetFileName($resolved)) -Force
        throw
    }
    Remove-Item -LiteralPath $bakPath -Force -ErrorAction SilentlyContinue

    # ── 6. Reopen compacted database ──
    try {
        $script:AccessSession.App.OpenCurrentDatabase($resolved)
        $script:AccessSession.DbPath = $resolved
    } catch {
        if ($_.Exception.Message -match 'already have the database open') {
            $script:AccessSession.DbPath = $resolved
        } else {
            throw "Decompile+compact OK but failed to reopen: $_"
        }
    }

    $saved = $originalSize - $compactedSize
    $pct   = if ($originalSize -gt 0) { [math]::Round($saved / $originalSize * 100, 1) } else { 0 }

    Format-AccessOutput -AsJson:$AsJson -Data @{
        original_size   = $originalSize
        decompile_size  = $decompileSize
        compacted_size  = $compactedSize
        saved_bytes     = $saved
        saved_pct       = $pct
        status          = 'decompiled_and_compacted'
    }
}

# ── 2.3 Get-AccessObject ─────────────────────────────────────────────────

function Get-AccessObject {
    <#
    .SYNOPSIS
        List objects in the database by type.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type to list: table, query, form, report, macro, module, or all (default).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessObject -DbPath "C:\Data\mydb.accdb"
        Get-AccessObject -DbPath "C:\Data\mydb.accdb" -ObjectType form
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [ValidateSet('all','table','query','form','report','macro','module')]
        [string]$ObjectType = 'all',
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $containers = [ordered]@{
        table  = $app.CurrentData.AllTables
        query  = $app.CurrentData.AllQueries
        form   = $app.CurrentProject.AllForms
        report = $app.CurrentProject.AllReports
        macro  = $app.CurrentProject.AllMacros
        module = $app.CurrentProject.AllModules
    }

    $keys = if ($ObjectType -eq 'all') { $containers.Keys } else { @($ObjectType) }
    $result = [ordered]@{}

    foreach ($k in $keys) {
        if (-not $containers.Contains($k)) { continue }
        $col = $containers[$k]
        $names = for ($i = 0; $i -lt $col.Count; $i++) { $col.Item($i).Name }
        if ($k -eq 'table') {
            $names = @($names | Where-Object { $_ -notlike 'MSys*' -and $_ -notlike '~*' })
        }
        $result[$k] = @($names)
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 2.4 Get-AccessCode ───────────────────────────────────────────────────

function Get-AccessCode {
    <#
    .SYNOPSIS
        Export an Access object to text via SaveAsText.
        For forms/reports, binary sections are stripped (restored automatically by Set-AccessCode).
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type of object: query, form, report, macro, module.
    .PARAMETER Name
        Name of the object.
    .EXAMPLE
        Get-AccessCode -DbPath "C:\Data\mydb.accdb" -ObjectType module -Name "Module1"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('query','form','report','macro','module')]
        [string]$ObjectType,
        [Parameter(Mandatory)][string]$Name
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $tmp = [System.IO.Path]::GetTempFileName()
    try {
        $app.SaveAsText($script:AC_TYPE[$ObjectType], $Name, $tmp)
        $fileResult = Read-TempFile -Path $tmp
        $text = $fileResult.Content
        if ($ObjectType -in 'form', 'report') {
            $text = Remove-BinarySections -Text $text
        }
        return $text
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}

# ── 2.5 Set-AccessCode ───────────────────────────────────────────────────

function Set-AccessCode {
    <#
    .SYNOPSIS
        Import text as an Access object definition (creates or overwrites).
        For forms/reports, binary sections are auto-restored from current export.
        If code contains CodeBehindForm/CodeBehindReport, VBA is separated and
        injected via VBE after import.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type of object: query, form, report, macro, module.
    .PARAMETER Name
        Name of the object.
    .PARAMETER Code
        The text definition to import.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        $code = Get-AccessCode -DbPath "C:\db.accdb" -ObjectType module -Name "Module1"
        Set-AccessCode -DbPath "C:\db.accdb" -ObjectType module -Name "Module1" -Code $code
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('query','form','report','macro','module')]
        [string]$ObjectType,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Code,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath

    # Separate CodeBehindForm/CodeBehindReport if present
    $vbaCode = ''
    if ($ObjectType -in 'form', 'report') {
        $split = Split-CodeBehind -Code $Code
        $Code    = $split.FormText
        $vbaCode = $split.VbaCode

        # Remove HasModule from form text — will be set when injecting VBA
        if ($vbaCode) {
            $Code = $Code -replace '(?m)^\s*HasModule\s*=.*$', ''
        }
    }

    # Restore binary sections if they were stripped
    if ($ObjectType -in 'form', 'report') {
        $hasBinary = $false
        foreach ($s in $script:BINARY_SECTIONS) {
            if ($Code.Contains($s)) { $hasBinary = $true; break }
        }
        if (-not $hasBinary) {
            Write-Verbose "Restoring binary sections for '$Name'"
            $Code = Restore-BinarySections -App $app -ObjectType $ObjectType -Name $Name -NewCode $Code
        }
    }

    # Backup existing object
    $backupTmp = $null
    if ($ObjectType -in 'form', 'report', 'module') {
        try {
            $backupTmp = [System.IO.Path]::GetTempFileName()
            $app.SaveAsText($script:AC_TYPE[$ObjectType], $Name, $backupTmp)
        } catch {
            if ($backupTmp) { Remove-Item -LiteralPath $backupTmp -Force -ErrorAction SilentlyContinue }
            $backupTmp = $null
        }
    }

    $tmp = [System.IO.Path]::GetTempFileName()
    try {
        # Modules expect cp1252; forms/reports/queries/macros expect UTF-16
        $enc = if ($ObjectType -eq 'module') { 'cp1252' } else { 'utf-16' }
        Write-TempFile -Path $tmp -Content $Code -Encoding $enc

        try {
            $app.LoadFromText($script:AC_TYPE[$ObjectType], $Name, $tmp)
        } catch {
            # Restore backup if exists
            if ($backupTmp -and (Test-Path -LiteralPath $backupTmp)) {
                Write-Warning "Import failed, restoring backup of '$Name'"
                try {
                    $app.LoadFromText($script:AC_TYPE[$ObjectType], $Name, $backupTmp)
                } catch {
                    Write-Warning "Could not restore backup of '$Name'"
                }
            }
            throw
        }

        # Invalidate caches
        $cacheKey = "${ObjectType}:${Name}"
        $script:AccessSession.VbeCodeCache.Remove($cacheKey)
        $script:AccessSession.CmCache.Remove($cacheKey)
        $script:AccessSession.ControlsCache.Remove($cacheKey)

        # Inject VBA if CodeBehindForm was present
        $vbaMsg = ''
        if ($vbaCode) {
            Invoke-VbaAfterImport -App $app -ObjectType $ObjectType -Name $Name -VbaCode $vbaCode
            $vbaMsg = ' (with VBA injected via VBE)'
        }

        $msg = "OK: '$Name' ($ObjectType) imported successfully${vbaMsg}"
        Format-AccessOutput -AsJson:$AsJson -Data @{ status = 'imported'; message = $msg; object_type = $ObjectType; name = $Name }
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        if ($backupTmp) { Remove-Item -LiteralPath $backupTmp -Force -ErrorAction SilentlyContinue }
    }
}

function Invoke-VbaAfterImport {
    <#
    .SYNOPSIS
        Internal: Inject VBA code into a form/report after LoadFromText import.
        Opens in design, enables HasModule, then injects via VBE CodeModule.
    #>
    param(
        [Parameter(Mandatory)]$App,
        [Parameter(Mandatory)][string]$ObjectType,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$VbaCode
    )

    if (-not $VbaCode.Trim()) { return }

    # 1. Open in design and enable HasModule
    $acObjType = if ($ObjectType -eq 'form') { $script:AC_FORM } else { $script:AC_REPORT }
    if ($ObjectType -eq 'form') {
        $App.DoCmd.OpenForm($Name, $script:AC_DESIGN)
    } else {
        $App.DoCmd.OpenReport($Name, $script:AC_DESIGN)
    }
    try {
        $obj = if ($ObjectType -eq 'form') { $App.Forms($Name) } else { $App.Reports($Name) }
        $obj.HasModule = $true
    } finally {
        $App.DoCmd.Close($acObjType, $Name, $script:AC_SAVE_YES)
    }

    # 2. Clear VBE cache
    $cacheKey = "${ObjectType}:${Name}"
    $script:AccessSession.CmCache.Remove($cacheKey)
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)

    # 3. Get CodeModule and inject VBA
    $prefix = $script:VBE_PREFIX[$ObjectType]
    $compName = "${prefix}${Name}"
    $cm = $App.VBE.ActiveVBProject.VBComponents($compName).CodeModule

    $total = $cm.CountOfLines
    if ($total -gt 0) {
        $cm.DeleteLines(1, $total)
    }

    # Normalize line endings to CRLF
    $VbaCode = $VbaCode -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
    if (-not $VbaCode.EndsWith("`r`n")) { $VbaCode += "`r`n" }

    $cm.InsertLines(1, $VbaCode)

    # Invalidate caches
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)
    $script:AccessSession.CmCache.Remove($cacheKey)
}

# ── 2.6 Remove-AccessObject ──────────────────────────────────────────────

function Remove-AccessObject {
    <#
    .SYNOPSIS
        Delete an Access object (module, form, report, query, macro).
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type of object: query, form, report, macro, module.
    .PARAMETER Name
        Name of the object to delete.
    .PARAMETER Confirm
        Must be specified to confirm destructive operation.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-AccessObject -DbPath "C:\db.accdb" -ObjectType module -Name "Module1" -Confirm
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('query','form','report','macro','module')]
        [string]$ObjectType,
        [Parameter(Mandatory)][string]$Name,
        [switch]$Confirm,
        [switch]$AsJson
    )

    if (-not $Confirm) {
        throw "Destructive operation: -Confirm is required to delete an object."
    }

    $app = Connect-AccessDB -DbPath $DbPath
    try {
        $app.DoCmd.DeleteObject($script:AC_TYPE[$ObjectType], $Name)
    } catch {
        throw "Error deleting $ObjectType '$Name': $_"
    } finally {
        Clear-AccessCaches
    }

    Format-AccessOutput -AsJson:$AsJson -Data @{
        action      = 'deleted'
        object_type = $ObjectType
        object_name = $Name
    }
}

# ── 2.7 Export-AccessStructure ────────────────────────────────────────────

function Export-AccessStructure {
    <#
    .SYNOPSIS
        Generate a Markdown file with the complete database structure.
        Includes VBA module signatures, forms, reports, queries, and macros.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER OutputPath
        Path for the output Markdown file. Defaults to db_structure.md next to the database.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Export-AccessStructure -DbPath "C:\Data\mydb.accdb"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [string]$OutputPath,
        [switch]$AsJson
    )

    if (-not $OutputPath) {
        $dir = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($DbPath))
        $OutputPath = Join-Path $dir 'db_structure.md'
    }

    $objects = Get-AccessObject -DbPath $DbPath -ObjectType all
    $modules = @($objects.module)
    $forms   = @($objects.form)
    $reports = @($objects.report)
    $queries = @($objects.query)
    $macros  = @($objects.macro)

    $dbName = [System.IO.Path]::GetFileName($DbPath)
    $now    = (Get-Date).ToString('yyyy-MM-dd HH:mm')

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("# Structure of ``$dbName``")
    $lines.Add("")
    $lines.Add("**Path**: ``$DbPath``  ")
    $lines.Add("**Generated**: $now  ")
    $lines.Add("**Summary**: $($modules.Count) modules, $($forms.Count) forms, $($reports.Count) reports, $($queries.Count) queries, $($macros.Count) macros")
    $lines.Add("")

    # VBA Modules with signatures
    $app = $script:AccessSession.App
    $lines.Add("## VBA Modules ($($modules.Count))")
    $lines.Add("")
    foreach ($modName in $modules) {
        $lines.Add("### ``$modName``")
        try {
            $prefix = $script:VBE_PREFIX['module']
            $compName = "${prefix}${modName}"
            $cm = $app.VBE.ActiveVBProject.VBComponents($compName).CodeModule
            $total = $cm.CountOfLines
            $code = if ($total -gt 0) { $cm.Lines(1, $total) } else { '' }
            $sigs = foreach ($codeLine in $code.Split("`n")) {
                $s = $codeLine.Trim()
                if ($s -match '^(Public\s+|Private\s+|Friend\s+)?(Function|Sub)\s+\w+') {
                    "  - ``$s``"
                }
            }
            if ($sigs) {
                $lines.AddRange([string[]]@($sigs))
            } else {
                $lines.Add('  *(no public functions/subs)*')
            }
        } catch {
            $lines.Add("  *(error reading: $_)*")
        }
        $lines.Add('')
    }

    # Forms
    $lines.Add("## Forms ($($forms.Count))")
    $lines.Add("")
    if ($forms.Count -gt 0) {
        foreach ($n in $forms) { $lines.Add("- ``$n``") }
    } else {
        $lines.Add('*(none)*')
    }
    $lines.Add('')

    # Reports
    $lines.Add("## Reports ($($reports.Count))")
    $lines.Add("")
    if ($reports.Count -gt 0) {
        foreach ($n in $reports) { $lines.Add("- ``$n``") }
    } else {
        $lines.Add('*(none)*')
    }
    $lines.Add('')

    # Queries
    $lines.Add("## Queries ($($queries.Count))")
    $lines.Add("")
    if ($queries.Count -gt 0) {
        foreach ($n in $queries) { $lines.Add("- ``$n``") }
    } else {
        $lines.Add('*(none)*')
    }
    $lines.Add('')

    # Macros
    if ($macros.Count -gt 0) {
        $lines.Add("## Macros ($($macros.Count))")
        $lines.Add("")
        foreach ($n in $macros) { $lines.Add("- ``$n``") }
        $lines.Add('')
    }

    $content = $lines -join "`n"
    [System.IO.File]::WriteAllText($OutputPath, $content, [System.Text.Encoding]::UTF8)

    Format-AccessOutput -AsJson:$AsJson -Data @{
        output_path = $OutputPath
        content     = $content
        status      = 'exported'
    }
}

# ── 2.8 Invoke-AccessSQL ─────────────────────────────────────────────────

function Invoke-AccessSQL {
    <#
    .SYNOPSIS
        Execute SQL against the database via DAO.
        SELECT returns rows; INSERT/UPDATE/DELETE returns affected_rows.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER SQL
        SQL statement to execute.
    .PARAMETER Limit
        Max rows to return for SELECT (default 500, max 10000).
    .PARAMETER ConfirmDestructive
        Required for DELETE/DROP/TRUNCATE/ALTER statements.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-AccessSQL -DbPath "C:\db.accdb" -SQL "SELECT * FROM Users"
        Invoke-AccessSQL -DbPath "C:\db.accdb" -SQL "DELETE FROM Users WHERE ID=1" -ConfirmDestructive
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$SQL,
        [int]$Limit = 500,
        [switch]$ConfirmDestructive,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $normalized = $SQL.Trim().ToUpper()

    if ($normalized.StartsWith('SELECT')) {
        $Limit = [math]::Max(1, [math]::Min($Limit, 10000))
        $rs = $null
        try {
            $rs = $db.OpenRecordset($SQL)
        } catch {
            $firstErr = $_
            try {
                $rs = $db.OpenRecordset($SQL, 2, $script:DB_SEE_CHANGES)  # dbOpenDynaset + dbSeeChanges
            } catch {
                throw $firstErr
            }
        }

        $fieldCount = $rs.Fields.Count
        [string[]]$fieldNames = for ($i = 0; $i -lt $fieldCount; $i++) { $rs.Fields($i).Name }
        $rows = [System.Collections.Generic.List[object]]::new()
        $truncated = $false

        if (-not $rs.EOF) {
            $rs.MoveFirst()
            while (-not $rs.EOF -and $rows.Count -lt $Limit) {
                $row = [ordered]@{}
                for ($i = 0; $i -lt $fieldCount; $i++) {
                    $row[$fieldNames[$i]] = ConvertTo-SafeValue -Value $rs.Fields($i).Value
                }
                $rows.Add([PSCustomObject]$row)
                $rs.MoveNext()
            }
            $truncated = -not $rs.EOF
        }
        $rs.Close()

        $result = [ordered]@{
            rows  = @($rows)
            count = $rows.Count
        }
        if ($truncated) { $result['truncated'] = $true }

        Format-AccessOutput -AsJson:$AsJson -Data $result
    }
    else {
        # Check for destructive SQL
        foreach ($prefix in $script:DESTRUCTIVE_PREFIXES) {
            if ($normalized.StartsWith($prefix)) {
                if (-not $ConfirmDestructive) {
                    $msg = "Destructive SQL detected. Use -ConfirmDestructive to execute: $($SQL.Substring(0, [math]::Min(100, $SQL.Length)))"
                    return Format-AccessOutput -AsJson:$AsJson -Data @{ error = $msg }
                }
                break
            }
        }

        try {
            $db.Execute($SQL)
        } catch {
            $firstErr = $_
            try {
                $db.Execute($SQL, $script:DB_SEE_CHANGES)
            } catch {
                throw $firstErr
            }
        }

        Format-AccessOutput -AsJson:$AsJson -Data @{ affected_rows = $db.RecordsAffected }
    }
}

# ── 2.9 Invoke-AccessSQLBatch ────────────────────────────────────────────

function Invoke-AccessSQLBatch {
    <#
    .SYNOPSIS
        Execute multiple SQL statements in a single call.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER Statements
        Array of hashtables: @{ sql = "..."; label = "optional label" }
    .PARAMETER StopOnError
        Stop at first error (default $true). Set $false to continue and report all.
    .PARAMETER ConfirmDestructive
        Required if any statement is DELETE/DROP/TRUNCATE/ALTER.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        $stmts = @(
            @{ sql = "INSERT INTO Users (Name) VALUES ('Alice')"; label = "add alice" }
            @{ sql = "SELECT COUNT(*) AS cnt FROM Users"; label = "count" }
        )
        Invoke-AccessSQLBatch -DbPath "C:\db.accdb" -Statements $stmts
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][object[]]$Statements,
        [bool]$StopOnError = $true,
        [switch]$ConfirmDestructive,
        [switch]$AsJson
    )

    if ($Statements.Count -eq 0) {
        return Format-AccessOutput -AsJson:$AsJson -Data @{ error = 'No SQL statements provided.' }
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    # Pre-scan for destructive SQL
    if (-not $ConfirmDestructive) {
        for ($i = 0; $i -lt $Statements.Count; $i++) {
            $sqlUpper = $Statements[$i].sql.Trim().ToUpper()
            foreach ($prefix in $script:DESTRUCTIVE_PREFIXES) {
                if ($sqlUpper.StartsWith($prefix)) {
                    $label = if ($Statements[$i].label) { $Statements[$i].label } else { "statement[$i]" }
                    return Format-AccessOutput -AsJson:$AsJson -Data @{
                        error = "Destructive SQL in '$label'. Use -ConfirmDestructive to execute."
                    }
                }
            }
        }
    }

    $results   = [System.Collections.Generic.List[object]]::new()
    $succeeded = 0
    $failed    = 0

    for ($i = 0; $i -lt $Statements.Count; $i++) {
        $sql   = $Statements[$i].sql.Trim()
        $label = $Statements[$i].label
        $entry = [ordered]@{ index = $i }
        if ($label) { $entry['label'] = $label }

        try {
            $sqlUpper = $sql.ToUpper()
            if ($sqlUpper.StartsWith('SELECT')) {
                $rs = $null
                try {
                    $rs = $db.OpenRecordset($sql)
                } catch {
                    $firstErr = $_
                    try {
                        $rs = $db.OpenRecordset($sql, 2, $script:DB_SEE_CHANGES)
                    } catch {
                        throw $firstErr
                    }
                }

                $fieldCount = $rs.Fields.Count
                [string[]]$fieldNames = for ($j = 0; $j -lt $fieldCount; $j++) { $rs.Fields($j).Name }
                $rows = [System.Collections.Generic.List[object]]::new()
                $selectLimit = 100

                if (-not $rs.EOF) {
                    $rs.MoveFirst()
                    while (-not $rs.EOF -and $rows.Count -lt $selectLimit) {
                        $row = [ordered]@{}
                        for ($j = 0; $j -lt $fieldCount; $j++) {
                            $row[$fieldNames[$j]] = ConvertTo-SafeValue -Value $rs.Fields($j).Value
                        }
                        $rows.Add([PSCustomObject]$row)
                        $rs.MoveNext()
                    }
                    $truncated = -not $rs.EOF
                }
                $rs.Close()

                $entry['status'] = 'ok'
                $entry['rows']   = @($rows)
                $entry['count']  = $rows.Count
                if ($truncated) { $entry['truncated'] = $true }
            }
            else {
                try {
                    $db.Execute($sql)
                } catch {
                    $firstErr = $_
                    try {
                        $db.Execute($sql, $script:DB_SEE_CHANGES)
                    } catch {
                        throw $firstErr
                    }
                }
                $entry['status']        = 'ok'
                $entry['affected_rows'] = $db.RecordsAffected
            }
            $succeeded++
        }
        catch {
            $entry['status'] = 'error'
            $entry['error']  = $_.ToString()
            $failed++
            if ($StopOnError) {
                $results.Add([PSCustomObject]$entry)
                return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
                    total      = $Statements.Count
                    succeeded  = $succeeded
                    failed     = $failed
                    stopped_at = $i
                    results    = @($results)
                })
            }
        }

        $results.Add([PSCustomObject]$entry)
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        total     = $Statements.Count
        succeeded = $succeeded
        failed    = $failed
        results   = @($results)
    })
}

# ── 2.10 Get-AccessTableInfo ─────────────────────────────────────────────

function Get-AccessTableInfo {
    <#
    .SYNOPSIS
        Get the structure of an Access table: fields, types, sizes, record count, linked info.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER TableName
        Name of the table.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessTableInfo -DbPath "C:\db.accdb" -TableName "Users"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$TableName,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    try {
        $td = $db.TableDefs($TableName)
    } catch {
        throw "Table '$TableName' not found: $_"
    }

    $isLinked = [bool]$td.Connect
    $fields = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $td.Fields.Count; $i++) {
        $fld   = $td.Fields($i)
        $ftype = $fld.Type
        $typeName = $script:DAO_FIELD_TYPE[[int]$ftype]
        if (-not $typeName) { $typeName = "Type$ftype" }

        # AutoNumber detection: Long (4) + dbAutoIncrField attribute (16)
        if ($ftype -eq 4 -and ($fld.Attributes -band $script:DB_AUTO_INCR_FIELD)) {
            $typeName = 'AutoNumber'
        }

        $fields.Add([PSCustomObject][ordered]@{
            name     = $fld.Name
            type     = $typeName
            size     = $fld.Size
            required = [bool]$fld.Required
        })
    }

    # Record count (may fail on linked tables)
    $recordCount = -1
    try {
        $recordCount = $td.RecordCount
        if ($recordCount -eq -1) {
            $rs = $db.OpenRecordset("SELECT COUNT(*) AS cnt FROM [$TableName]")
            $recordCount = $rs.Fields(0).Value
            $rs.Close()
        }
    } catch {}

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        table_name   = $TableName
        fields       = @($fields)
        record_count = $recordCount
        is_linked    = $isLinked
        source_table = if ($isLinked) { $td.SourceTableName } else { '' }
        connect      = if ($isLinked) { $td.Connect } else { '' }
    })
}

# ── 2.11 New-AccessTable ─────────────────────────────────────────────────

function New-AccessTable {
    <#
    .SYNOPSIS
        Create an Access table via DAO with full type support, defaults, descriptions, and primary key.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER TableName
        Name for the new table (must not already exist).
    .PARAMETER Fields
        Array of field definitions: @{ name="ID"; type="autonumber"; primary_key=$true },
        @{ name="Name"; type="text"; size=100; required=$true; default="Unknown"; description="User name" }
        Supported types: autonumber, autoincrement, long, integer, short, byte, text, memo,
        currency, double, single, float, datetime, date, boolean, yesno, bit, guid, ole, bigint.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        $fields = @(
            @{ name = "ID"; type = "autonumber"; primary_key = $true }
            @{ name = "Name"; type = "text"; size = 100; required = $true }
        )
        New-AccessTable -DbPath "C:\db.accdb" -TableName "Users" -Fields $fields
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][object[]]$Fields,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    # Check table doesn't exist
    $existing = for ($i = 0; $i -lt $db.TableDefs.Count; $i++) { $db.TableDefs($i).Name }
    if ($TableName -in $existing) {
        throw "Table '$TableName' already exists."
    }

    $td = $db.CreateTableDef($TableName)
    $pkFields      = [System.Collections.Generic.List[string]]::new()
    $createdFields = [System.Collections.Generic.List[object]]::new()

    foreach ($fdef in $Fields) {
        $name     = $fdef.name
        $ftype    = ($fdef.type ?? 'text').ToLower()
        $size     = [int]($fdef.size ?? 0)
        $required = [bool]($fdef.required)
        $pk       = [bool]($fdef.primary_key)

        $daoType = $script:FIELD_TYPE_MAP[$ftype]
        if ($null -eq $daoType) {
            $validTypes = ($script:FIELD_TYPE_MAP.Keys | Sort-Object -Unique) -join ', '
            throw "Unknown field type: '$ftype'. Valid types: $validTypes"
        }

        $isAuto = $ftype -in 'autonumber', 'autoincrement'

        # Text needs size
        if ($daoType -eq 10 -and $size -eq 0) { $size = 255 }

        $fld = if ($size -gt 0) {
            $td.CreateField($name, $daoType, $size)
        } else {
            $td.CreateField($name, $daoType)
        }

        if ($isAuto) {
            $fld.Attributes = $fld.Attributes -bor $script:DB_AUTO_INCR_FIELD
        }

        $fld.Required = $required -or $pk

        $td.Fields.Append($fld)

        if ($pk) { $pkFields.Add($name) }

        $createdFields.Add([PSCustomObject][ordered]@{
            name = $name
            type = $ftype
            size = if ($size -gt 0) { $size } else { $null }
        })
    }

    # Create primary key index
    if ($pkFields.Count -gt 0) {
        $idx = $td.CreateIndex('PrimaryKey')
        $idx.Primary = $true
        $idx.Unique  = $true
        foreach ($pkName in $pkFields) {
            $idxFld = $idx.CreateField($pkName)
            $idx.Fields.Append($idxFld)
        }
        $td.Indexes.Append($idx)
    }

    $db.TableDefs.Append($td)
    $db.TableDefs.Refresh()

    # Set defaults and descriptions via field properties (post-creation)
    foreach ($fdef in $Fields) {
        $name = $fdef.name
        if ($null -ne $fdef.default) {
            try {
                Set-FieldProperty -Db $db -TableName $TableName -FieldName $name -PropertyName 'DefaultValue' -Value ([string]$fdef.default)
            } catch {
                Write-Warning "Error setting default for ${TableName}.${name}: $_"
            }
        }
        if ($null -ne $fdef.description) {
            try {
                Set-FieldProperty -Db $db -TableName $TableName -FieldName $name -PropertyName 'Description' -Value $fdef.description
            } catch {
                Write-Warning "Error setting description for ${TableName}.${name}: $_"
            }
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        table_name  = $TableName
        fields      = @($createdFields)
        primary_key = @($pkFields)
        status      = 'created'
    })
}

# ═══════════════════════════════════════════════════════════════════════════
# PHASE 3 — VBE/VBA Operations
# ═══════════════════════════════════════════════════════════════════════════

# ── Internal helpers ──────────────────────────────────────────────────────

function Test-TextMatch {
    <#
    .SYNOPSIS
        Internal: match needle against haystack (plain substring or regex).
    #>
    param(
        [Parameter(Mandatory)][string]$Needle,
        [Parameter(Mandatory)][string]$Haystack,
        [bool]$MatchCase = $false,
        [bool]$UseRegex = $false
    )

    if ($UseRegex) {
        $opts = [System.Text.RegularExpressions.RegexOptions]::None
        if (-not $MatchCase) { $opts = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
        return [regex]::IsMatch($Haystack, $Needle, $opts)
    }
    if (-not $MatchCase) {
        return $Haystack.IndexOf($Needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
    }
    return $Haystack.Contains($Needle)
}

function Get-CodeModule {
    <#
    .SYNOPSIS
        Internal: Get cached VBE CodeModule COM object for a module/form/report.
    #>
    param(
        [Parameter(Mandatory)]$App,
        [Parameter(Mandatory)][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName
    )

    if (-not $script:VBE_PREFIX.ContainsKey($ObjectType)) {
        throw "object_type '$ObjectType' does not support VBE. Use 'module', 'form', or 'report'."
    }

    $cacheKey = "${ObjectType}:${ObjectName}"
    $cm = $script:AccessSession.CmCache[$cacheKey]
    if ($null -ne $cm) { return $cm }

    $compName = $script:VBE_PREFIX[$ObjectType] + $ObjectName
    try {
        $project = $App.VBE.VBProjects(1)
        $component = $project.VBComponents($compName)
        $cm = $component.CodeModule
        $script:AccessSession.CmCache[$cacheKey] = $cm
        return $cm
    } catch {
        $script:AccessSession.CmCache.Remove($cacheKey)
        throw "Cannot access CodeModule '$compName'. Is 'Trust access to the VBA project object model' enabled in Access Trust Center? Error: $_"
    }
}

function Get-AllModuleCode {
    <#
    .SYNOPSIS
        Internal: Get full module text using VbeCodeCache.
    #>
    param(
        [Parameter(Mandatory)]$CodeModule,
        [Parameter(Mandatory)][string]$CacheKey
    )

    if (-not $script:AccessSession.VbeCodeCache.ContainsKey($CacheKey)) {
        $total = $CodeModule.CountOfLines
        $text = if ($total -gt 0) { $CodeModule.Lines(1, $total) } else { '' }
        $script:AccessSession.VbeCodeCache[$CacheKey] = $text
    }
    return $script:AccessSession.VbeCodeCache[$CacheKey]
}

# ── 3.3 Get-AccessVbeLine ────────────────────────────────────────────────

function Get-AccessVbeLine {
    <#
    .SYNOPSIS
        Read a range of lines from a VBE CodeModule.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER StartLine
        First line to read (1-based).
    .PARAMETER Count
        Number of lines to read.
    .EXAMPLE
        Get-AccessVbeLine -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -StartLine 1 -Count 10
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('module','form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][int]$StartLine,
        [Parameter(Mandatory)][int]$Count
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $cacheKey = "${ObjectType}:${ObjectName}"
    $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey
    $allLines = $allCode.Split("`n")
    $total = $allLines.Count

    if ($StartLine -lt 1 -or $StartLine -gt $total) {
        throw "start_line $StartLine out of range (1-$total)"
    }
    $actual = [math]::Min($Count, $total - $StartLine + 1)
    $result = $allLines[($StartLine - 1) .. ($StartLine - 1 + $actual - 1)] -join "`n"
    return $result.TrimEnd("`r")
}

# ── 3.4 Get-AccessVbeProc ────────────────────────────────────────────────

function Get-AccessVbeProc {
    <#
    .SYNOPSIS
        Extract a procedure by name from a VBE module.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER ProcName
        Name of the procedure (Sub/Function/Property).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessVbeProc -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -ProcName "MyFunc"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('module','form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][string]$ProcName,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName

    try {
        $start = $cm.ProcStartLine($ProcName, 0)  # 0 = vbext_pk_Proc
        $body  = $cm.ProcBodyLine($ProcName, 0)
        $count = $cm.ProcCountLines($ProcName, 0)
    } catch {
        throw "Procedure '$ProcName' not found in '$ObjectName': $_"
    }

    $cacheKey = "${ObjectType}:${ObjectName}"
    $allLines = (Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey).Split("`n")
    $total = $allLines.Count
    $count = [math]::Min($count, $total - $start + 1)
    $code = ($allLines[($start - 1) .. ($start - 1 + $count - 1)] -join "`n").TrimEnd("`r")

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        proc_name  = $ProcName
        start_line = $start
        body_line  = $body
        count      = $count
        code       = $code
    })
}

# ── 3.5 Get-AccessVbeModuleInfo ──────────────────────────────────────────

function Get-AccessVbeModuleInfo {
    <#
    .SYNOPSIS
        Enumerate all procedures in a VBE module with their positions.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessVbeModuleInfo -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('module','form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $cacheKey = "${ObjectType}:${ObjectName}"
    $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey
    $allLines = $allCode.Split("`n")
    $total = $allLines.Count

    $procs = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    for ($i = 0; $i -lt $total; $i++) {
        $line = $allLines[$i].Trim()
        if ($line -match '^(?:Public\s+|Private\s+|Friend\s+)?(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+(\w+)') {
            $pname = $Matches[1]
            if (-not $seen.Add($pname)) { continue }
            try {
                $pstart = $cm.ProcStartLine($pname, 0)
                $pbody  = $cm.ProcBodyLine($pname, 0)
                $pcount = $cm.ProcCountLines($pname, 0)
                $pcount = [math]::Min($pcount, $total - $pstart + 1)
                $procs.Add([PSCustomObject][ordered]@{
                    name       = $pname
                    start_line = $pstart
                    body_line  = $pbody
                    count      = $pcount
                })
            } catch {
                $procs.Add([PSCustomObject][ordered]@{
                    name       = $pname
                    start_line = ($i + 1)
                })
            }
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        total_lines = $total
        procs       = @($procs)
    })
}

# ── 3.6 Set-AccessVbeLine ────────────────────────────────────────────────

function Set-AccessVbeLine {
    <#
    .SYNOPSIS
        Replace lines in a VBE module. count=0 inserts without deleting. Empty NewCode deletes only.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER StartLine
        First line to replace (1-based).
    .PARAMETER Count
        Number of lines to delete before inserting.
    .PARAMETER NewCode
        Code to insert at StartLine (can be multiline).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessVbeLine -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -StartLine 5 -Count 3 -NewCode "' replaced lines"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('module','form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][int]$StartLine,
        [int]$Count = 0,
        [string]$NewCode = '',
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $total = $cm.CountOfLines

    if ($StartLine -lt 1 -or $StartLine -gt ($total + 1)) {
        throw "start_line $StartLine out of range (1-$total)"
    }

    $clamped = $false
    if ($Count -gt 0) {
        $maxCount = $total - $StartLine + 1
        if ($Count -gt $maxCount) { $Count = $maxCount; $clamped = $true }
        $cm.DeleteLines($StartLine, $Count)
    }

    $inserted = 0
    if ($NewCode) {
        $normalized = $NewCode -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
        $cm.InsertLines($StartLine, $normalized)
        $inserted = $NewCode.Split("`n").Count
    }

    # Invalidate cache
    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines
    $clampNote = if ($clamped) { ' (count clamped to module boundary)' } else { '' }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status        = "Lines $StartLine replaced ($Count deleted, $inserted inserted)$clampNote -> module now has $newTotal lines"
        deleted       = $Count
        inserted      = $inserted
        new_total     = $newTotal
    })
}

# ── 3.7 Set-AccessVbeProc ────────────────────────────────────────────────

function Set-AccessVbeProc {
    <#
    .SYNOPSIS
        Replace an entire procedure by name. Auto-locates via ProcStartLine/ProcCountLines.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER ProcName
        Name of the procedure to replace.
    .PARAMETER NewCode
        New code for the procedure. Empty string deletes it.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessVbeProc -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -ProcName "OldFunc" -NewCode "Public Sub OldFunc()`r`n  MsgBox ""Hello""`r`nEnd Sub"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('module','form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][string]$ProcName,
        [string]$NewCode = '',
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath

    # Close form/report if open in design (avoids COM conflicts with VBE)
    if ($ObjectType -in 'form', 'report') {
        $acObjType = if ($ObjectType -eq 'form') { $script:AC_FORM } else { $script:AC_REPORT }
        try { $app.DoCmd.Close($acObjType, $ObjectName, $script:AC_SAVE_YES) } catch {}
    }

    # Invalidate cache in case CodeModule is stale
    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.CmCache.Remove($cacheKey)

    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    try {
        $start = $cm.ProcStartLine($ProcName, 0)
        $count = $cm.ProcCountLines($ProcName, 0)
    } catch {
        throw "Procedure '$ProcName' not found in '$ObjectName': $_"
    }

    $total = $cm.CountOfLines
    $count = [math]::Min($count, $total - $start + 1)

    $cm.DeleteLines($start, $count)

    $inserted = 0
    if ($NewCode) {
        $normalized = $NewCode -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
        $cm.InsertLines($start, $normalized)
        $inserted = $NewCode.Split("`n").Count
    }

    # Invalidate caches
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)
    $script:AccessSession.CmCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines
    $action = if ($NewCode) { 'replaced' } else { 'deleted' }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status   = "Proc '$ProcName' $action ($count deleted, $inserted inserted) -> module now has $newTotal lines"
        action   = $action
        proc     = $ProcName
        deleted  = $count
        inserted = $inserted
        new_total = $newTotal
    })
}

# ── 3.7b Patch-AccessVbeProc ─────────────────────────────────────────────

function Test-WsNormalizedMatch {
    <#
    .SYNOPSIS
        Whitespace-tolerant matching: strips leading whitespace from each line
        and does a sliding-window search. Returns hashtable with start/end 0-based
        line indices, or $null if no match.
    #>
    param(
        [string]$ProcCode,
        [string]$FindText
    )
    $procLines = $ProcCode -split "`r?`n"
    $findLines = @($FindText -split "`r?`n")
    # Remove empty trailing lines from find text
    while ($findLines.Count -gt 0 -and -not $findLines[-1].Trim()) {
        $findLines = $findLines[0..($findLines.Count - 2)]
    }
    if ($findLines.Count -eq 0) { return $null }

    $procStripped = @($procLines | ForEach-Object { $_.TrimStart() })
    $findStripped = @($findLines | ForEach-Object { $_.TrimStart() })
    $window = $findStripped.Count

    for ($i = 0; $i -le ($procStripped.Count - $window); $i++) {
        $match = $true
        for ($j = 0; $j -lt $window; $j++) {
            if ($procStripped[$i + $j] -cne $findStripped[$j]) {
                $match = $false
                break
            }
        }
        if ($match) {
            return @{ start = $i; end = ($i + $window - 1) }
        }
    }
    return $null
}

function Get-ClosestMatchContext {
    <#
    .SYNOPSIS
        When both exact and ws-normalized match fail, finds the most similar line
        using character-level similarity and returns a contextual snippet.
    #>
    param(
        [string]$ProcCode,
        [string]$FindText,
        [string]$ProcName
    )
    $procLines = $ProcCode -split "`r?`n"
    $findLines = @(($FindText -split "`r?`n") | Where-Object { $_.Trim() })
    if ($findLines.Count -eq 0) { return "Empty find text in proc '$ProcName'" }

    $ref = $findLines[0].Trim()
    $bestRatio = 0.0
    $bestIdx = 0

    for ($i = 0; $i -lt $procLines.Count; $i++) {
        $candidate = $procLines[$i].Trim()
        if (-not $candidate) { continue }
        # Longest Common Subsequence ratio (simplified SequenceMatcher)
        $shorter = if ($ref.Length -lt $candidate.Length) { $ref } else { $candidate }
        $longer  = if ($ref.Length -lt $candidate.Length) { $candidate } else { $ref }
        $matchCount = 0
        $usedIdx = -1
        foreach ($ch in $shorter.ToCharArray()) {
            $pos = $longer.IndexOf($ch, $usedIdx + 1)
            if ($pos -ge 0) { $matchCount++; $usedIdx = $pos }
        }
        $ratio = if (($ref.Length + $candidate.Length) -gt 0) { (2.0 * $matchCount) / ($ref.Length + $candidate.Length) } else { 0 }
        if ($ratio -gt $bestRatio) {
            $bestRatio = $ratio
            $bestIdx = $i
        }
    }

    # Build context: 3 lines around best candidate
    $ctxStart = [math]::Max(0, $bestIdx - 1)
    $ctxEnd   = [math]::Min($procLines.Count - 1, $bestIdx + 1)
    $contextLines = @()
    for ($j = $ctxStart; $j -le $ctxEnd; $j++) {
        $marker = if ($j -eq $bestIdx) { '>>>' } else { '   ' }
        $contextLines += "  $marker L$($j + 1): $($procLines[$j].TrimEnd())"
    }
    $pct = [math]::Round($bestRatio * 100)
    $refSnippet = if ($ref.Length -gt 80) { $ref.Substring(0, 80) } else { $ref }
    return "Best match (${pct}% similar) near line $($bestIdx + 1) of '${ProcName}':`n$($contextLines -join "`n")`n  Looking for: '$refSnippet'"
}

function Update-AccessVbeProc {
    <#
    .SYNOPSIS
        Surgically patch code within a procedure using find/replace with whitespace tolerance.
    .DESCRIPTION
        Applies one or more patches (find/replace pairs) to a procedure without rewriting the entire proc.
        Three-layer matching: (1) exact string, (2) whitespace-normalized fallback, (3) context error reporting.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER ProcName
        Name of the procedure to patch.
    .PARAMETER Patches
        Array of hashtables, each with 'find' and 'replace' keys.
        Example: @(@{find='old code'; replace='new code'}, @{find='more old'; replace='more new'})
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Patch-AccessVbeProc -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" `
            -ProcName "MyFunc" -Patches @(@{find='MsgBox "Old"'; replace='MsgBox "New"'})
    .EXAMPLE
        # Multiple patches in one call
        $patches = @(
            @{ find = '    x = 1'; replace = '    x = 2' }
            @{ find = '    y = 3'; replace = '    y = 4' }
        )
        Patch-AccessVbeProc -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" `
            -ProcName "Form_Load" -Patches $patches
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('module','form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][string]$ProcName,
        [Parameter(Mandatory)][array]$Patches,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath

    # Close form/report if open in design (avoids COM conflicts with VBE)
    if ($ObjectType -in 'form', 'report') {
        $acObjType = if ($ObjectType -eq 'form') { $script:AC_FORM } else { $script:AC_REPORT }
        try { $app.DoCmd.Close($acObjType, $ObjectName, $script:AC_SAVE_YES) } catch {}
    }

    # Invalidate cache
    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.CmCache.Remove($cacheKey)

    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName

    # Locate procedure — try kind=0 (Sub/Function), fallback to kind=3 (Property)
    $kind = 0
    try {
        $start = $cm.ProcStartLine($ProcName, 0)
        $count = $cm.ProcCountLines($ProcName, 0)
    } catch {
        try {
            $start = $cm.ProcStartLine($ProcName, 3)
            $count = $cm.ProcCountLines($ProcName, 3)
            $kind = 3
        } catch {
            throw "Procedure '$ProcName' not found in '$ObjectName': $_"
        }
    }

    $total = $cm.CountOfLines
    $count = [math]::Min($count, $total - $start + 1)

    # Get current proc code
    $procCode = $cm.Lines($start, $count)
    $backupCode = $procCode

    # Apply patches sequentially
    $applied = 0
    $notFound = @()
    $wsFallbackNotes = @()

    for ($pi = 0; $pi -lt $Patches.Count; $pi++) {
        $patch = $Patches[$pi]
        $findText    = "$($patch['find'])"
        $replaceText = if ($patch.ContainsKey('replace')) { "$($patch['replace'])" } else { '' }

        # Layer 1: Exact match
        if ($procCode.Contains($findText)) {
            $idx = $procCode.IndexOf($findText)
            $procCode = $procCode.Substring(0, $idx) + $replaceText + $procCode.Substring($idx + $findText.Length)
            $applied++
        }
        else {
            # Layer 2: Whitespace-normalized fallback
            $wsMatch = Test-WsNormalizedMatch -ProcCode $procCode -FindText $findText
            if ($null -ne $wsMatch) {
                $codeLines = @($procCode -split "`r?`n")
                $replaceNorm = $replaceText
                if ($replaceNorm -and -not $replaceNorm.EndsWith("`n")) {
                    $replaceNorm += "`r`n"
                }
                $before = if ($wsMatch.start -gt 0) { $codeLines[0..($wsMatch.start - 1)] } else { @() }
                $after  = if ($wsMatch.end -lt ($codeLines.Count - 1)) { $codeLines[($wsMatch.end + 1)..($codeLines.Count - 1)] } else { @() }
                $middle = if ($replaceNorm) { @($replaceNorm.TrimEnd("`r`n")) } else { @() }
                $procCode = ($before + $middle + $after) -join "`r`n"
                $applied++
                $wsFallbackNotes += "patch[$pi]: matched via ws-normalized fallback"
            }
            else {
                # Layer 3: Context error reporting
                $ctx = Get-ClosestMatchContext -ProcCode $procCode -FindText $findText -ProcName $ProcName
                $notFound += "patch[$pi]: not found. $ctx"
            }
        }
    }

    if ($applied -eq 0) {
        $errMsg = "NOOP: no patches matched in '$ProcName'. Errors:`n$($notFound -join "`n")"
        return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            status     = $errMsg
            applied    = 0
            total      = $Patches.Count
            not_found  = $notFound
        })
    }

    # Strip Option lines if proc is NOT at the top of the module
    $optionWarnings = @()
    if ($start -gt 5) {
        $optionRe = '^\s*Option\s+(Explicit|Compare\s+\w+)\s*$'
        $cleanLines = @()
        foreach ($line in ($procCode -split "`r?`n")) {
            if ($line -match $optionRe) {
                $optionWarnings += "Stripped misplaced Option line: '$($line.Trim())'"
            } else {
                $cleanLines += $line
            }
        }
        $procCode = $cleanLines -join "`r`n"
    }

    # Replace entire proc with patched code
    try {
        $cm.DeleteLines($start, $count)
        if ($procCode.Trim()) {
            $normalized = $procCode -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
            $cm.InsertLines($start, $normalized)
        }
    } catch {
        # Rollback: restore backup
        try { $cm.InsertLines($start, $backupCode) } catch {}
        throw "Patch failed (rolled back): $_"
    }

    # Invalidate caches
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)
    $script:AccessSession.CmCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines
    $newCount = try { $cm.ProcCountLines($ProcName, $kind) } catch { 0 }

    # Build result
    $resultParts = @("OK: $applied/$($Patches.Count) patches applied in '$ProcName' ($count -> $newCount lines) -> module now has $newTotal lines")
    if ($wsFallbackNotes.Count -gt 0) { $resultParts += "WS-fallback: $($wsFallbackNotes -join '; ')" }
    if ($optionWarnings.Count -gt 0)  { $resultParts += $optionWarnings }
    if ($notFound.Count -gt 0)        { $resultParts += "Not found:"; $resultParts += $notFound }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status          = $resultParts -join "`n"
        applied         = $applied
        total           = $Patches.Count
        proc            = $ProcName
        old_lines       = $count
        new_lines       = $newCount
        module_lines    = $newTotal
        ws_fallback     = $wsFallbackNotes
        not_found       = $notFound
    })
}

# ── 3.8 Add-AccessVbeCode ────────────────────────────────────────────────

function Add-AccessVbeCode {
    <#
    .SYNOPSIS
        Append code at the end of a VBE module.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER Code
        Code to append (can be multiline).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-AccessVbeCode -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -Code "Public Sub NewSub()`r`nEnd Sub"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('module','form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][string]$Code,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $total = $cm.CountOfLines

    $normalized = $Code -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
    $cm.InsertLines($total + 1, $normalized)
    $inserted = $Code.Split("`n").Count

    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status    = "$inserted lines appended -> module now has $newTotal lines"
        inserted  = $inserted
        new_total = $newTotal
    })
}

# ── 3.9 Find-AccessVbeText ───────────────────────────────────────────────

function Find-AccessVbeText {
    <#
    .SYNOPSIS
        Search for text (or regex) in a VBE module.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Find-AccessVbeText -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -SearchText "MsgBox"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('module','form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $cacheKey = "${ObjectType}:${ObjectName}"
    $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey

    if (-not $allCode) {
        return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{ found = $false; match_count = 0; matches = @() })
    }

    $matchList = [System.Collections.Generic.List[object]]::new()
    $lineNum = 0
    foreach ($line in $allCode.Split("`n")) {
        $lineNum++
        if (Test-TextMatch -Needle $SearchText -Haystack $line -MatchCase:$MatchCase -UseRegex:$UseRegex) {
            $matchList.Add([PSCustomObject][ordered]@{
                line    = $lineNum
                content = $line.TrimEnd("`r")
            })
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        found       = ($matchList.Count -gt 0)
        match_count = $matchList.Count
        matches     = @($matchList)
    })
}

# ── 3.10 Search-AccessVbe ────────────────────────────────────────────────

function Search-AccessVbe {
    <#
    .SYNOPSIS
        Search for text (or regex) across ALL VBA modules, forms, and reports.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER MaxResults
        Maximum total matches to return (default 100).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Search-AccessVbe -DbPath "C:\db.accdb" -SearchText "DoCmd"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [int]$MaxResults = 100,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $objects = Get-AccessObject -DbPath $DbPath -ObjectType all
    $results = [System.Collections.Generic.List[object]]::new()
    $total = 0
    $truncated = $false

    foreach ($objType in @('module', 'form', 'report')) {
        if ($truncated) { break }
        $names = @($objects.$objType)
        foreach ($objName in $names) {
            if ($truncated) { break }
            try {
                $cm = Get-CodeModule -App $app -ObjectType $objType -ObjectName $objName
                $cacheKey = "${objType}:${objName}"
                $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey
                if (-not $allCode) { continue }

                $objMatches = [System.Collections.Generic.List[object]]::new()
                $lineNum = 0
                foreach ($line in $allCode.Split("`n")) {
                    $lineNum++
                    if (Test-TextMatch -Needle $SearchText -Haystack $line -MatchCase:$MatchCase -UseRegex:$UseRegex) {
                        $objMatches.Add([PSCustomObject][ordered]@{
                            line    = $lineNum
                            content = $line.TrimEnd("`r")
                        })
                        $total++
                        if ($total -ge $MaxResults) { $truncated = $true; break }
                    }
                }
                if ($objMatches.Count -gt 0) {
                    $results.Add([PSCustomObject][ordered]@{
                        object_type = $objType
                        object_name = $objName
                        matches     = @($objMatches)
                    })
                }
            } catch { continue }
        }
    }

    $out = [ordered]@{ total_matches = $total; results = @($results) }
    if ($truncated) { $out['truncated'] = $true }
    Format-AccessOutput -AsJson:$AsJson -Data $out
}

# ── 3.11 Search-AccessQuery ──────────────────────────────────────────────

function Search-AccessQuery {
    <#
    .SYNOPSIS
        Search for text (or regex) in the SQL of all queries.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER MaxResults
        Maximum results to return (default 100).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Search-AccessQuery -DbPath "C:\db.accdb" -SearchText "Users"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [int]$MaxResults = 100,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()
    $results = [System.Collections.Generic.List[object]]::new()
    $total = 0

    foreach ($qd in $db.QueryDefs) {
        $name = $qd.Name
        if ($name.StartsWith('~')) { continue }
        $sql = $qd.SQL
        if (Test-TextMatch -Needle $SearchText -Haystack $sql -MatchCase:$MatchCase -UseRegex:$UseRegex) {
            $results.Add([PSCustomObject][ordered]@{
                query_name = $name
                sql        = $sql.Trim()
            })
            $total++
            if ($total -ge $MaxResults) { break }
        }
    }

    $out = [ordered]@{ total_matches = $total; results = @($results) }
    if ($total -ge $MaxResults) { $out['truncated'] = $true }
    Format-AccessOutput -AsJson:$AsJson -Data $out
}

# ── 3.12 Invoke-AccessMacro ──────────────────────────────────────────────

function Invoke-AccessMacro {
    <#
    .SYNOPSIS
        Run an Access macro by name.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER MacroName
        Name of the macro to run.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-AccessMacro -DbPath "C:\db.accdb" -MacroName "AutoExec"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$MacroName,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    try {
        $app.DoCmd.RunMacro($MacroName)
    } catch {
        throw "Error running macro '$MacroName': $_"
    }

    Format-AccessOutput -AsJson:$AsJson -Data @{
        macro_name = $MacroName
        status     = 'executed'
    }
}

# ── 3.13 Invoke-AccessVba ────────────────────────────────────────────────

function Invoke-AccessVba {
    <#
    .SYNOPSIS
        Call a VBA Sub/Function via Application.Run or Forms COM access.
    .DESCRIPTION
        Supports two syntaxes:
        - 'ModuleName.ProcName' or 'ProcName' -> Application.Run (standard modules)
        - 'Forms.FormName.Method' -> COM Forms() access (form must be open)
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER Procedure
        Procedure name or qualified path (e.g., "Module1.MyFunc" or "Forms.Form1.Calculate").
    .PARAMETER Arguments
        Arguments to pass to the procedure (max 30 for Application.Run).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-AccessVba -DbPath "C:\db.accdb" -Procedure "Module1.Calculate" -Arguments @(42)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$Procedure,
        [object[]]$Arguments,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $callArgs = if ($Arguments) { $Arguments } else { @() }

    if ($callArgs.Count -gt 30) {
        throw 'Application.Run supports max 30 arguments.'
    }

    # Forms.FormName.Method -> direct COM access
    if ($Procedure.Contains('.')) {
        $parts = $Procedure.Split('.', 3)
        if ($parts[0] -eq 'Forms' -and $parts.Count -eq 3) {
            $formName   = $parts[1]
            $methodName = $parts[2]
            try {
                $form = $app.Forms($formName)
                if ($callArgs.Count -gt 0) {
                    $result = $form.GetType().InvokeMember(
                        $methodName,
                        [System.Reflection.BindingFlags]::InvokeMethod,
                        $null, $form, $callArgs
                    )
                } else {
                    # Try method call, fall back to property
                    try {
                        $result = $form.GetType().InvokeMember(
                            $methodName,
                            [System.Reflection.BindingFlags]::InvokeMethod,
                            $null, $form, @()
                        )
                    } catch {
                        $result = $form.GetType().InvokeMember(
                            $methodName,
                            [System.Reflection.BindingFlags]::GetProperty,
                            $null, $form, @()
                        )
                    }
                }
            } catch {
                throw "Error calling Forms('$formName').$methodName : $_. Make sure the form is open."
            }

            return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
                procedure = $Procedure
                result    = $result
                status    = 'executed'
            })
        }
    }

    # Standard Application.Run
    # PowerShell COM late binding handles optional params natively
    try {
        $result = switch ($callArgs.Count) {
            0  { $app.Run($Procedure) }
            1  { $app.Run($Procedure, $callArgs[0]) }
            2  { $app.Run($Procedure, $callArgs[0], $callArgs[1]) }
            3  { $app.Run($Procedure, $callArgs[0], $callArgs[1], $callArgs[2]) }
            4  { $app.Run($Procedure, $callArgs[0], $callArgs[1], $callArgs[2], $callArgs[3]) }
            5  { $app.Run($Procedure, $callArgs[0], $callArgs[1], $callArgs[2], $callArgs[3], $callArgs[4]) }
            default {
                # Build args array for Invoke
                $invokeArgs = @($Procedure) + $callArgs
                $app.GetType().InvokeMember('Run', [System.Reflection.BindingFlags]::InvokeMethod, $null, $app, $invokeArgs)
            }
        }
    } catch {
        throw "Error running '$Procedure': $_"
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        procedure = $Procedure
        result    = $result
        status    = 'executed'
    })
}

# ── 3.14 Invoke-AccessEval ───────────────────────────────────────────────

function Invoke-AccessEval {
    <#
    .SYNOPSIS
        Evaluate a VBA/Access expression via Application.Eval.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER Expression
        Expression to evaluate (e.g., "Date()", "DLookup(...)").
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-AccessEval -DbPath "C:\db.accdb" -Expression "Date()"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$Expression,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    try {
        $result = $app.Eval($Expression)
    } catch {
        throw "Error evaluating '$Expression': $_"
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        expression = $Expression
        result     = $result
        status     = 'evaluated'
    })
}

# ── 3.15 Test-AccessVbaCompile ────────────────────────────────────────────

function Test-AccessVbaCompile {
    <#
    .SYNOPSIS
        Compile and save all VBA modules. Returns status and error location if compilation fails.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Test-AccessVbaCompile -DbPath "C:\db.accdb"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath

    try {
        $app.RunCommand($script:AC_CMD_COMPILE)
    } catch {
        # Try to get error location from VBE
        $errLoc = $null
        try {
            $pane = $app.VBE.ActiveCodePane
            if ($null -ne $pane) {
                $cm = $pane.CodeModule
                $startLine = 0; $startCol = 0; $endLine = 0; $endCol = 0
                $pane.GetSelection([ref]$startLine, [ref]$startCol, [ref]$endLine, [ref]$endCol)
                $errCode = ''
                if ($startLine -gt 0 -and $cm.CountOfLines -ge $startLine) {
                    $errCode = $cm.Lines($startLine, 1)
                }
                $errLoc = [ordered]@{
                    component = $cm.Parent.Name
                    line      = $startLine
                    code      = $errCode
                }
            }
        } catch {}

        $result = [ordered]@{
            status       = 'error'
            error_detail = "VBA compilation error: $_"
        }
        if ($errLoc) { $result['error_location'] = $errLoc }
        return Format-AccessOutput -AsJson:$AsJson -Data $result
    }

    # Invalidate caches — compilation may change module state
    $script:AccessSession.VbeCodeCache = @{}
    $script:AccessSession.CmCache = @{}

    Format-AccessOutput -AsJson:$AsJson -Data @{ status = 'compiled' }
}

# ── 3.16 Find-AccessUsage ────────────────────────────────────────────────

function Find-AccessUsage {
    <#
    .SYNOPSIS
        Search for a name across VBA code, query SQL, and form/report control properties.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER MaxResults
        Maximum total matches (default 200).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Find-AccessUsage -DbPath "C:\db.accdb" -SearchText "Users"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [int]$MaxResults = 200,
        [switch]$AsJson
    )

    # 1. VBA matches
    $vbaResult = Search-AccessVbe -DbPath $DbPath -SearchText $SearchText -MatchCase:$MatchCase -UseRegex:$UseRegex -MaxResults $MaxResults
    $vbaMatches = [System.Collections.Generic.List[object]]::new()
    foreach ($group in @($vbaResult.results)) {
        foreach ($m in @($group.matches)) {
            $vbaMatches.Add([PSCustomObject][ordered]@{
                object_type = $group.object_type
                object_name = $group.object_name
                line        = $m.line
                content     = $m.content
            })
        }
    }
    $total = $vbaMatches.Count
    $truncated = [bool]$vbaResult.truncated

    # 2. Query matches
    $queryMatches = [System.Collections.Generic.List[object]]::new()
    if (-not $truncated) {
        $remaining = $MaxResults - $total
        $qryResult = Search-AccessQuery -DbPath $DbPath -SearchText $SearchText -MatchCase:$MatchCase -UseRegex:$UseRegex -MaxResults $remaining
        foreach ($q in @($qryResult.results)) { $queryMatches.Add($q) }
        $total += $qryResult.total_matches
        $truncated = [bool]$qryResult.truncated
    }

    # 3. Control property matches — search form/report exports
    $controlMatches = [System.Collections.Generic.List[object]]::new()
    if (-not $truncated) {
        $app = $script:AccessSession.App
        $objects = Get-AccessObject -DbPath $DbPath -ObjectType all
        foreach ($objType in @('form', 'report')) {
            if ($truncated) { break }
            foreach ($objName in @($objects.$objType)) {
                if ($truncated) { break }
                try {
                    $tmp = [System.IO.Path]::GetTempFileName()
                    try {
                        $app.SaveAsText($script:AC_TYPE[$objType], $objName, $tmp)
                        $rawText = (Read-TempFile -Path $tmp).Content
                    } finally {
                        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                    }
                    foreach ($line in $rawText.Split("`n")) {
                        $stripped = $line.Trim()
                        foreach ($prop in $script:CONTROL_SEARCH_PROPS) {
                            if ($stripped.StartsWith("$prop =")) {
                                $valuePart = $stripped.Substring($prop.Length + 2).Trim()
                                if (Test-TextMatch -Needle $SearchText -Haystack $valuePart -MatchCase:$MatchCase -UseRegex:$UseRegex) {
                                    $controlMatches.Add([PSCustomObject][ordered]@{
                                        object_type = $objType
                                        object_name = $objName
                                        property    = $prop
                                        value       = $valuePart
                                    })
                                    $total++
                                    if ($total -ge $MaxResults) { $truncated = $true }
                                    break
                                }
                            }
                        }
                    }
                } catch { continue }
            }
        }
    }

    $out = [ordered]@{
        search_text     = $SearchText
        vba_matches     = @($vbaMatches)
        query_matches   = @($queryMatches)
        control_matches = @($controlMatches)
        total_matches   = $total
    }
    if ($truncated) { $out['truncated'] = $true }
    Format-AccessOutput -AsJson:$AsJson -Data $out
}

# ═══════════════════════════════════════════════════════════════════════════
# PHASE 4 — FORM / REPORT OPERATIONS
# ═══════════════════════════════════════════════════════════════════════════

# ── 4.1 Open-InDesignView (internal) ─────────────────────────────────────

function Open-InDesignView {
    <#
    .SYNOPSIS
        Open a form or report in Design view (internal helper).
        Uses $script:AccessSession.App directly for COM reliability.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName
    )
    try {
        if ($ObjectType -eq 'form') {
            $script:AccessSession.App.DoCmd.OpenForm($ObjectName, $script:AC_DESIGN)
        } else {
            $script:AccessSession.App.DoCmd.OpenReport($ObjectName, $script:AC_DESIGN)
        }
    } catch {
        throw "Cannot open '$ObjectName' in Design view. If it is open in another view, close it first. Error: $_"
    }
}

# ── 4.2 Get-DesignObject (internal) ──────────────────────────────────────

function Get-DesignObject {
    <#
    .SYNOPSIS
        Return the COM Form/Report object currently open in Design view (internal helper).
        Uses Screen.ActiveForm/ActiveReport — the Forms/Reports collection
        cannot be accessed reliably from dot-sourced functions due to a
        PowerShell COM marshaling issue.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName
    )
    $sessionApp = $script:AccessSession.App
    if ($ObjectType -eq 'form') {
        $result = $sessionApp.Screen.ActiveForm
    } else {
        $result = $sessionApp.Screen.ActiveReport
    }
    if ($null -eq $result -or $result.Name -ne $ObjectName) {
        throw "Cannot get '$ObjectName' ($ObjectType) — is it open in Design view?"
    }
    $result
}

# ── 4.3 Save-AndCloseDesign (internal) ───────────────────────────────────

function Save-AndCloseDesign {
    <#
    .SYNOPSIS
        Save and close a form/report open in Design view, invalidate caches (internal helper).
        Uses $script:AccessSession.App directly for COM reliability.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName
    )
    $acType = if ($ObjectType -eq 'form') { $script:AC_TYPE['form'] } else { $script:AC_TYPE['report'] }
    try {
        $script:AccessSession.App.DoCmd.Close($acType, $ObjectName, $script:AC_SAVE_YES)
    } catch {
        Write-Warning "Error closing '$ObjectName': $_"
    }
    # Invalidate caches
    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.ControlsCache.Remove($cacheKey) | Out-Null
    $script:AccessSession.CmCache.Remove($cacheKey) | Out-Null
    $script:AccessSession.VbeCodeCache.Remove($cacheKey) | Out-Null
}

# ── 4.4 ConvertFrom-ControlBlock (internal) ──────────────────────────────

function ConvertFrom-ControlBlock {
    <#
    .SYNOPSIS
        Parse the SaveAsText export of a form/report and extract control blocks (internal helper).
    .DESCRIPTION
        Returns a hashtable with:
          controls       — array of controls with properties and line positions
          form_indent    — indentation of the Begin Form/Report line
          ctrl_indent    — indent of the first control found (legacy compat)
          form_begin_idx — 0-based line index of Begin Form/Report
          form_end_idx   — 0-based line index of the closing End
    #>
    param(
        [Parameter(Mandatory)][string]$FormText
    )

    $lines = $FormText -split "`r?`n"
    $result = [ordered]@{
        controls       = @()
        form_indent    = ''
        ctrl_indent    = ''
        form_begin_idx = -1
        form_end_idx   = -1
    }

    # Known control type names for fast lookup
    $ctrlTypeNames = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]($script:CTRL_TYPE.Values),
        [System.StringComparer]::Ordinal
    )

    # 1. Locate "Begin Form" or "Begin Report"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $s = $lines[$i].TrimStart()
        if ($s -match '^Begin\s+(Form|Report)\s*$') {
            $raw = $lines[$i]
            $result['form_indent'] = $raw.Substring(0, $raw.Length - $raw.TrimStart().Length)
            $result['form_begin_idx'] = $i
            break
        }
    }
    if ($result['form_begin_idx'] -eq -1) { return $result }

    $formBegin = $result['form_begin_idx']

    # 2. Find the matching "End" (depth tracking, including "Property = Begin" blocks)
    $depth = 0
    for ($i = $formBegin; $i -lt $lines.Count; $i++) {
        $s = $lines[$i].TrimStart()
        if ($s -match '^Begin\b' -or $s -match '^\w+\s*=\s*Begin\s*$') {
            $depth++
        } elseif ($s -eq 'End') {
            $depth--
            if ($depth -eq 0) {
                $result['form_end_idx'] = $i
                break
            }
        }
    }
    if ($result['form_end_idx'] -eq -1) { return $result }

    # Container types whose children get a "parent" field
    $containerTypes = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@('Page', 'OptionGroup'),
        [System.StringComparer]::Ordinal
    )

    # 3. Scan all "Begin <TypeName>" blocks where TypeName is a known control type
    $controls = [System.Collections.Generic.List[object]]::new()
    $containerStack = [System.Collections.Generic.List[object]]::new()  # @{name; end_idx}
    $currentSection = ''
    $i = $formBegin + 1
    while ($i -lt $result['form_end_idx']) {
        # Clean up containers we've passed
        while ($containerStack.Count -gt 0 -and $i -gt $containerStack[$containerStack.Count - 1].end_idx) {
            $containerStack.RemoveAt($containerStack.Count - 1)
        }

        $raw = $lines[$i]
        $s = $raw.TrimStart()
        $indent = $raw.Substring(0, $raw.Length - $s.Length)

        # Skip ClassModule — contains VBA, not controls
        if ($s -match '^Begin\s+ClassModule\s*$') { break }

        # Track current section (Detail, FormHeader, FormFooter, PageHeader, PageFooter, etc.)
        if ($s -match '^Begin\s+Section\s*$') {
            # Look ahead for the section Name property
            $secName = ''
            $secDepth = 1
            for ($si = $i + 1; $si -lt $lines.Count; $si++) {
                $secLine = $lines[$si].TrimStart()
                if ($secLine -match '^Begin\b') { $secDepth++ }
                elseif ($secLine -eq 'End') {
                    $secDepth--
                    if ($secDepth -eq 0) { break }
                }
                if ($secDepth -eq 1) {
                    $secMatch = [regex]::Match($secLine, '^Name\s*=\s*"?([^"]*)"?\s*$')
                    if ($secMatch.Success) { $secName = $secMatch.Groups[1].Value; break }
                }
            }
            $currentSection = $secName
        }

        # Detect "Begin <TypeName>"
        $mCtrl = [regex]::Match($s, '^Begin\s+(\w+)\s*$')
        if ($mCtrl.Success -and $ctrlTypeNames.Contains($mCtrl.Groups[1].Value)) {
            $ctrlStart = $i
            $block = [System.Collections.Generic.List[string]]::new()
            $block.Add($lines[$i])
            $props = @{}
            $blkDepth = 1
            $ctrlEnd = $i
            $j = $i + 1
            while ($j -lt $lines.Count) {
                $bl = $lines[$j]
                $blS = $bl.TrimStart()
                $block.Add($bl)
                # Parse top-level properties only (depth == 1)
                if ($blkDepth -eq 1) {
                    $mProp = [regex]::Match($blS, '^(\w+)\s*=(.*)')
                    if ($mProp.Success) {
                        $props[$mProp.Groups[1].Value] = $mProp.Groups[2].Value.Trim().Trim('"')
                    }
                }
                if ($blS -match '^Begin\b') {
                    $blkDepth++
                } elseif ($blS -eq 'End') {
                    $blkDepth--
                    if ($blkDepth -eq 0) {
                        $ctrlEnd = $j
                        break
                    }
                }
                $j++
            }

            $name = if ($props.ContainsKey('Name')) { $props['Name'] } elseif ($props.ContainsKey('ControlName')) { $props['ControlName'] } else { '' }
            $ctype = -1
            if ($props.ContainsKey('ControlType')) {
                $parsed = 0
                if ([int]::TryParse($props['ControlType'], [ref]$parsed)) { $ctype = $parsed }
            }

            $rawText = $block -join "`r`n"
            $fmtCount = ($block | Where-Object { $_ -match '^\s+ConditionalFormat\d*\s*=\s*Begin\s*$' }).Count

            if (-not $result['ctrl_indent'] -and $name) {
                $result['ctrl_indent'] = $indent
            }

            $ctrlEntry = [ordered]@{
                name           = $name
                control_type   = $ctype
                type_name      = if ($script:CTRL_TYPE.ContainsKey($ctype)) { $script:CTRL_TYPE[$ctype] } else { $mCtrl.Groups[1].Value }
                caption        = if ($props.ContainsKey('Caption')) { $props['Caption'] } else { '' }
                control_source = if ($props.ContainsKey('ControlSource')) { $props['ControlSource'] } else { '' }
                left           = if ($props.ContainsKey('Left'))   { $props['Left'] }   else { '' }
                top            = if ($props.ContainsKey('Top'))    { $props['Top'] }    else { '' }
                width          = if ($props.ContainsKey('Width'))  { $props['Width'] }  else { '' }
                height         = if ($props.ContainsKey('Height')) { $props['Height'] } else { '' }
                visible        = if ($props.ContainsKey('Visible')){ $props['Visible'] } else { '' }
                section        = $currentSection
                parent         = if ($containerStack.Count -gt 0) { $containerStack[$containerStack.Count - 1].name } else { '' }
                start_line     = $ctrlStart + 1   # 1-based
                end_line       = $ctrlEnd + 1     # 1-based inclusive
                raw_block      = $rawText
            }
            if ($fmtCount -gt 0) { $ctrlEntry['format_conditions'] = $fmtCount }
            $controls.Add([PSCustomObject]$ctrlEntry)

            # Container types: re-scan inside instead of skipping past
            if ($containerTypes.Contains($mCtrl.Groups[1].Value)) {
                $containerStack.Add([PSCustomObject]@{ name = $name; end_idx = $ctrlEnd })
                $i = $ctrlStart + 1  # re-scan inside the container
            } else {
                $i = $ctrlEnd + 1
            }
            continue
        }
        $i++
    }

    $result['controls'] = @($controls)
    return $result
}

# ── 4.5 Get-ParsedControls (internal, cached) ───────────────────────────

function Get-ParsedControls {
    <#
    .SYNOPSIS
        Return parsed controls for a form/report, using the ControlsCache (internal helper).
    #>
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName
    )
    $cacheKey = "${ObjectType}:${ObjectName}"
    if (-not $script:AccessSession.ControlsCache.ContainsKey($cacheKey)) {
        $text = Get-AccessCode -DbPath $DbPath -ObjectType $ObjectType -Name $ObjectName
        $script:AccessSession.ControlsCache[$cacheKey] = ConvertFrom-ControlBlock -FormText $text
    }
    return $script:AccessSession.ControlsCache[$cacheKey]
}

# ── 4.6 New-AccessForm ───────────────────────────────────────────────────

function New-AccessForm {
    <#
    .SYNOPSIS
        Create a new blank form.
    .DESCRIPTION
        Uses CreateForm() which auto-names the form (Form1, Form2...).
        Saves, closes, then renames to the desired name.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER FormName
        Desired name for the new form.
    .PARAMETER HasHeader
        Toggle form header/footer sections.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-AccessForm -DbPath "C:\db.accdb" -FormName "frmCustomers"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$FormName,
        [switch]$HasHeader,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $autoName = $null
    try {
        $form = $app.CreateForm()
        $autoName = $form.Name

        if ($HasHeader) {
            try {
                $app.Visible = $true
                $app.RunCommand(36)  # acCmdFormHdrFtr — toggle header/footer
            } catch {
                Write-Warning "Could not toggle header/footer via RunCommand: $_"
            }
        }

        # Save with auto-name (no dialog)
        $app.DoCmd.Save($script:AC_TYPE['form'], $autoName)
        # Close without prompt (already saved)
        $app.DoCmd.Close($script:AC_TYPE['form'], $autoName, 2)  # acSaveNo=2

        # Rename to desired name
        if ($autoName -ne $FormName) {
            $app.DoCmd.Rename($FormName, $script:AC_TYPE['form'], $autoName)
        }

        return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            name         = $FormName
            created_from = $autoName
            has_header   = [bool]$HasHeader
        })
    } catch {
        if ($autoName) {
            try { $app.DoCmd.Close($script:AC_TYPE['form'], $autoName, 2) } catch {}
            try { $app.DoCmd.DeleteObject($script:AC_TYPE['form'], $autoName) } catch {}
        }
        throw "Error creating form '$FormName': $_"
    } finally {
        $script:AccessSession.VbeCodeCache = @{}
        $script:AccessSession.ControlsCache = @{}
        $script:AccessSession.CmCache = @{}
    }
}

# ── 4.7 Get-AccessFormProperty ───────────────────────────────────────────

function Get-AccessFormProperty {
    <#
    .SYNOPSIS
        Read properties from a form or report by opening it in Design view.
    .DESCRIPTION
        If PropertyNames is omitted, reads all readable properties.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER PropertyNames
        Array of property names to read. If omitted, reads all.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessFormProperty -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -PropertyNames "Caption","RecordSource"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [string[]]$PropertyNames,
        [switch]$AsJson
    )

    $null = Connect-AccessDB -DbPath $DbPath
    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    $properties = [ordered]@{}
    $errors = [ordered]@{}
    try {
        # Access properties directly via Screen.ActiveForm/ActiveReport
        # (COM collections/objects lose their RCW when passed through PS function returns or if-expressions)
        if ($PropertyNames) {
            foreach ($pname in $PropertyNames) {
                try {
                    if ($ObjectType -eq 'form') {
                        $val = $script:AccessSession.App.Screen.ActiveForm.Properties.Item($pname).Value
                    } else {
                        $val = $script:AccessSession.App.Screen.ActiveReport.Properties.Item($pname).Value
                    }
                    $properties[$pname] = ConvertTo-SafeValue -Value $val
                } catch {
                    $errors[$pname] = "$_"
                }
            }
        } else {
            if ($ObjectType -eq 'form') {
                $cnt = $script:AccessSession.App.Screen.ActiveForm.Properties.Count
            } else {
                $cnt = $script:AccessSession.App.Screen.ActiveReport.Properties.Count
            }
            for ($i = 0; $i -lt $cnt; $i++) {
                try {
                    if ($ObjectType -eq 'form') {
                        $pName  = $script:AccessSession.App.Screen.ActiveForm.Properties.Item($i).Name
                        $pValue = $script:AccessSession.App.Screen.ActiveForm.Properties.Item($i).Value
                    } else {
                        $pName  = $script:AccessSession.App.Screen.ActiveReport.Properties.Item($i).Name
                        $pValue = $script:AccessSession.App.Screen.ActiveReport.Properties.Item($i).Value
                    }
                    $properties[$pName] = ConvertTo-SafeValue -Value $pValue
                } catch { }
            }
        }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    $result = [ordered]@{
        object     = $ObjectName
        type       = $ObjectType
        properties = $properties
    }
    if ($errors.Count -gt 0) { $result['errors'] = $errors }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 4.8 Set-AccessFormProperty ───────────────────────────────────────────

function Set-AccessFormProperty {
    <#
    .SYNOPSIS
        Set properties on a form or report by opening it in Design view.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER Properties
        Hashtable of property name/value pairs to set.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessFormProperty -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -Properties @{ Caption = "My Form"; RecordSource = "tblCustomers" }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][hashtable]$Properties,
        [switch]$AsJson
    )

    $null = Connect-AccessDB -DbPath $DbPath
    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    $applied = [System.Collections.Generic.List[string]]::new()
    $errors = [ordered]@{}
    try {
        foreach ($key in $Properties.Keys) {
            try {
                $coerced = ConvertTo-CoercedProp -Value $Properties[$key]
                if ($ObjectType -eq 'form') {
                    $script:AccessSession.App.Screen.ActiveForm.$key = $coerced
                } else {
                    $script:AccessSession.App.Screen.ActiveReport.$key = $coerced
                }
                $applied.Add($key)
            } catch {
                $errors[$key] = "$_"
            }
        }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    $result = [ordered]@{
        applied = @($applied)
        errors  = $errors
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 4.9 Get-AccessControl ────────────────────────────────────────────────

function Get-AccessControl {
    <#
    .SYNOPSIS
        List controls in a form or report (from cached parsed export).
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessControl -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [switch]$AsJson
    )

    $parsed = Get-ParsedControls -DbPath $DbPath -ObjectType $ObjectType -ObjectName $ObjectName
    $controls = @(
        $parsed.controls | Where-Object { $_.name.Trim() } | ForEach-Object {
            $c = [ordered]@{}
            foreach ($prop in $_.PSObject.Properties) {
                if ($prop.Name -ne 'raw_block') {
                    $c[$prop.Name] = $prop.Value
                }
            }
            [PSCustomObject]$c
        }
    )

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        count    = $controls.Count
        controls = $controls
    })
}

# ── 4.10 Get-AccessControlDetail ─────────────────────────────────────────

function Get-AccessControlDetail {
    <#
    .SYNOPSIS
        Get the full definition of a single control by name (includes raw_block).
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER ControlName
        Name of the control.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessControlDetail -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -ControlName "txtLastName"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][string]$ControlName,
        [switch]$AsJson
    )

    $parsed = Get-ParsedControls -DbPath $DbPath -ObjectType $ObjectType -ObjectName $ObjectName
    $found = $parsed.controls | Where-Object { $_.name -ieq $ControlName } | Select-Object -First 1
    if (-not $found) {
        $names = @($parsed.controls | ForEach-Object { $_.name })
        throw "Control '$ControlName' not found in '$ObjectName'. Available controls: $($names -join ', ')"
    }

    Format-AccessOutput -AsJson:$AsJson -Data $found
}

# ── 4.11 New-AccessControl ───────────────────────────────────────────────

function New-AccessControl {
    <#
    .SYNOPSIS
        Create a new control on a form or report.
    .DESCRIPTION
        Opens the form/report in Design view, calls CreateControl/CreateReportControl,
        sets properties, saves and closes.
        Structural properties (section, parent, column_name, left, top, width, height) are
        passed to CreateControl. All other properties are set via COM after creation.
        For ActiveX controls (type 119), pass ClassName for the ProgID.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER ControlType
        Control type: name ('CommandButton') or number (104).
    .PARAMETER Properties
        Hashtable of properties. Special keys: section, parent, column_name, left, top, width, height.
    .PARAMETER ClassName
        ProgID for ActiveX controls (type 119, e.g., 'Shell.Explorer.2').
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-AccessControl -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -ControlType "CommandButton" -Properties @{ Name = "btnSave"; Caption = "Save" }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)]$ControlType,
        [hashtable]$Properties = @{},
        [string]$ClassName,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath

    # Resolve control type
    $ctype = $ControlType
    if ($ctype -is [string]) {
        $key = $ctype.ToLower()
        if ($script:CTRL_TYPE_BY_NAME.ContainsKey($key)) {
            $ctype = $script:CTRL_TYPE_BY_NAME[$key]
        } else {
            $intVal = 0
            if ([int]::TryParse($ctype, [ref]$intVal)) { $ctype = $intVal }
            else { throw "Unknown control type: '$ControlType'" }
        }
    }
    $ctype = [int]$ctype

    # Extract structural params from Properties (don't set as COM properties later)
    $p = @{} + $Properties  # copy
    $section    = 0
    $parent     = ''
    $columnName = ''
    $left       = -1
    $top        = -1
    $width      = -1
    $height     = -1

    if ($p.ContainsKey('section')) {
        $secVal = "$($p['section'])".ToLower()
        if ($script:SECTION_MAP.ContainsKey($secVal)) { $section = $script:SECTION_MAP[$secVal] }
        else { $section = [int](ConvertTo-CoercedProp -Value $p['section']) }
        $p.Remove('section')
    }
    if ($p.ContainsKey('parent'))      { $parent     = "$($p['parent'])";      $p.Remove('parent') }
    if ($p.ContainsKey('column_name')) { $columnName = "$($p['column_name'])"; $p.Remove('column_name') }
    if ($p.ContainsKey('left'))   { $left   = [int](ConvertTo-CoercedProp -Value $p['left']);   $p.Remove('left') }
    if ($p.ContainsKey('top'))    { $top    = [int](ConvertTo-CoercedProp -Value $p['top']);    $p.Remove('top') }
    if ($p.ContainsKey('width'))  { $width  = [int](ConvertTo-CoercedProp -Value $p['width']);  $p.Remove('width') }
    if ($p.ContainsKey('height')) { $height = [int](ConvertTo-CoercedProp -Value $p['height']); $p.Remove('height') }

    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    try {
        try {
            if ($ObjectType -eq 'form') {
                $ctrl = $app.CreateControl($ObjectName, $ctype, $section, $parent, $columnName, $left, $top, $width, $height)
            } else {
                $ctrl = $app.CreateReportControl($ObjectName, $ctype, $section, $parent, $columnName, $left, $top, $width, $height)
            }
        } catch {
            $secNames = @($script:SECTION_MAP.GetEnumerator() | Where-Object { $_.Value -eq $section } | ForEach-Object { $_.Key })
            throw ("Error creating control in section=$section ($($secNames -join ', ')): $_. " +
                   "Valid sections: 0=Detail, 1=Header, 2=Footer, 3=PageHeader, 4=PageFooter")
        }

        # ActiveX: set ProgID via Class property
        if ($ClassName -and $ctype -eq 119) {
            try { $ctrl.Class = $ClassName } catch { Write-Warning "Could not set Class='$ClassName': $_" }
        }

        $propErrors = [ordered]@{}
        foreach ($key in $p.Keys) {
            try {
                $ctrl.$key = ConvertTo-CoercedProp -Value $p[$key]
            } catch {
                $propErrors[$key] = "$_"
            }
        }

        $resultData = [ordered]@{
            name         = $ctrl.Name
            control_type = $ctype
            type_name    = if ($script:CTRL_TYPE.ContainsKey($ctype)) { $script:CTRL_TYPE[$ctype] } else { "Type$ctype" }
        }
        if ($propErrors.Count -gt 0) { $resultData['property_errors'] = $propErrors }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    Format-AccessOutput -AsJson:$AsJson -Data $resultData
}

# ── 4.12 Remove-AccessControl ────────────────────────────────────────────

function Remove-AccessControl {
    <#
    .SYNOPSIS
        Delete a control from a form or report.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER ControlName
        Name of the control to delete.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-AccessControl -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -ControlName "txtOld"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][string]$ControlName,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    try {
        if ($ObjectType -eq 'form') {
            $app.DeleteControl($ObjectName, $ControlName)
        } else {
            $app.DeleteReportControl($ObjectName, $ControlName)
        }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status  = "OK: control '$ControlName' deleted from '$ObjectName'"
        control = $ControlName
        object  = $ObjectName
    })
}

# ── 4.13 Set-AccessControlProperty ───────────────────────────────────────

function Set-AccessControlProperty {
    <#
    .SYNOPSIS
        Modify properties of an existing control on a form or report.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER ControlName
        Name of the control to modify.
    .PARAMETER Properties
        Hashtable of property name/value pairs to set.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessControlProperty -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -ControlName "txtName" -Properties @{ Caption = "Full Name"; Width = 3000 }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][string]$ControlName,
        [Parameter(Mandatory)][hashtable]$Properties,
        [switch]$AsJson
    )

    $null = Connect-AccessDB -DbPath $DbPath
    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    $applied = [System.Collections.Generic.List[string]]::new()
    $errors = [ordered]@{}
    try {
        foreach ($key in $Properties.Keys) {
            try {
                $coerced = ConvertTo-CoercedProp -Value $Properties[$key]
                if ($ObjectType -eq 'form') {
                    $script:AccessSession.App.Screen.ActiveForm.Controls.Item($ControlName).$key = $coerced
                } else {
                    $script:AccessSession.App.Screen.ActiveReport.Controls.Item($ControlName).$key = $coerced
                }
                $applied.Add($key)
            } catch {
                $errors[$key] = "$_"
            }
        }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        applied = @($applied)
        errors  = $errors
    })
}

# ── 4.14 Set-AccessControlBatch ──────────────────────────────────────────

function Set-AccessControlBatch {
    <#
    .SYNOPSIS
        Modify properties of multiple controls in a single Design view session.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER Controls
        Array of hashtables, each with 'name' (string) and 'props' (hashtable).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessControlBatch -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -Controls @(
            @{ name = "txtFirst"; props = @{ Width = 3000 } },
            @{ name = "txtLast";  props = @{ Width = 3000 } }
        )
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][array]$Controls,
        [switch]$AsJson
    )

    if ($Controls.Count -eq 0) {
        return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{ error = 'No controls provided.' })
    }

    $null = Connect-AccessDB -DbPath $DbPath
    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    $results = [System.Collections.Generic.List[object]]::new()
    try {
        foreach ($ctrlSpec in $Controls) {
            $ctrlName  = $ctrlSpec['name']
            $ctrlProps = $ctrlSpec['props']
            if (-not $ctrlProps) { $ctrlProps = @{} }
            $applied = [System.Collections.Generic.List[string]]::new()
            $errors  = [ordered]@{}
            try {
                foreach ($key in $ctrlProps.Keys) {
                    try {
                        $coerced = ConvertTo-CoercedProp -Value $ctrlProps[$key]
                        if ($ObjectType -eq 'form') {
                            $script:AccessSession.App.Screen.ActiveForm.Controls.Item($ctrlName).$key = $coerced
                        } else {
                            $script:AccessSession.App.Screen.ActiveReport.Controls.Item($ctrlName).$key = $coerced
                        }
                        $applied.Add($key)
                    } catch {
                        $errors[$key] = "$_"
                    }
                }
            } catch {
                $errors['_control'] = "Control '$ctrlName' not found: $_"
            }
            $entry = [ordered]@{ name = $ctrlName; applied = @($applied) }
            if ($errors.Count -gt 0) { $entry['errors'] = $errors }
            $results.Add([PSCustomObject]$entry)
        }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{ results = @($results) })
}

# ═══════════════════════════════════════════════════════════════════════════
# PHASE 5 — Structure & Metadata Operations
# ═══════════════════════════════════════════════════════════════════════════

# ── 5.1 Edit-AccessTable ─────────────────────────────────────────────────

function Edit-AccessTable {
    <#
    .SYNOPSIS
        Add, delete, or rename fields in an existing table via DAO.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][ValidateSet('add_field','delete_field','rename_field')][string]$Action,
        [Parameter(Mandatory)][string]$FieldName,
        [string]$NewName,
        [string]$FieldType = 'text',
        [int]$Size = 0,
        [switch]$Required,
        $Default,
        [string]$Description,
        [switch]$ConfirmDelete,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $td  = $db.TableDefs($TableName)

    switch ($Action) {
        'add_field' {
            $ftype   = $FieldType.ToLower()
            $daoType = $script:FIELD_TYPE_MAP[$ftype]
            if ($null -eq $daoType) {
                $validTypes = ($script:FIELD_TYPE_MAP.Keys | Sort-Object -Unique) -join ', '
                throw "Unknown type: '$ftype'. Valid: $validTypes"
            }
            $isAuto = $ftype -in 'autonumber', 'autoincrement'

            if ($daoType -eq 10 -and $Size -eq 0) { $Size = 255 }

            $fld = if ($Size -gt 0) {
                $td.CreateField($FieldName, $daoType, $Size)
            } else {
                $td.CreateField($FieldName, $daoType)
            }

            if ($isAuto) {
                $fld.Attributes = $fld.Attributes -bor $script:DB_AUTO_INCR_FIELD
            }
            $fld.Required = [bool]$Required

            $td.Fields.Append($fld)
            $td.Fields.Refresh()

            if ($null -ne $Default) {
                try { Set-FieldProperty -Db $db -TableName $TableName -FieldName $FieldName -PropertyName 'DefaultValue' -Value ([string]$Default) } catch {}
            }
            if (-not [string]::IsNullOrEmpty($Description)) {
                try { Set-FieldProperty -Db $db -TableName $TableName -FieldName $FieldName -PropertyName 'Description' -Value $Description } catch {}
            }

            $result = [ordered]@{ action = 'field_added'; table = $TableName; field = $FieldName; type = $ftype }
        }
        'delete_field' {
            if (-not $ConfirmDelete) {
                $result = [ordered]@{ error = "Deleting field '$FieldName' from '$TableName' is destructive. Use -ConfirmDelete to confirm." }
            } else {
                $td.Fields.Delete($FieldName)
                $result = [ordered]@{ action = 'field_deleted'; table = $TableName; field = $FieldName }
            }
        }
        'rename_field' {
            if ([string]::IsNullOrEmpty($NewName)) {
                throw "rename_field requires -NewName"
            }
            $fld = $td.Fields($FieldName)
            $fld.Name = $NewName
            $result = [ordered]@{ action = 'field_renamed'; table = $TableName; old_name = $FieldName; new_name = $NewName }
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 5.2 Get-AccessFieldProperty ─────────────────────────────────────────

function Get-AccessFieldProperty {
    <#
    .SYNOPSIS
        Read all DAO properties from a table field.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][string]$FieldName,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $fld = $db.TableDefs($TableName).Fields($FieldName)

    $props = [ordered]@{}
    for ($i = 0; $i -lt $fld.Properties.Count; $i++) {
        try {
            $p   = $fld.Properties($i)
            $val = $p.Value
            if ($val -is [string] -or $val -is [int] -or $val -is [long] -or
                $val -is [double] -or $val -is [float] -or $val -is [bool] -or $null -eq $val) {
                $props[$p.Name] = $val
            }
        } catch {
            # Skip unreadable properties
        }
    }

    $result = [ordered]@{
        table_name  = $TableName
        field_name  = $FieldName
        properties  = $props
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 5.3 Set-AccessFieldProperty ─────────────────────────────────────────

function Set-AccessFieldProperty {
    <#
    .SYNOPSIS
        Set or create a DAO property on a table field.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][string]$FieldName,
        [Parameter(Mandatory)][string]$PropertyName,
        [Parameter(Mandatory)]$Value,
        [switch]$AsJson
    )

    $app     = Connect-AccessDB -DbPath $DbPath
    $db      = $app.CurrentDb()
    $fld     = $db.TableDefs($TableName).Fields($FieldName)
    $coerced = ConvertTo-CoercedProp $Value

    # Try updating existing property first
    $actionTaken = 'updated'
    try {
        $fld.Properties($PropertyName).Value = $coerced
    } catch {
        # Property doesn't exist — create it
        $propType = if ($coerced -is [bool]) { 1 }        # dbBoolean
                    elseif ($coerced -is [int])  { 4 }     # dbLong
                    else                         { 10 }    # dbText

        $prop = $fld.CreateProperty($PropertyName, $propType, $coerced)
        $fld.Properties.Append($prop)
        $actionTaken = 'created'
    }

    $result = [ordered]@{
        table_name    = $TableName
        field_name    = $FieldName
        property_name = $PropertyName
        value         = $coerced
        action        = $actionTaken
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 5.4 Get-AccessLinkedTable ────────────────────────────────────────────

function Get-AccessLinkedTable {
    <#
    .SYNOPSIS
        List all linked (attached) tables in the database.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    $linked = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $db.TableDefs.Count; $i++) {
        $td   = $db.TableDefs($i)
        $conn = $td.Connect
        if ([string]::IsNullOrEmpty($conn)) { continue }

        $name = $td.Name
        if ($name.StartsWith('~') -or $name.StartsWith('MSys')) { continue }

        $linked.Add([PSCustomObject][ordered]@{
            name           = $name
            source_table   = $td.SourceTableName
            connect_string = $conn
            is_odbc        = $conn.ToUpper().StartsWith('ODBC;')
        })
    }

    $result = [ordered]@{
        count         = $linked.Count
        linked_tables = @($linked)
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 5.5 Set-AccessLinkedTable ────────────────────────────────────────────

function Set-AccessLinkedTable {
    <#
    .SYNOPSIS
        Relink a linked table (or all tables sharing the same connection) to a new data source.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][string]$NewConnect,
        [switch]$RelinkAll,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    # Verify the reference table is actually linked
    $refTd = $db.TableDefs($TableName)
    if ([string]::IsNullOrEmpty($refTd.Connect)) {
        throw "'$TableName' is not a linked table."
    }

    $relinked = [System.Collections.Generic.List[object]]::new()

    $relinkOne = {
        param([string]$tName, [string]$oldConn)
        $t = $db.TableDefs($tName)
        $t.Connect = $NewConnect
        $t.RefreshLink()
        $relinked.Add([PSCustomObject][ordered]@{
            name        = $tName
            old_connect = $oldConn
            new_connect = $NewConnect
        })
    }

    if ($RelinkAll) {
        $oldConnect = $refTd.Connect
        $namesToRelink = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $db.TableDefs.Count; $i++) {
            $td = $db.TableDefs($i)
            if ($td.Connect -eq $oldConnect) {
                $namesToRelink.Add([PSCustomObject]@{ Name = $td.Name; Connect = $td.Connect })
            }
        }
        foreach ($item in $namesToRelink) {
            & $relinkOne $item.Name $item.Connect
        }
    } else {
        & $relinkOne $TableName $refTd.Connect
    }

    $result = [ordered]@{
        relinked_count = $relinked.Count
        tables         = @($relinked)
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 5.6 Get-AccessRelationship ───────────────────────────────────────────

function Get-AccessRelationship {
    <#
    .SYNOPSIS
        List all non-system relationships in the database.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $rels = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $db.Relations.Count; $i++) {
        $rel  = $db.Relations($i)
        $name = $rel.Name
        if ($name.StartsWith('MSys')) { continue }

        $fields = [System.Collections.Generic.List[object]]::new()
        for ($j = 0; $j -lt $rel.Fields.Count; $j++) {
            $fld = $rel.Fields($j)
            $fields.Add([ordered]@{ local = $fld.Name; foreign = $fld.ForeignName })
        }

        $attrs     = [int]$rel.Attributes
        $attrFlags = foreach ($bit in $script:REL_ATTR.Keys) {
            if ($attrs -band $bit) { $script:REL_ATTR[$bit] }
        }
        if ($null -eq $attrFlags) { $attrFlags = @() }

        $rels.Add([PSCustomObject][ordered]@{
            name            = $name
            table           = $rel.Table
            foreign_table   = $rel.ForeignTable
            fields          = @($fields)
            attributes      = $attrs
            attribute_flags = @($attrFlags)
        })
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        count         = $rels.Count
        relationships = @($rels)
    })
}

# ── 5.7 New-AccessRelationship ───────────────────────────────────────────

function New-AccessRelationship {
    <#
    .SYNOPSIS
        Create a new relationship between two tables.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Table,
        [Parameter(Mandatory)][string]$ForeignTable,
        [Parameter(Mandatory)][array]$Fields,
        [int]$Attributes = 0,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    $rel = $db.CreateRelation($Name, $Table, $ForeignTable, $Attributes)
    foreach ($fmap in $Fields) {
        $localName   = $fmap['local']
        $foreignName = $fmap['foreign']
        if (-not $localName -or -not $foreignName) {
            throw "Each field mapping must have 'local' and 'foreign' keys."
        }
        $fld = $rel.CreateField($localName)
        $fld.ForeignName = $foreignName
        $rel.Fields.Append($fld)
    }
    $db.Relations.Append($rel)

    $attrFlags = foreach ($bit in $script:REL_ATTR.Keys) {
        if ($Attributes -band $bit) { $script:REL_ATTR[$bit] }
    }
    if ($null -eq $attrFlags) { $attrFlags = @() }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        name            = $Name
        table           = $Table
        foreign_table   = $ForeignTable
        fields          = @($Fields)
        attributes      = $Attributes
        attribute_flags = @($attrFlags)
        status          = 'created'
    })
}

# ── 5.8 Remove-AccessRelationship ────────────────────────────────────────

function Remove-AccessRelationship {
    <#
    .SYNOPSIS
        Delete a relationship by name.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$Name,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $db.Relations.Delete($Name)

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        action = 'deleted'
        name   = $Name
    })
}

# ── 5.9 Get-AccessReference ─────────────────────────────────────────────

function Get-AccessReference {
    <#
    .SYNOPSIS
        List all VBA project references in the database.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [switch]$AsJson
    )

    $app     = Connect-AccessDB -DbPath $DbPath
    $refsCol = $app.VBE.ActiveVBProject.References
    $refs    = [System.Collections.Generic.List[object]]::new()

    for ($i = 1; $i -le $refsCol.Count; $i++) {
        $ref = $refsCol.Item($i)

        $isBroken = $true
        try { $isBroken = [bool]$ref.IsBroken } catch {}

        $builtIn = $false
        try { $builtIn = [bool]$ref.BuiltIn } catch {}

        $guid = ''
        try { if ($ref.GUID) { $guid = $ref.GUID } } catch {}

        $refs.Add([PSCustomObject][ordered]@{
            name        = $ref.Name
            description = $ref.Description
            full_path   = $ref.FullPath
            guid        = $guid
            major       = [int]$ref.Major
            minor       = [int]$ref.Minor
            is_broken   = $isBroken
            built_in    = $builtIn
        })
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        count      = $refs.Count
        references = @($refs)
    })
}

# ── 5.10 Set-AccessReference ────────────────────────────────────────────

function Set-AccessReference {
    <#
    .SYNOPSIS
        Add or remove a VBA project reference.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('add','remove')][string]$Action,
        [string]$Name,
        [string]$RefPath,
        [string]$Guid,
        [int]$Major = 0,
        [int]$Minor = 0,
        [switch]$AsJson
    )

    $app  = Connect-AccessDB -DbPath $DbPath
    $refs = $app.VBE.ActiveVBProject.References

    if ($Action -eq 'add') {
        if ($Guid) {
            $ref    = $refs.AddFromGuid($Guid, $Major, $Minor)
            $result = [ordered]@{
                action = 'added'; name = $ref.Name; guid = $Guid; major = $Major; minor = $Minor
            }
        } elseif ($RefPath) {
            $ref    = $refs.AddFromFile($RefPath)
            $result = [ordered]@{
                action = 'added'; name = $ref.Name; full_path = $RefPath
            }
        } else {
            throw "Action 'add' requires either -Guid or -RefPath."
        }
    } else {
        # remove
        if (-not $Name) { throw "Action 'remove' requires -Name." }
        $found = $null
        for ($i = 1; $i -le $refs.Count; $i++) {
            $ref = $refs.Item($i)
            if ($ref.Name -ieq $Name) { $found = $ref; break }
        }
        if ($null -eq $found) { throw "Reference '$Name' not found." }
        try { if ($found.BuiltIn) { throw "Cannot remove built-in reference '$Name'." } } catch [System.Management.Automation.PropertyNotFoundException] {}
        $refs.Remove($found)
        $result = [ordered]@{ action = 'removed'; name = $Name }
    }

    # Clear VBE caches
    $script:AccessSession.VbeCodeCache = @{}
    $script:AccessSession.CmCache     = @{}

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 5.11 Set-AccessQuery ─────────────────────────────────────────────────

function Set-AccessQuery {
    <#
    .SYNOPSIS
        Create, modify, delete, rename, or retrieve SQL for an Access query.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('create','modify','delete','rename','get_sql')][string]$Action,
        [Parameter(Mandatory)][string]$QueryName,
        [string]$Sql,
        [string]$NewName,
        [switch]$ConfirmDelete,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    switch ($Action) {
        'create' {
            if (-not $Sql) { throw "create requires -Sql" }
            $null = $db.CreateQueryDef($QueryName, $Sql)
            $result = [ordered]@{ action = 'created'; query_name = $QueryName; sql = $Sql }
        }
        'modify' {
            if (-not $Sql) { throw "modify requires -Sql" }
            $qd = $db.QueryDefs($QueryName)
            $qd.SQL = $Sql
            $result = [ordered]@{ action = 'modified'; query_name = $QueryName; sql = $Sql }
        }
        'delete' {
            if (-not $ConfirmDelete) {
                $result = [ordered]@{ error = "Deleting query '$QueryName' requires -ConfirmDelete" }
            } else {
                $null = $db.QueryDefs($QueryName)   # verify exists
                $db.QueryDefs.Delete($QueryName)
                $result = [ordered]@{ action = 'deleted'; query_name = $QueryName }
            }
        }
        'rename' {
            if (-not $NewName) { throw "rename requires -NewName" }
            $qd = $db.QueryDefs($QueryName)
            $qd.Name = $NewName
            $result = [ordered]@{ action = 'renamed'; old_name = $QueryName; new_name = $NewName }
        }
        'get_sql' {
            $qd = $db.QueryDefs($QueryName)
            $qdType = $script:QUERYDEF_TYPE[[int]$qd.Type]
            if (-not $qdType) { $qdType = "Unknown($($qd.Type))" }
            $result = [ordered]@{ query_name = $QueryName; sql = $qd.SQL; type = $qdType }
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 5.12 Get-AccessIndex ─────────────────────────────────────────────────

function Get-AccessIndex {
    <#
    .SYNOPSIS
        List all indexes on an Access table with field details.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$TableName,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $td  = $db.TableDefs($TableName)

    $indexes = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $td.Indexes.Count; $i++) {
        $idx = $td.Indexes($i)
        $fields = [System.Collections.Generic.List[object]]::new()

        for ($j = 0; $j -lt $idx.Fields.Count; $j++) {
            $f = $idx.Fields($j)
            $fields.Add([ordered]@{
                name  = $f.Name
                order = if ($f.Attributes -band 1) { 'desc' } else { 'asc' }
            })
        }

        $indexes.Add([ordered]@{
            name    = $idx.Name
            fields  = @($fields)
            primary = [bool]$idx.Primary
            unique  = [bool]$idx.Unique
            foreign = [bool]$idx.Foreign
        })
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        table_name = $TableName
        count      = $indexes.Count
        indexes    = @($indexes)
    })
}

# ── 5.13 Set-AccessIndex ─────────────────────────────────────────────────

function Set-AccessIndex {
    <#
    .SYNOPSIS
        Create or delete an index on an Access table.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][ValidateSet('create','delete')][string]$Action,
        [Parameter(Mandatory)][string]$IndexName,
        [array]$Fields,
        [switch]$Primary,
        [switch]$Unique,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $td  = $db.TableDefs($TableName)

    switch ($Action) {
        'create' {
            if (-not $Fields -or $Fields.Count -eq 0) { throw "create requires -Fields" }
            $idx = $td.CreateIndex($IndexName)
            $idx.Primary = [bool]$Primary
            $idx.Unique  = [bool]$Unique

            foreach ($fdef in $Fields) {
                if ($fdef -is [string]) {
                    $fname = $fdef
                    $fld   = $idx.CreateField($fname)
                } else {
                    $fname = $fdef['name']
                    $fld   = $idx.CreateField($fname)
                    if ($fdef.ContainsKey('order') -and $fdef['order'] -eq 'desc') {
                        $fld.Attributes = 1   # dbDescending
                    }
                }
                $idx.Fields.Append($fld)
            }

            $td.Indexes.Append($idx)
            $result = [ordered]@{
                action     = 'created'
                table_name = $TableName
                index_name = $IndexName
                fields     = $Fields
                primary    = [bool]$Primary
                unique     = [bool]$Unique
            }
        }
        'delete' {
            $null = $td.Indexes($IndexName)   # verify exists
            $td.Indexes.Delete($IndexName)
            $result = [ordered]@{
                action     = 'deleted'
                table_name = $TableName
                index_name = $IndexName
            }
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 5.14 Get-AccessStartupOption ─────────────────────────────────────────

function Get-AccessStartupOption {
    <#
    .SYNOPSIS
        List Access startup/application options from database properties and application settings.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    $options = [System.Collections.Generic.List[object]]::new()

    foreach ($name in $script:STARTUP_PROPS) {
        $val    = $null
        $source = '<not set>'

        try {
            $val    = $db.Properties($name).Value
            $source = 'database'
        } catch {
            try {
                $val    = $app.GetOption($name)
                $source = 'application'
            } catch {}
        }

        $options.Add([ordered]@{
            name   = $name
            value  = $val
            source = $source
        })
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        count   = $options.Count
        options = @($options)
    })
}

# ═══════════════════════════════════════════════════════════════════════════
# PHASE 6 — Export, Transfer & Properties
# ═══════════════════════════════════════════════════════════════════════════

# ── 6.1 Export-AccessReport ────────────────────────────────────────────────

function Export-AccessReport {
    <#
    .SYNOPSIS
        Export a report, table, query, or form to PDF, XLSX, RTF, or TXT via DoCmd.OutputTo.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$ObjectName,
        [ValidateSet('report','table','query','form')]
        [string]$ObjectType = 'report',
        [ValidateSet('pdf','xlsx','rtf','txt')]
        [string]$OutputFormat = 'pdf',
        [string]$OutputPath,
        [switch]$OpenAfterExport,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath

    $objectTypeMap = @{ report = 3; table = 0; query = 1; form = 2 }
    $acType = $objectTypeMap[$ObjectType]

    $formatString = $script:OUTPUT_FORMATS[$OutputFormat]
    if (-not $formatString) {
        throw "Unsupported output format '$OutputFormat'. Valid: $($script:OUTPUT_FORMATS.Keys -join ', ')"
    }

    if (-not $OutputPath) {
        $extMap = @{ pdf = '.pdf'; xlsx = '.xlsx'; rtf = '.rtf'; txt = '.txt' }
        $safeName = $ObjectName -replace '[\\/:*?"<>|]', '_'
        $OutputPath = Join-Path $env:TEMP "$safeName$($extMap[$OutputFormat])"
    }
    $OutputPath = [System.IO.Path]::GetFullPath($OutputPath)

    try {
        $app.DoCmd.OutputTo($acType, $ObjectName, $formatString, $OutputPath, [bool]$OpenAfterExport)
    } catch {
        throw "Error exporting '$ObjectName' ($ObjectType) to $OutputFormat : $_"
    }

    $fileSize = 0
    if (Test-Path -LiteralPath $OutputPath) {
        $fileSize = (Get-Item -LiteralPath $OutputPath).Length
    }

    $result = [ordered]@{
        object_name   = $ObjectName
        object_type   = $ObjectType
        output_format = $OutputFormat
        output_path   = $OutputPath
        file_size     = $fileSize
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 6.2 Copy-AccessData ──────────────────────────────────────────────────

function Copy-AccessData {
    <#
    .SYNOPSIS
        Import or export data via DoCmd.TransferSpreadsheet / DoCmd.TransferText.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('import','export')][string]$Action,
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$TableName,
        [bool]$HasHeaders = $true,
        [ValidateSet('xlsx','xls','excel','csv','txt','text')]
        [string]$FileType = 'xlsx',
        [string]$Range,
        [string]$SpecName,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $FilePath = [System.IO.Path]::GetFullPath($FilePath)
    $ft = $FileType.ToLower()

    if ($Action -eq 'import') {
        $transferTypeSpreadsheet = $script:AC_IMPORT       # 0
        $transferTypeText        = $script:AC_IMPORT       # 0
    } else {
        $transferTypeSpreadsheet = $script:AC_EXPORT       # 1
        $transferTypeText        = $script:AC_EXPORT_DELIM # 2
    }

    try {
        if ($ft -in @('xlsx', 'xls', 'excel')) {
            $rangeArg = if ($Range) { $Range } else { '' }
            $app.DoCmd.TransferSpreadsheet(
                $transferTypeSpreadsheet,
                $script:AC_SPREADSHEET_XLSX,
                $TableName,
                $FilePath,
                $HasHeaders,
                $rangeArg
            )
        } elseif ($ft -in @('csv', 'txt', 'text')) {
            $specArg = if ($SpecName) { $SpecName } else { '' }
            $app.DoCmd.TransferText(
                $transferTypeText,
                $specArg,
                $TableName,
                $FilePath,
                $HasHeaders
            )
        }
    } catch {
        throw "Error transferring data ($Action, $ft) for table '$TableName': $_"
    }

    $result = [ordered]@{
        action     = $Action
        file_type  = $ft
        table_name = $TableName
        file_path  = $FilePath
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 6.3 Get-AccessDatabaseProperty ──────────────────────────────────────

function Get-AccessDatabaseProperty {
    <#
    .SYNOPSIS
        Read a database property from CurrentDb().Properties or Application.GetOption.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$Name,
        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    try {
        $val = $db.Properties($Name).Value
        $result = [ordered]@{ name = $Name; value = $val; source = 'database' }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    } catch {}

    try {
        $val = $app.GetOption($Name)
        $result = [ordered]@{ name = $Name; value = $val; source = 'application' }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    } catch {
        throw "Property '$Name' not found in CurrentDb().Properties or Application.GetOption"
    }
}

# ── 6.4 Set-AccessDatabaseProperty ──────────────────────────────────────

function Set-AccessDatabaseProperty {
    <#
    .SYNOPSIS
        Set or create a database property in CurrentDb().Properties or Application.SetOption.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Value,
        [int]$PropType = -1,
        [switch]$AsJson
    )

    $app     = Connect-AccessDB -DbPath $DbPath
    $db      = $app.CurrentDb()
    $coerced = ConvertTo-CoercedProp -Value $Value

    try {
        $db.Properties($Name).Value = $coerced
        $result = [ordered]@{ name = $Name; value = $coerced; source = 'database'; action = 'updated' }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    } catch {}

    try {
        $app.SetOption($Name, $coerced)
        $result = [ordered]@{ name = $Name; value = $coerced; source = 'application'; action = 'updated' }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    } catch {}

    # Create new database property
    if ($PropType -eq -1) {
        if ($coerced -is [bool])                          { $PropType = 1 }
        elseif ($coerced -is [int] -or $coerced -is [long]) { $PropType = 4 }
        else                                               { $PropType = 10 }
    }

    $prop = $db.CreateProperty($Name, $PropType, $coerced)
    $db.Properties.Append($prop)

    $result = [ordered]@{ name = $Name; value = $coerced; source = 'database'; action = 'created' }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ── 6.5 Get-AccessTip ───────────────────────────────────────────────────

function Get-AccessTip {
    <#
    .SYNOPSIS
        Return tips and gotchas by topic for working with Access.
    #>
    [CmdletBinding()]
    param(
        [string]$Topic,
        [switch]$AsJson
    )

    $tips = [ordered]@{
        eval = @"
Invoke-AccessEval can query the Access Object Model without new tools:
  Application.IsCompiled - check if VBA is compiled
  SysCmd(10, 2, "formName") - check if form is open
  Application.BrokenReference - True if any ref is broken
  Screen.ActiveForm.Name / Screen.ActiveControl.Name - active form/control
  Forms.Count - number of open forms
  TempVars("x") - session-persistent variables
  DLookup/DCount/DSum - domain aggregate functions
  TypeName(expr) - inspect type
  Eval only works for expressions/functions, NOT statements/Subs.
"@
        controls = @"
Control types for New-AccessControl:
  119 = acCustomControl (ActiveX) - use ClassName for ProgID
  128 = acWebBrowser (native, NOT ActiveX)
  Common: 100=Label, 109=TextBox, 106=ComboBox, 105=ListBox, 104=CommandButton,
          110=CheckBox, 114=SubForm, 122=Image, 101=Rectangle

  FormatConditions: Get-AccessControl / Get-AccessControlDetail show
  format_conditions count. Use VBA via Invoke-AccessVba to read/modify details.
"@
        gotchas = @"
COM & ODBC:
  dbSeeChanges (512) - REQUIRED for DELETE/UPDATE on ODBC linked tables
  LIKE wildcards - use % for ODBC (not *)
  ListBox.Value - use .Column(0) explicitly
  dbAttachSavePWD = 131072 (NOT 65536)
  Multiple JOINs - Access requires nested parentheses

VBA:
  Str() adds leading space - use CStr()
  IIf() evaluates ALL three args (not short-circuit) - use If/Then/Else
  Dim X As New ClassName in a loop only creates ONE instance
  Chr(128) truncates MsgBox - use ChrW(8364) for euro
"@
        sql = @"
Jet SQL DDL:
  YESNO is not valid - use BIT
  DEFAULT not supported in CREATE TABLE - use Set-AccessFieldProperty
  AUTOINCREMENT works as a type
  Use SHORT instead of SMALLINT, LONG instead of INT
  Prefer New-AccessTable over CREATE TABLE SQL

ODBC pass-through:
  QueryDef.Connect limit 255 chars
"@
        vbe = @"
VBE line numbers are 1-based.
ProcCountLines can inflate the last proc count past end - always clamp.
Access must be Visible=True for VBE COM access.
'Trust access to the VBA project object model' must be enabled.
After design operations, close form before accessing VBE CodeModule.
"@
        compile = @"
Test-AccessVbaCompile tips:
  RunCommand(126) shows MsgBox on error - use timeout param.
  Before compiling: Eval('Application.BrokenReference') for broken refs.
  After error: use Get-AccessVbeLine to read problematic code.
"@
        design = @"
Design view + VBE conflict:
  After design ops, form may remain open in Design view.
  Set-AccessVbeProc closes the form (acSaveYes) before VBE access.
  All design operations invalidate caches.

SaveAsText encoding:
  Modules (.bas) - cp1252 (ANSI, no BOM)
  Forms/reports - utf-16 (UTF-16LE with BOM)
"@
    }

    if (-not $Topic -or $Topic.Trim() -eq '') {
        $result = [ordered]@{
            topics = @($tips.Keys)
            hint   = 'Pass -Topic <name> for details. Fuzzy matching supported.'
        }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    }

    $key = $Topic.Trim().ToLower()

    # Exact match
    if ($tips.Contains($key)) {
        $result = [ordered]@{ topic = $key; tip = $tips[$key] }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    }

    # Fuzzy match
    $matched = [System.Collections.Generic.List[object]]::new()
    foreach ($kv in $tips.GetEnumerator()) {
        if ($kv.Key -like "*$key*" -or $kv.Value -like "*$key*") {
            $matched.Add([ordered]@{ topic = $kv.Key; tip = $kv.Value })
        }
    }

    if ($matched.Count -gt 0) {
        if ($matched.Count -eq 1) {
            return (Format-AccessOutput -AsJson:$AsJson -Data $matched[0])
        }
        $result = [ordered]@{ query = $Topic; matches = @($matched) }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    }

    $result = [ordered]@{
        query            = $Topic
        error            = "No tips found matching '$Topic'"
        available_topics = @($tips.Keys)
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

# ═══════════════════════════════════════════════════════════════════════════
# NATIVE INTEROP — Phase 7 UI Automation (Win32 + GDI)
# ═══════════════════════════════════════════════════════════════════════════

if (-not ([System.Management.Automation.PSTypeName]'AccessPoshUI').Type) {
    Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class AccessPoshUI
{
    [DllImport("user32.dll")]
    public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hwnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hwnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern short VkKeyScanW(char ch);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    public static extern IntPtr GetWindowDC(IntPtr hwnd);

    [DllImport("user32.dll")]
    public static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int width, int height);

    [DllImport("gdi32.dll")]
    public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hObject);

    [DllImport("gdi32.dll")]
    public static extern bool DeleteObject(IntPtr hObject);

    [DllImport("gdi32.dll")]
    public static extern bool DeleteDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    public static extern bool BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest,
                                     IntPtr hdcSrc, int xSrc, int ySrc, uint rop);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    // Constants
    public const uint PW_RENDERFULLCONTENT = 2;
    public const int  SW_RESTORE           = 9;
    public const uint SRCCOPY              = 0x00CC0020;
}
'@
}

# ── 7.2 Get-AccessScreenshot ─────────────────────────────────────────────

function Get-AccessScreenshot {
    <#
    .SYNOPSIS
        Capture a screenshot of the Access window and save as PNG.
    .DESCRIPTION
        Optionally opens a form or report first, captures the Access window
        via PrintWindow (PW_RENDERFULLCONTENT), optionally scales down to
        MaxWidth, and saves as PNG. Returns path, dimensions, and file size.
    .EXAMPLE
        Get-AccessScreenshot -DbPath C:\my.accdb -AsJson
    .EXAMPLE
        Get-AccessScreenshot -DbPath C:\my.accdb -ObjectType form -ObjectName MainMenu -MaxWidth 1024
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DbPath,

        [ValidateSet('form','report')]
        [string]$ObjectType,

        [string]$ObjectName,

        [string]$OutputPath,

        [int]$WaitMs = 300,

        [int]$MaxWidth = 1920,

        [switch]$AsJson
    )

    $app = Connect-AccessDB -DbPath $DbPath

    $weOpened = $false
    try {
        # Optionally open a form or report
        if ($ObjectType -and $ObjectName) {
            switch ($ObjectType) {
                'form'   { $app.DoCmd.OpenForm($ObjectName)   }
                'report' { $app.DoCmd.OpenReport($ObjectName, 2 <# acViewPreview #>) }
            }
            $weOpened = $true
            Start-Sleep -Milliseconds $WaitMs
        }

        # Get window handle
        $hwnd = Get-AccessHwnd -App $app
        if ($hwnd -eq 0) {
            throw 'Could not obtain Access window handle.'
        }
        $hWndPtr = [IntPtr]::new($hwnd)

        # Restore if minimized
        if ([AccessPoshUI]::IsIconic($hWndPtr)) {
            [AccessPoshUI]::ShowWindow($hWndPtr, [AccessPoshUI]::SW_RESTORE) | Out-Null
            Start-Sleep -Milliseconds 200
        }

        # Bring to foreground for reliable capture
        [AccessPoshUI]::SetForegroundWindow($hWndPtr) | Out-Null
        Start-Sleep -Milliseconds 100

        # Get window dimensions
        $rect = New-Object AccessPoshUI+RECT
        [AccessPoshUI]::GetWindowRect($hWndPtr, [ref]$rect) | Out-Null
        $w = $rect.Right  - $rect.Left
        $h = $rect.Bottom - $rect.Top
        if ($w -le 0 -or $h -le 0) {
            throw "Invalid window dimensions: ${w}x${h}"
        }

        $origW = $w
        $origH = $h

        # Capture via PrintWindow into a System.Drawing.Bitmap
        $bmp = [System.Drawing.Bitmap]::new($w, $h)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $hdc = $g.GetHdc()
        try {
            [AccessPoshUI]::PrintWindow($hWndPtr, $hdc, [AccessPoshUI]::PW_RENDERFULLCONTENT) | Out-Null
        } finally {
            $g.ReleaseHdc($hdc)
        }

        # Resize if wider than MaxWidth
        if ($w -gt $MaxWidth) {
            $ratio  = $MaxWidth / $w
            $newH   = [int]($h * $ratio)
            $resized = [System.Drawing.Bitmap]::new($bmp, $MaxWidth, $newH)
            $bmp.Dispose()
            $bmp = $resized
            $w = $MaxWidth
            $h = $newH
        }

        # Determine output path
        if (-not $OutputPath) {
            $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
            $OutputPath = [System.IO.Path]::Combine(
                [System.IO.Path]::GetTempPath(),
                "access_screenshot_${stamp}.png"
            )
        }

        # Save PNG
        $bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $fileSize = (Get-Item -LiteralPath $OutputPath).Length

        $result = [ordered]@{
            status              = 'captured'
            path                = $OutputPath
            width               = $w
            height              = $h
            original_width      = $origW
            original_height     = $origH
            file_size           = $fileSize
        }
        if ($ObjectType -and $ObjectName) {
            $result['object_type'] = $ObjectType
            $result['object_name'] = $ObjectName
        }

        Format-AccessOutput -AsJson:$AsJson -Data $result
    } catch {
        $err = [ordered]@{ status = 'error'; error = $_.Exception.Message }
        Format-AccessOutput -AsJson:$AsJson -Data $err
    } finally {
        # Clean up GDI resources
        if ($null -ne $bmp) { $bmp.Dispose() }
        if ($null -ne $g)   { $g.Dispose()   }

        # Close the form/report if we opened it
        if ($weOpened -and $ObjectType -and $ObjectName) {
            try {
                switch ($ObjectType) {
                    'form'   { $app.DoCmd.Close(2, $ObjectName, 1 <# acSaveNo #>) }
                    'report' { $app.DoCmd.Close(3, $ObjectName, 1 <# acSaveNo #>) }
                }
            } catch {
                Write-Verbose "Could not close $ObjectType '$ObjectName': $_"
            }
        }
    }
}

# ── 7.3 Send-AccessClick ─────────────────────────────────────────────────

function Send-AccessClick {
    <#
    .SYNOPSIS
        Send a mouse click to the Access window at image-relative coordinates.
    .DESCRIPTION
        Scales X/Y from reference image coordinates to actual screen coordinates
        using the Access window rect and ImageWidth, then performs left, double,
        or right click via mouse_event.
    .EXAMPLE
        Send-AccessClick -DbPath C:\my.accdb -X 150 -Y 200 -ImageWidth 1024
    .EXAMPLE
        Send-AccessClick -DbPath C:\my.accdb -X 300 -Y 50 -ImageWidth 1920 -ClickType double
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DbPath,

        [Parameter(Mandatory)]
        [int]$X,

        [Parameter(Mandatory)]
        [int]$Y,

        [Parameter(Mandatory)]
        [int]$ImageWidth,

        [ValidateSet('left','double','right')]
        [string]$ClickType = 'left',

        [int]$WaitAfterMs = 200,

        [switch]$AsJson
    )

    # mouse_event flag constants
    $LEFTDOWN  = [uint32]0x0002
    $LEFTUP    = [uint32]0x0004
    $RIGHTDOWN = [uint32]0x0008
    $RIGHTUP   = [uint32]0x0010

    $app = Connect-AccessDB -DbPath $DbPath

    try {
        $hwnd = Get-AccessHwnd -App $app
        if ($hwnd -eq 0) { throw 'Could not obtain Access window handle.' }
        $hWndPtr = [IntPtr]::new($hwnd)

        # Restore if minimized
        if ([AccessPoshUI]::IsIconic($hWndPtr)) {
            [AccessPoshUI]::ShowWindow($hWndPtr, [AccessPoshUI]::SW_RESTORE) | Out-Null
            Start-Sleep -Milliseconds 200
        }

        [AccessPoshUI]::SetForegroundWindow($hWndPtr) | Out-Null
        Start-Sleep -Milliseconds 50

        # Get window rect and compute scale
        $rect = New-Object AccessPoshUI+RECT
        [AccessPoshUI]::GetWindowRect($hWndPtr, [ref]$rect) | Out-Null
        $winW = $rect.Right - $rect.Left
        if ($winW -le 0) { throw "Invalid window width: $winW" }

        $scale   = $winW / $ImageWidth
        $screenX = [int]($rect.Left + $X * $scale)
        $screenY = [int]($rect.Top  + $Y * $scale)

        # Move cursor
        [AccessPoshUI]::SetCursorPos($screenX, $screenY) | Out-Null
        Start-Sleep -Milliseconds 30

        # Perform click
        switch ($ClickType) {
            'left' {
                [AccessPoshUI]::mouse_event($LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
                [AccessPoshUI]::mouse_event($LEFTUP,   0, 0, 0, [UIntPtr]::Zero)
            }
            'double' {
                [AccessPoshUI]::mouse_event($LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
                [AccessPoshUI]::mouse_event($LEFTUP,   0, 0, 0, [UIntPtr]::Zero)
                Start-Sleep -Milliseconds 50
                [AccessPoshUI]::mouse_event($LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
                [AccessPoshUI]::mouse_event($LEFTUP,   0, 0, 0, [UIntPtr]::Zero)
            }
            'right' {
                [AccessPoshUI]::mouse_event($RIGHTDOWN, 0, 0, 0, [UIntPtr]::Zero)
                [AccessPoshUI]::mouse_event($RIGHTUP,   0, 0, 0, [UIntPtr]::Zero)
            }
        }

        Start-Sleep -Milliseconds $WaitAfterMs

        Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            status     = 'clicked'
            screen_x   = $screenX
            screen_y   = $screenY
            image_x    = $X
            image_y    = $Y
            click_type = $ClickType
            scale      = [math]::Round($scale, 4)
        })
    } catch {
        Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            status = 'error'
            error  = $_.Exception.Message
        })
    }
}

# ── 7.4 Send-AccessKeyboard ─────────────────────────────────────────────

function Send-AccessKeyboard {
    <#
    .SYNOPSIS
        Send keyboard input (text or special keys) to the Access window.
    .DESCRIPTION
        Types text via WM_CHAR SendMessage, or sends special-key combos with
        optional modifiers (ctrl, shift, alt) via keybd_event.
    .EXAMPLE
        Send-AccessKeyboard -DbPath C:\my.accdb -Text "Hello World"
    .EXAMPLE
        Send-AccessKeyboard -DbPath C:\my.accdb -Key enter
    .EXAMPLE
        Send-AccessKeyboard -DbPath C:\my.accdb -Key "s" -Modifiers "ctrl"
    .EXAMPLE
        Send-AccessKeyboard -DbPath C:\my.accdb -Key "a" -Modifiers "ctrl+shift"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DbPath,

        [string]$Text,

        [string]$Key,

        [string]$Modifiers,

        [int]$WaitAfterMs = 100,

        [switch]$AsJson
    )

    # Virtual-key code map for special keys
    $VK_MAP = @{
        enter     = 0x0D; tab       = 0x09; escape    = 0x1B
        backspace = 0x08; delete    = 0x2E; space     = 0x20
        up        = 0x26; down      = 0x28; left      = 0x25; right = 0x27
        home      = 0x24; 'end'     = 0x23
        pageup    = 0x21; pagedown  = 0x22
        f1  = 0x70; f2  = 0x71; f3  = 0x72;  f4  = 0x73
        f5  = 0x74; f6  = 0x75; f7  = 0x76;  f8  = 0x77
        f9  = 0x78; f10 = 0x79; f11 = 0x7A;  f12 = 0x7B
    }

    # Modifier name → virtual-key code
    $MOD_MAP = @{
        ctrl  = 0x11
        shift = 0x10
        alt   = 0x12
    }

    $KEYEVENTF_KEYUP = [uint32]2

    if (-not $Text -and -not $Key) {
        throw 'At least one of -Text or -Key must be specified.'
    }

    $app = Connect-AccessDB -DbPath $DbPath

    try {
        $hwnd = Get-AccessHwnd -App $app
        if ($hwnd -eq 0) { throw 'Could not obtain Access window handle.' }
        $hWndPtr = [IntPtr]::new($hwnd)

        # Restore if minimized
        if ([AccessPoshUI]::IsIconic($hWndPtr)) {
            [AccessPoshUI]::ShowWindow($hWndPtr, [AccessPoshUI]::SW_RESTORE) | Out-Null
            Start-Sleep -Milliseconds 200
        }

        [AccessPoshUI]::SetForegroundWindow($hWndPtr) | Out-Null
        Start-Sleep -Milliseconds 50

        $action = ''

        # ── Type text via WM_CHAR ──
        if ($Text) {
            foreach ($ch in $Text.ToCharArray()) {
                [AccessPoshUI]::SendMessage($hWndPtr, 0x0102, [IntPtr]::new([int][char]$ch), [IntPtr]::Zero) | Out-Null
            }
            $action = "typed $($Text.Length) character(s)"
        }

        # ── Send special key / key combo ──
        if ($Key) {
            $keyLower = $Key.ToLower()

            # Resolve virtual key code
            if ($VK_MAP.ContainsKey($keyLower)) {
                $vk = [byte]$VK_MAP[$keyLower]
            } else {
                # Single character — use VkKeyScanW
                if ($Key.Length -eq 1) {
                    $scan = [AccessPoshUI]::VkKeyScanW([char]$Key)
                    $vk = [byte]($scan -band 0xFF)
                } else {
                    throw "Unknown key name: '$Key'. Use a VK_MAP name or a single character."
                }
            }

            # Parse modifier string (e.g. "ctrl+shift")
            $modVks = @()
            if ($Modifiers) {
                foreach ($m in ($Modifiers.ToLower() -split '\+')) {
                    $m = $m.Trim()
                    if (-not $MOD_MAP.ContainsKey($m)) {
                        throw "Unknown modifier: '$m'. Use ctrl, shift, or alt."
                    }
                    $modVks += [byte]$MOD_MAP[$m]
                }
            }

            # Press modifiers down
            foreach ($mv in $modVks) {
                [AccessPoshUI]::keybd_event($mv, 0, 0, [UIntPtr]::Zero)
            }

            # Press and release the main key
            [AccessPoshUI]::keybd_event($vk, 0, 0, [UIntPtr]::Zero)
            [AccessPoshUI]::keybd_event($vk, 0, $KEYEVENTF_KEYUP, [UIntPtr]::Zero)

            # Release modifiers in reverse order
            for ($i = $modVks.Count - 1; $i -ge 0; $i--) {
                [AccessPoshUI]::keybd_event($modVks[$i], 0, $KEYEVENTF_KEYUP, [UIntPtr]::Zero)
            }

            $keyDesc = if ($Modifiers) { "$Modifiers+$Key" } else { $Key }
            $action = if ($action) { "$action; sent key $keyDesc" } else { "sent key $keyDesc" }
        }

        Start-Sleep -Milliseconds $WaitAfterMs

        $result = [ordered]@{
            status    = 'sent'
            action    = $action
        }
        if ($Modifiers) { $result['modifiers'] = $Modifiers }
        if ($Key)       { $result['key']       = $Key }
        if ($Text)      { $result['text_length'] = $Text.Length }

        Format-AccessOutput -AsJson:$AsJson -Data $result
    } catch {
        Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            status = 'error'
            error  = $_.Exception.Message
        })
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# Script loaded message
# ═══════════════════════════════════════════════════════════════════════════
Write-Host 'Access-POSH loaded. Use Close-AccessDatabase to release COM when done.' -ForegroundColor Cyan

