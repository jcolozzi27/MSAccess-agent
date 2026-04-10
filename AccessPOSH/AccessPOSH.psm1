<#
.SYNOPSIS
    Access-POSH — PowerShell Access Database Automation

.DESCRIPTION
    Provides full COM automation of Microsoft Access databases (.accdb/.mdb).
    Port of the Python MCP-Access server (59 tools) to native PowerShell.
    No MCP server needed — AI agents call functions directly via terminal.

    Usage:
        Import-Module .\AccessPOSH -Force
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
# DOT-SOURCE ALL SUB-FILES
# ═══════════════════════════════════════════════════════════════════════════

# Private helpers first (session, utilities, binary, vbe, design)
foreach ($file in (Get-ChildItem -Path "$PSScriptRoot\Private\*.ps1" -ErrorAction SilentlyContinue)) {
    . $file.FullName
}

# Public domain files (database, table, vbe, form, metadata, export, ui)
foreach ($file in (Get-ChildItem -Path "$PSScriptRoot\Public\*.ps1" -ErrorAction SilentlyContinue)) {
    . $file.FullName
}

# ═══════════════════════════════════════════════════════════════════════════
# CLEANUP ON EXIT
# ═══════════════════════════════════════════════════════════════════════════

Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    if ($null -ne $script:AccessSession -and $null -ne $script:AccessSession.App) {
        try { $script:AccessSession.App.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:AccessSession.App) } catch {}
        $script:AccessSession.App    = $null
        $script:AccessSession.DbPath = $null
    }
} | Out-Null

# ═══════════════════════════════════════════════════════════════════════════
# LOADED
# ═══════════════════════════════════════════════════════════════════════════
Write-Host 'AccessPOSH module loaded. Use Close-AccessDatabase to release COM when done.' -ForegroundColor Cyan
