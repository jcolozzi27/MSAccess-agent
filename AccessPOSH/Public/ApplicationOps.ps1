# Public/ApplicationOps.ps1 — Application environment, runtime detection, and file info

function Get-AccessApplicationInfo {
    <#
    .SYNOPSIS
        Get comprehensive Access application information.
    .DESCRIPTION
        Returns version, build, bitness, and runtime detection for the running Access instance.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessApplicationInfo'
    $app = Connect-AccessDB -DbPath $DbPath

    $version = $app.Version
    $build = $app.Build
    $productCode = ''
    try { $productCode = $app.ProductCode } catch {}

    # Detect bitness from the process
    $hwnd = $null
    try {
        $hwnd = $app.hWndAccessApp
    } catch {
        try { $hwnd = $app.hWndAccessApp() } catch {}
    }

    $bitness = 'unknown'
    if ($hwnd) {
        try {
            $processId = [uint32]0
            [void][AccessPoshNative]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$processId)
            $proc = Get-Process -Id $processId -ErrorAction SilentlyContinue
            if ($proc) {
                # Check if the process is 32-bit on 64-bit OS
                if ([Environment]::Is64BitOperatingSystem) {
                    # If the module file path contains 'x86', it's 32-bit
                    if ($proc.Path -match 'Program Files \(x86\)') {
                        $bitness = '32-bit'
                    } else {
                        $bitness = '64-bit'
                    }
                } else {
                    $bitness = '32-bit'
                }
            }
        } catch {}
    }

    # Detect runtime vs full
    $isRuntime = $false
    try {
        # Runtime edition throws errors when accessing the VBE
        $null = $app.VBE.Version
        $isRuntime = $false
    } catch {
        $isRuntime = $true
    }

    $result = [ordered]@{
        database     = (Split-Path $DbPath -Leaf)
        version      = $version
        build        = $build
        product_code = $productCode
        bitness      = $bitness
        is_runtime   = $isRuntime
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Test-AccessRuntime {
    <#
    .SYNOPSIS
        Quick check if Access is the Runtime edition.
    .DESCRIPTION
        Returns whether the running Access instance is the limited Runtime edition (no VBE access).
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Test-AccessRuntime'
    $app = Connect-AccessDB -DbPath $DbPath

    $isRuntime = $false
    try {
        # Access Runtime does not expose VBE at all
        $null = $app.VBE.Version
        $isRuntime = $false
    } catch {
        $isRuntime = $true
    }

    $result = [ordered]@{
        database   = (Split-Path $DbPath -Leaf)
        is_runtime = $isRuntime
        version    = $app.Version
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Get-AccessFileInfo {
    <#
    .SYNOPSIS
        Get database file-level information.
    .DESCRIPTION
        Returns file size, dates, format, version, and object counts for the specified Access database.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessFileInfo'
    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()

    $resolvedPath = (Resolve-Path $DbPath).Path
    $fileInfo = Get-Item -LiteralPath $resolvedPath

    # Get database version/format from DAO
    $dbVersion = $db.Version

    # Determine file format from extension and version
    $format = switch -Wildcard ($fileInfo.Extension.ToLower()) {
        '.accdb' { 'Access 2007+ (.accdb)' }
        '.accde' { 'Access 2007+ compiled (.accde)' }
        '.accdr' { 'Access 2007+ runtime (.accdr)' }
        '.mdb'   { 'Access 97-2003 (.mdb)' }
        '.mde'   { 'Access 97-2003 compiled (.mde)' }
        default  { $fileInfo.Extension }
    }

    # Count objects
    $tableCount = 0
    for ($i = 0; $i -lt $db.TableDefs.Count; $i++) {
        $name = $db.TableDefs($i).Name
        if (-not $name.StartsWith('MSys') -and -not $name.StartsWith('~')) { $tableCount++ }
    }

    $queryCount = $db.QueryDefs.Count

    $result = [ordered]@{
        database       = $fileInfo.Name
        full_path      = $resolvedPath
        file_size_kb   = [math]::Round($fileInfo.Length / 1024, 1)
        file_size_mb   = [math]::Round($fileInfo.Length / 1MB, 2)
        created        = $fileInfo.CreationTime.ToString('o')
        last_modified  = $fileInfo.LastWriteTime.ToString('o')
        format         = $format
        db_version     = $dbVersion
        table_count    = $tableCount
        query_count    = $queryCount
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}
