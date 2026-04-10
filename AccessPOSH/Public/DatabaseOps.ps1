# Public/DatabaseOps.ps1 — Database lifecycle, object CRUD, SQL execution

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
        [string]$DbPath,
        [switch]$AsJson
    )

    if (-not $DbPath) { throw "New-AccessDatabase: -DbPath is required." }

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
        # Suppress dialogs for non-interactive automation
        try {
            $script:AccessSession.App.DisplayAlerts = $false
            $script:AccessSession.App.AutomationSecurity = 1  # msoAutomationSecurityForceDisable
        } catch {}
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
        [string]$DbPath,
        [switch]$AsJson
    )

    if (-not $DbPath) { throw "Repair-AccessDatabase: -DbPath is required." }

    $resolved = [System.IO.Path]::GetFullPath($DbPath)
    $app = Connect-AccessDB -DbPath $resolved
    $originalSize = (Get-Item -LiteralPath $resolved).Length

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    try {
        $app.CloseCurrentDatabase()
    } catch {
        throw "Could not close database for compact/repair: $_"
    }
    $script:AccessSession.DbPath = $null
    Clear-AccessCaches

    Start-Sleep -Milliseconds 500

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

        Rename-Item -LiteralPath $resolved -NewName ([System.IO.Path]::GetFileName($bakPath)) -Force
        try {
            Rename-Item -LiteralPath $tmpPath -NewName ([System.IO.Path]::GetFileName($resolved)) -Force
        } catch {
            Rename-Item -LiteralPath $bakPath -NewName ([System.IO.Path]::GetFileName($resolved)) -Force
            throw
        }

        Remove-Item -LiteralPath $bakPath -Force -ErrorAction SilentlyContinue

    } catch {
        try {
            if (Test-Path -LiteralPath $resolved) {
                $app.OpenCurrentDatabase($resolved)
                $script:AccessSession.DbPath = $resolved
            }
        } catch {}
        throw
    }

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
        [string]$DbPath,
        [switch]$AsJson
    )

    if (-not $DbPath) { throw "Invoke-AccessDecompile: -DbPath is required." }

    $resolved = [System.IO.Path]::GetFullPath($DbPath)
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw "Database not found: $resolved"
    }
    $originalSize = (Get-Item -LiteralPath $resolved).Length

    # 1. Close COM session completely so the file is unlocked
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

    # 2. Find MSACCESS.EXE
    $msaccessCandidates = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\MSACCESS.EXE"
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\MSACCESS.EXE"
    )
    $msaccess = $msaccessCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $msaccess) {
        throw 'MSACCESS.EXE not found in known Office 16 paths'
    }

    # 3. Launch MSACCESS /decompile with SHIFT held
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

    Start-Sleep -Seconds 3
    if ($shiftHeld) {
        try { [AccessPoshNative]::keybd_event($VK_SHIFT, 0, $KEYEVENTF_KEYUP, [UIntPtr]::Zero) } catch {}
        Write-Verbose 'SHIFT released'
    }

    Start-Sleep -Seconds 5
    if ($null -ne $proc -and -not $proc.HasExited) {
        Write-Verbose "Killing decompile Access process (PID $($proc.Id))"
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
    }

    $decompileSize = (Get-Item -LiteralPath $resolved).Length

    # 4. Reopen via COM and try to recompile VBA
    Write-Verbose 'Relaunching COM after decompile...'
    try {
        $script:AccessSession.App = New-Object -ComObject 'Access.Application'
    } catch {
        throw "Failed to relaunch Access COM after decompile: $_"
    }
    # Suppress dialogs for non-interactive automation
    try {
        $script:AccessSession.App.DisplayAlerts = $false
        $script:AccessSession.App.AutomationSecurity = 1  # msoAutomationSecurityForceDisable
    } catch {}
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

    try {
        $script:AccessSession.App.RunCommand(137)  # acCmdCompileAllModules
        Write-Verbose 'VBA recompiled after decompile'
    } catch {
        Write-Verbose "VBA recompile skipped: $_"
    }

    try {
        $script:AccessSession.App.CloseCurrentDatabase()
    } catch {
        Write-Verbose "Error closing DB before compact: $_"
    }
    $script:AccessSession.DbPath = $null
    Clear-AccessCaches

    # 5. Compact & Repair
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

    Rename-Item -LiteralPath $resolved -NewName ([System.IO.Path]::GetFileName($bakPath)) -Force
    try {
        Rename-Item -LiteralPath $tmpPath -NewName ([System.IO.Path]::GetFileName($resolved)) -Force
    } catch {
        Rename-Item -LiteralPath $bakPath -NewName ([System.IO.Path]::GetFileName($resolved)) -Force
        throw
    }
    Remove-Item -LiteralPath $bakPath -Force -ErrorAction SilentlyContinue

    # 6. Reopen compacted database
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
        [string]$DbPath,
        [ValidateSet('all','table','query','form','report','macro','module')]
        [string]$ObjectType = 'all',
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessObject'

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
        [string]$DbPath,
        [ValidateSet('query','form','report','macro','module')]
        [string]$ObjectType,
        [string]$Name
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessCode'
    if (-not $ObjectType) { throw "Get-AccessCode: -ObjectType is required (query, form, report, macro, module)." }
    if (-not $Name) { throw "Get-AccessCode: -Name is required." }

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
        [string]$DbPath,
        [ValidateSet('query','form','report','macro','module')]
        [string]$ObjectType,
        [string]$Name,
        [string]$Code,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessCode'
    if (-not $ObjectType) { throw "Set-AccessCode: -ObjectType is required (query, form, report, macro, module)." }
    if (-not $Name) { throw "Set-AccessCode: -Name is required." }
    if (-not $Code) { throw "Set-AccessCode: -Code is required." }

    $app = Connect-AccessDB -DbPath $DbPath

    $vbaCode = ''
    if ($ObjectType -in 'form', 'report') {
        $split = Split-CodeBehind -Code $Code
        $Code    = $split.FormText
        $vbaCode = $split.VbaCode

        if ($vbaCode) {
            $Code = $Code -replace '(?m)^\s*HasModule\s*=.*$', ''
        }
    }

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
        $enc = if ($ObjectType -eq 'module') { 'cp1252' } else { 'utf-16' }
        Write-TempFile -Path $tmp -Content $Code -Encoding $enc

        try {
            $app.LoadFromText($script:AC_TYPE[$ObjectType], $Name, $tmp)
        } catch {
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

        $cacheKey = "${ObjectType}:${Name}"
        $script:AccessSession.VbeCodeCache.Remove($cacheKey)
        $script:AccessSession.CmCache.Remove($cacheKey)
        $script:AccessSession.ControlsCache.Remove($cacheKey)

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
        [string]$DbPath,
        [ValidateSet('query','form','report','macro','module')]
        [string]$ObjectType,
        [string]$Name,
        [switch]$Confirm,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Remove-AccessObject'
    if (-not $ObjectType) { throw "Remove-AccessObject: -ObjectType is required (query, form, report, macro, module)." }
    if (-not $Name) { throw "Remove-AccessObject: -Name is required." }

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
        [string]$DbPath,
        [string]$OutputPath,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Export-AccessStructure'

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

    $lines.Add("## Forms ($($forms.Count))")
    $lines.Add("")
    if ($forms.Count -gt 0) {
        foreach ($n in $forms) { $lines.Add("- ``$n``") }
    } else {
        $lines.Add('*(none)*')
    }
    $lines.Add('')

    $lines.Add("## Reports ($($reports.Count))")
    $lines.Add("")
    if ($reports.Count -gt 0) {
        foreach ($n in $reports) { $lines.Add("- ``$n``") }
    } else {
        $lines.Add('*(none)*')
    }
    $lines.Add('')

    $lines.Add("## Queries ($($queries.Count))")
    $lines.Add("")
    if ($queries.Count -gt 0) {
        foreach ($n in $queries) { $lines.Add("- ``$n``") }
    } else {
        $lines.Add('*(none)*')
    }
    $lines.Add('')

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
        [string]$DbPath,
        [string]$SQL,
        [int]$Limit = 500,
        [switch]$ConfirmDestructive,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Invoke-AccessSQL'
    if (-not $SQL) { throw "Invoke-AccessSQL: -SQL is required." }

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
                $rs = $db.OpenRecordset($SQL, 2, $script:DB_SEE_CHANGES)
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
        [string]$DbPath,
        [object[]]$Statements,
        [bool]$StopOnError = $true,
        [switch]$ConfirmDestructive,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Invoke-AccessSQLBatch'
    if (-not $Statements -or $Statements.Count -eq 0) { throw "Invoke-AccessSQLBatch: -Statements is required." }

    if ($Statements.Count -eq 0) {
        return Format-AccessOutput -AsJson:$AsJson -Data @{ error = 'No SQL statements provided.' }
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

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
