# Public/ExportTransfer.ps1 — Export reports and import/export data

function Export-AccessReport {
    <#
    .SYNOPSIS
        Export a report, table, query, or form to PDF, XLSX, RTF, or TXT via DoCmd.OutputTo.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$ObjectName,
        [ValidateSet('report','table','query','form')]
        [string]$ObjectType = 'report',
        [ValidateSet('pdf','xlsx','rtf','txt')]
        [string]$OutputFormat = 'pdf',
        [string]$OutputPath,
        [switch]$OpenAfterExport,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Export-AccessReport'
    if (-not $ObjectName) { throw "Export-AccessReport: -ObjectName is required." }

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

function Copy-AccessData {
    <#
    .SYNOPSIS
        Import or export data via DoCmd.TransferSpreadsheet / DoCmd.TransferText.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateSet('import','export')][string]$Action,
        [string]$FilePath,
        [ValidateNotNullOrEmpty()]
        [string]$TableName,
        [bool]$HasHeaders = $true,
        [ValidateSet('xlsx','xls','excel','csv','txt','text')]
        [string]$FileType = 'xlsx',
        [string]$Range,
        [string]$SpecName,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Copy-AccessData'
    if (-not $Action) { throw "Copy-AccessData: -Action is required (import, export)." }
    if (-not $FilePath) { throw "Copy-AccessData: -FilePath is required." }
    if (-not $TableName) { throw "Copy-AccessData: -TableName is required." }

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

function Export-AccessSource {
    <#
    .SYNOPSIS
        Export all (or filtered) Access database objects to a folder structure.
        Uses SaveAsText for forms, reports, queries, macros, modules.
        Uses DAO metadata for table schema (CREATE TABLE DDL).
        Optionally exports table row data via TransferText.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER OutputFolder
        Destination folder. Created if it doesn't exist.
    .PARAMETER ObjectType
        One or more object types to export. Default exports all.
    .PARAMETER IncludeTableData
        Also export table row data as CSV files.
    .PARAMETER ClearFirst
        Remove existing OutputFolder contents before exporting.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Export-AccessSource -DbPath "C:\db.accdb" -OutputFolder "C:\src"
    .EXAMPLE
        Export-AccessSource -DbPath "C:\db.accdb" -OutputFolder "C:\src" -ObjectType form,module -AsJson
    .EXAMPLE
        Export-AccessSource -DbPath "C:\db.accdb" -OutputFolder "C:\src" -IncludeTableData -ClearFirst
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [string]$OutputFolder,
        [ValidateSet('form','report','query','macro','module','table')]
        [string[]]$ObjectType,
        [switch]$IncludeTableData,
        [switch]$ClearFirst,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Export-AccessSource'
    if (-not $OutputFolder) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-OutputFolder is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $OutputFolder
            )
        )
    }

    $OutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)
    $app = Connect-AccessDB -DbPath $DbPath

    # Clear folder if requested
    if ($ClearFirst -and (Test-Path -LiteralPath $OutputFolder)) {
        Remove-Item -LiteralPath $OutputFolder -Recurse -Force
    }

    # Create subfolders
    $subfolders = @('forms', 'reports', 'queries', 'macros', 'modules', 'tables')
    foreach ($sub in $subfolders) {
        $p = [System.IO.Path]::Combine($OutputFolder, $sub)
        if (-not (Test-Path -LiteralPath $p)) {
            New-Item -Path $p -ItemType Directory -Force | Out-Null
        }
    }

    # Enumerate all objects
    $allObjects = Get-AccessObject -DbPath $DbPath

    # Determine which types to export
    $typesToExport = if ($ObjectType) { $ObjectType } else { @('form','report','query','macro','module','table') }

    $files   = [System.Collections.Generic.List[object]]::new()
    $errors  = [System.Collections.Generic.List[object]]::new()
    $objects = [ordered]@{
        forms   = @(); reports = @(); queries = @()
        macros  = @(); modules = @(); tables  = @()
    }

    # ── Forms & Reports: SaveAsText → temp → strip binaries → write ──
    foreach ($type in @('form', 'report')) {
        if ($type -notin $typesToExport) { continue }
        $pluralKey = "${type}s"
        $subDir    = [System.IO.Path]::Combine($OutputFolder, "${type}s")
        $names     = @($allObjects."$type")
        if (-not $names) { $names = @() }

        foreach ($name in $names) {
            $tmp = [System.IO.Path]::GetTempFileName()
            try {
                $app.SaveAsText($script:AC_TYPE[$type], $name, $tmp)
                $fileResult = Read-TempFile -Path $tmp
                $text = Remove-BinarySections -Text $fileResult.Content
                $outPath = [System.IO.Path]::Combine($subDir, "$name.txt")
                Write-TempFile -Path $outPath -Content $text -Encoding $fileResult.Encoding
                $relPath = "${type}s/$name.txt"
                $files.Add([PSCustomObject][ordered]@{ type = $type; name = $name; path = $relPath })
                $objects[$pluralKey] = @($objects[$pluralKey]) + @($name)
            } catch {
                $errors.Add([PSCustomObject][ordered]@{ type = $type; name = $name; error = $_.Exception.Message })
            } finally {
                Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # ── Queries & Macros: SaveAsText directly ──
    foreach ($type in @('query', 'macro')) {
        if ($type -notin $typesToExport) { continue }
        $pluralKey = if ($type -eq 'query') { 'queries' } else { 'macros' }
        $subDir    = [System.IO.Path]::Combine($OutputFolder, $pluralKey)
        $names     = @($allObjects."$type")
        if (-not $names) { $names = @() }

        foreach ($name in $names) {
            try {
                $safeName = $name -replace '[\\/:*?"<>|]', '_'
                $outPath = [System.IO.Path]::Combine($subDir, "$safeName.txt")
                $app.SaveAsText($script:AC_TYPE[$type], $name, $outPath)
                $relPath = "$pluralKey/$safeName.txt"
                $files.Add([PSCustomObject][ordered]@{ type = $type; name = $name; path = $relPath })
                $objects[$pluralKey] = @($objects[$pluralKey]) + @($name)
            } catch {
                $errors.Add([PSCustomObject][ordered]@{ type = $type; name = $name; error = $_.Exception.Message })
            }
        }
    }

    # ── Modules: SaveAsText directly (.bas extension) ──
    if ('module' -in $typesToExport) {
        $subDir = [System.IO.Path]::Combine($OutputFolder, 'modules')
        $names  = @($allObjects.module)
        if (-not $names) { $names = @() }

        foreach ($name in $names) {
            try {
                $safeName = $name -replace '[\\/:*?"<>|]', '_'
                $outPath = [System.IO.Path]::Combine($subDir, "$safeName.bas")
                $app.SaveAsText($script:AC_TYPE['module'], $name, $outPath)
                $relPath = "modules/$safeName.bas"
                $files.Add([PSCustomObject][ordered]@{ type = 'module'; name = $name; path = $relPath })
                $objects['modules'] = @($objects['modules']) + @($name)
            } catch {
                $errors.Add([PSCustomObject][ordered]@{ type = 'module'; name = $name; error = $_.Exception.Message })
            }
        }
    }

    # ── Tables: DDL schema + optional CSV data ──
    if ('table' -in $typesToExport) {
        $subDir = [System.IO.Path]::Combine($OutputFolder, 'tables')
        $names  = @($allObjects.table)
        if (-not $names) { $names = @() }

        foreach ($name in $names) {
            $safeName = $name -replace '[\\/:*?"<>|]', '_'

            # Schema DDL
            try {
                $ddl = Get-TableSchemaDDL -App $app -TableName $name
                $sqlPath = [System.IO.Path]::Combine($subDir, "$safeName.sql")
                [System.IO.File]::WriteAllText($sqlPath, $ddl, [System.Text.Encoding]::UTF8)
                $files.Add([PSCustomObject][ordered]@{ type = 'table_schema'; name = $name; path = "tables/$safeName.sql" })
                $objects['tables'] = @($objects['tables']) + @($name)
            } catch {
                $errors.Add([PSCustomObject][ordered]@{ type = 'table_schema'; name = $name; error = $_.Exception.Message })
            }

            # Data CSV (opt-in)
            if ($IncludeTableData) {
                try {
                    $csvPath = [System.IO.Path]::Combine($subDir, "$safeName.csv")
                    $app.DoCmd.TransferText($script:AC_EXPORT_DELIM, '', $name, $csvPath, $true)
                    $files.Add([PSCustomObject][ordered]@{ type = 'table_data'; name = $name; path = "tables/$safeName.csv" })
                } catch {
                    $errors.Add([PSCustomObject][ordered]@{ type = 'table_data'; name = $name; error = $_.Exception.Message })
                }
            }
        }
    }

    # ── Write manifest ──
    $manifest = [ordered]@{
        version     = '1.0'
        source_db   = $DbPath
        export_date = [DateTime]::UtcNow.ToString('o')
        objects     = $objects
        files       = @($files)
        errors      = @($errors)
        summary     = [ordered]@{
            exported = $files.Count
            failed   = $errors.Count
        }
    }

    $manifestPath = [System.IO.Path]::Combine($OutputFolder, 'manifest.json')
    $manifestJson = $manifest | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($manifestPath, $manifestJson, [System.Text.Encoding]::UTF8)

    Format-AccessOutput -AsJson:$AsJson -Data $manifest
}

function Import-AccessSource {
    <#
    .SYNOPSIS
        Import Access database objects from a folder previously created by Export-AccessSource.
        Imports in dependency order: tables → queries → modules → macros → forms → reports.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER InputFolder
        Source folder containing exported files (with manifest.json or standard subfolders).
    .PARAMETER ObjectType
        One or more object types to import. Default imports all found.
    .PARAMETER OverwriteExisting
        Delete existing objects before importing. Without this, existing objects are skipped.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Import-AccessSource -DbPath "C:\new.accdb" -InputFolder "C:\src"
    .EXAMPLE
        Import-AccessSource -DbPath "C:\db.accdb" -InputFolder "C:\src" -ObjectType form,module -OverwriteExisting
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [string]$InputFolder,
        [ValidateSet('form','report','query','macro','module','table')]
        [string[]]$ObjectType,
        [switch]$OverwriteExisting,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Import-AccessSource'
    if (-not $InputFolder) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-InputFolder is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $InputFolder
            )
        )
    }

    $InputFolder = [System.IO.Path]::GetFullPath($InputFolder)
    if (-not (Test-Path -LiteralPath $InputFolder -PathType Container)) {
        throw "Import-AccessSource: InputFolder not found: $InputFolder"
    }

    $app = Connect-AccessDB -DbPath $DbPath

    # Determine which types to import
    $typesToImport = if ($ObjectType) { $ObjectType } else { @('table','query','module','macro','form','report') }

    # Get existing objects for collision detection
    $existing = Get-AccessObject -DbPath $DbPath

    $imported = [System.Collections.Generic.List[object]]::new()
    $skipped  = [System.Collections.Generic.List[object]]::new()
    $errors   = [System.Collections.Generic.List[object]]::new()

    # Helper: check if object exists
    $existsCheck = {
        param([string]$type, [string]$name)
        $col = $existing."$type"
        if (-not $col) { return $false }
        return ($col -contains $name)
    }

    # ── 1. Tables (SQL DDL + CSV data) ──
    if ('table' -in $typesToImport) {
        $tableDir = [System.IO.Path]::Combine($InputFolder, 'tables')
        if (Test-Path -LiteralPath $tableDir) {
            $sqlFiles = Get-ChildItem -LiteralPath $tableDir -Filter '*.sql' -File -ErrorAction SilentlyContinue
            foreach ($sqlFile in $sqlFiles) {
                $tableName = [System.IO.Path]::GetFileNameWithoutExtension($sqlFile.Name)
                $tableExists = & $existsCheck 'table' $tableName

                if ($tableExists -and -not $OverwriteExisting) {
                    $skipped.Add([PSCustomObject][ordered]@{ type = 'table'; name = $tableName; reason = 'exists' })
                    continue
                }

                try {
                    if ($tableExists -and $OverwriteExisting) {
                        $app.DoCmd.DeleteObject($script:AC_TYPE['query'], $tableName) 2>$null
                        $db = $app.CurrentDb()
                        try { $db.TableDefs.Delete($tableName) } catch {}
                        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($db)
                    }

                    $ddl = [System.IO.File]::ReadAllText($sqlFile.FullName, [System.Text.Encoding]::UTF8)
                    $db = $app.CurrentDb()
                    try {
                        $db.Execute($ddl)
                    } finally {
                        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($db)
                    }

                    $imported.Add([PSCustomObject][ordered]@{ type = 'table_schema'; name = $tableName })

                    # Import CSV data if available
                    $csvPath = [System.IO.Path]::Combine($tableDir, "$tableName.csv")
                    if (Test-Path -LiteralPath $csvPath) {
                        $app.DoCmd.TransferText($script:AC_IMPORT, '', $tableName, $csvPath, $true)
                        $imported.Add([PSCustomObject][ordered]@{ type = 'table_data'; name = $tableName })
                    }
                } catch {
                    $errors.Add([PSCustomObject][ordered]@{ type = 'table'; name = $tableName; error = $_.Exception.Message })
                }
            }
        }
    }

    # ── 2. Queries ──
    if ('query' -in $typesToImport) {
        $queryDir = [System.IO.Path]::Combine($InputFolder, 'queries')
        if (Test-Path -LiteralPath $queryDir) {
            $queryFiles = Get-ChildItem -LiteralPath $queryDir -Filter '*.txt' -File -ErrorAction SilentlyContinue
            foreach ($f in $queryFiles) {
                $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                $objExists = & $existsCheck 'query' $name

                if ($objExists -and -not $OverwriteExisting) {
                    $skipped.Add([PSCustomObject][ordered]@{ type = 'query'; name = $name; reason = 'exists' })
                    continue
                }
                try {
                    if ($objExists -and $OverwriteExisting) {
                        $app.DoCmd.DeleteObject($script:AC_TYPE['query'], $name)
                    }
                    $app.LoadFromText($script:AC_TYPE['query'], $name, $f.FullName)
                    $imported.Add([PSCustomObject][ordered]@{ type = 'query'; name = $name })
                } catch {
                    $errors.Add([PSCustomObject][ordered]@{ type = 'query'; name = $name; error = $_.Exception.Message })
                }
            }
        }
    }

    # ── 3. Modules ──
    if ('module' -in $typesToImport) {
        $modDir = [System.IO.Path]::Combine($InputFolder, 'modules')
        if (Test-Path -LiteralPath $modDir) {
            $modFiles = Get-ChildItem -LiteralPath $modDir -Filter '*.bas' -File -ErrorAction SilentlyContinue
            foreach ($f in $modFiles) {
                $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                $objExists = & $existsCheck 'module' $name

                if ($objExists -and -not $OverwriteExisting) {
                    $skipped.Add([PSCustomObject][ordered]@{ type = 'module'; name = $name; reason = 'exists' })
                    continue
                }
                try {
                    if ($objExists -and $OverwriteExisting) {
                        $app.DoCmd.DeleteObject($script:AC_TYPE['module'], $name)
                    }
                    $app.LoadFromText($script:AC_TYPE['module'], $name, $f.FullName)
                    $imported.Add([PSCustomObject][ordered]@{ type = 'module'; name = $name })
                } catch {
                    $errors.Add([PSCustomObject][ordered]@{ type = 'module'; name = $name; error = $_.Exception.Message })
                }
            }
        }
    }

    # ── 4. Macros ──
    if ('macro' -in $typesToImport) {
        $macroDir = [System.IO.Path]::Combine($InputFolder, 'macros')
        if (Test-Path -LiteralPath $macroDir) {
            $macroFiles = Get-ChildItem -LiteralPath $macroDir -Filter '*.txt' -File -ErrorAction SilentlyContinue
            foreach ($f in $macroFiles) {
                $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                $objExists = & $existsCheck 'macro' $name

                if ($objExists -and -not $OverwriteExisting) {
                    $skipped.Add([PSCustomObject][ordered]@{ type = 'macro'; name = $name; reason = 'exists' })
                    continue
                }
                try {
                    if ($objExists -and $OverwriteExisting) {
                        $app.DoCmd.DeleteObject($script:AC_TYPE['macro'], $name)
                    }
                    $app.LoadFromText($script:AC_TYPE['macro'], $name, $f.FullName)
                    $imported.Add([PSCustomObject][ordered]@{ type = 'macro'; name = $name })
                } catch {
                    $errors.Add([PSCustomObject][ordered]@{ type = 'macro'; name = $name; error = $_.Exception.Message })
                }
            }
        }
    }

    # ── 5. Forms ──
    if ('form' -in $typesToImport) {
        $formDir = [System.IO.Path]::Combine($InputFolder, 'forms')
        if (Test-Path -LiteralPath $formDir) {
            $formFiles = Get-ChildItem -LiteralPath $formDir -Filter '*.txt' -File -ErrorAction SilentlyContinue
            foreach ($f in $formFiles) {
                $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                $objExists = & $existsCheck 'form' $name

                if ($objExists -and -not $OverwriteExisting) {
                    $skipped.Add([PSCustomObject][ordered]@{ type = 'form'; name = $name; reason = 'exists' })
                    continue
                }
                try {
                    if ($objExists -and $OverwriteExisting) {
                        $app.DoCmd.Close($script:AC_FORM, $name, $script:AC_SAVE_NO) 2>$null
                        $app.DoCmd.DeleteObject($script:AC_TYPE['form'], $name)
                    }

                    $fileResult = Read-TempFile -Path $f.FullName
                    $code = $fileResult.Content

                    # Split code-behind and restore binary sections
                    $vbaCode = ''
                    $split = Split-CodeBehind -Code $code
                    $code    = $split.FormText
                    $vbaCode = $split.VbaCode
                    if ($vbaCode) {
                        $code = $code -replace '(?m)^\s*HasModule\s*=.*$', ''
                    }

                    $code = Restore-BinarySections -App $app -ObjectType 'form' -Name $name -NewCode $code

                    $tmp = [System.IO.Path]::GetTempFileName()
                    try {
                        Write-TempFile -Path $tmp -Content $code -Encoding 'utf-16'
                        $app.LoadFromText($script:AC_TYPE['form'], $name, $tmp)

                        if ($vbaCode) {
                            Invoke-VbaAfterImport -App $app -ObjectType 'form' -Name $name -VbaCode $vbaCode
                        }
                    } finally {
                        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                    }

                    $imported.Add([PSCustomObject][ordered]@{ type = 'form'; name = $name })
                } catch {
                    $errors.Add([PSCustomObject][ordered]@{ type = 'form'; name = $name; error = $_.Exception.Message })
                }
            }
        }
    }

    # ── 6. Reports ──
    if ('report' -in $typesToImport) {
        $reportDir = [System.IO.Path]::Combine($InputFolder, 'reports')
        if (Test-Path -LiteralPath $reportDir) {
            $reportFiles = Get-ChildItem -LiteralPath $reportDir -Filter '*.txt' -File -ErrorAction SilentlyContinue
            foreach ($f in $reportFiles) {
                $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                $objExists = & $existsCheck 'report' $name

                if ($objExists -and -not $OverwriteExisting) {
                    $skipped.Add([PSCustomObject][ordered]@{ type = 'report'; name = $name; reason = 'exists' })
                    continue
                }
                try {
                    if ($objExists -and $OverwriteExisting) {
                        $app.DoCmd.Close($script:AC_REPORT, $name, $script:AC_SAVE_NO) 2>$null
                        $app.DoCmd.DeleteObject($script:AC_TYPE['report'], $name)
                    }

                    $fileResult = Read-TempFile -Path $f.FullName
                    $code = $fileResult.Content

                    $vbaCode = ''
                    $split = Split-CodeBehind -Code $code
                    $code    = $split.FormText
                    $vbaCode = $split.VbaCode
                    if ($vbaCode) {
                        $code = $code -replace '(?m)^\s*HasModule\s*=.*$', ''
                    }

                    $code = Restore-BinarySections -App $app -ObjectType 'report' -Name $name -NewCode $code

                    $tmp = [System.IO.Path]::GetTempFileName()
                    try {
                        Write-TempFile -Path $tmp -Content $code -Encoding 'utf-16'
                        $app.LoadFromText($script:AC_TYPE['report'], $name, $tmp)

                        if ($vbaCode) {
                            Invoke-VbaAfterImport -App $app -ObjectType 'report' -Name $name -VbaCode $vbaCode
                        }
                    } finally {
                        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                    }

                    $imported.Add([PSCustomObject][ordered]@{ type = 'report'; name = $name })
                } catch {
                    $errors.Add([PSCustomObject][ordered]@{ type = 'report'; name = $name; error = $_.Exception.Message })
                }
            }
        }
    }

    # Clear all caches after batch import
    Clear-AccessCaches

    $result = [ordered]@{
        target_db = $DbPath
        imported  = @($imported)
        skipped   = @($skipped)
        errors    = @($errors)
        summary   = [ordered]@{
            imported = $imported.Count
            skipped  = $skipped.Count
            failed   = $errors.Count
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}
