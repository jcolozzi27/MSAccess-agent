# Public/ExportTransfer.ps1 — Export reports and import/export data

function Export-AccessReport {
    <#
    .SYNOPSIS
        Export a report, table, query, or form to PDF, XLSX, RTF, or TXT via DoCmd.OutputTo.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
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
        [string]$DbPath,
        [ValidateSet('import','export')][string]$Action,
        [string]$FilePath,
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
