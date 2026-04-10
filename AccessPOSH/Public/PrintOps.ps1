# Public/PrintOps.ps1 — Filtered report export and direct printing

function Export-AccessFilteredReport {
    <#
    .SYNOPSIS
        Export an Access report to PDF/XLSX/RTF/TXT with an optional WhereCondition filter.
    .DESCRIPTION
        Opens the report in preview mode with the specified filter applied, exports it
        to the chosen format via DoCmd.OutputTo, then closes the report.  If no OutputPath
        is supplied a temp file is generated automatically.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$ReportName,
        [string]$WhereCondition,
        [string]$FilterName,
        [ValidateSet('pdf','xlsx','rtf','txt')]
        [string]$OutputFormat = 'pdf',
        [string]$OutputPath,
        [switch]$OpenAfterExport,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Export-AccessFilteredReport'
    if (-not $ReportName) { throw "Export-AccessFilteredReport: -ReportName is required." }

    $app = Connect-AccessDB -DbPath $DbPath

    $formatString = $script:OUTPUT_FORMATS[$OutputFormat]
    if (-not $formatString) {
        throw "Unsupported output format '$OutputFormat'. Valid: $($script:OUTPUT_FORMATS.Keys -join ', ')"
    }

    if (-not $OutputPath) {
        $extMap = @{ pdf = '.pdf'; xlsx = '.xlsx'; rtf = '.rtf'; txt = '.txt' }
        $safeName = $ReportName -replace '[\\/:*?"<>|]', '_'
        $OutputPath = Join-Path $env:TEMP "$safeName$($extMap[$OutputFormat])"
    }
    $OutputPath = [System.IO.Path]::GetFullPath($OutputPath)

    # Open the report with filter applied, then output
    # DoCmd.OpenReport reportName, acViewPreview, filterName, whereCondition
    # Then DoCmd.OutputTo acOutputReport, reportName, format, outputFile
    $filterArg = if ($FilterName) { $FilterName } else { [System.Reflection.Missing]::Value }
    $whereArg  = if ($WhereCondition) { $WhereCondition } else { [System.Reflection.Missing]::Value }

    try {
        # Open report in preview mode with filter (acViewPreview = 2)
        $app.DoCmd.OpenReport($ReportName, 2, $filterArg, $whereArg)

        # Export the currently open (filtered) report
        $app.DoCmd.OutputTo(3, $ReportName, $formatString, $OutputPath, [bool]$OpenAfterExport)

        # Close the report without saving
        $app.DoCmd.Close(3, $ReportName, 2)   # 3=acReport, 2=acSaveNo
    } catch {
        # Try to close the report if it was opened
        try { $app.DoCmd.Close(3, $ReportName, 2) } catch {}
        throw "Error exporting filtered report '$ReportName': $_"
    }

    $fileSize = 0
    if (Test-Path -LiteralPath $OutputPath) {
        $fileSize = (Get-Item -LiteralPath $OutputPath).Length
    }

    $result = [ordered]@{
        database        = (Split-Path $DbPath -Leaf)
        report          = $ReportName
        output_format   = $OutputFormat
        output_path     = $OutputPath
        file_size       = $fileSize
        where_condition = if ($WhereCondition) { $WhereCondition } else { '' }
        filter_name     = if ($FilterName) { $FilterName } else { '' }
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Send-AccessReportToPrinter {
    <#
    .SYNOPSIS
        Print an Access report directly to the default printer.
    .DESCRIPTION
        Sends a report to the printer, optionally applying a WhereCondition filter,
        setting the number of copies, and restricting the page range.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$ReportName,
        [string]$WhereCondition,
        [int]$Copies = 1,
        [ValidateSet('all','pages','selection')]
        [string]$PrintRange = 'all',
        [int]$FromPage,
        [int]$ToPage,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Send-AccessReportToPrinter'
    if (-not $ReportName) { throw "Send-AccessReportToPrinter: -ReportName is required." }

    $app = Connect-AccessDB -DbPath $DbPath

    $whereArg = if ($WhereCondition) { $WhereCondition } else { [System.Reflection.Missing]::Value }

    try {
        # DoCmd.OpenReport reportName, acViewNormal (=0, sends to printer), filterName, whereCondition
        # For page range, we use the Printer object

        if ($PrintRange -eq 'all' -and $Copies -le 1 -and -not $WhereCondition) {
            # Simple print — directly send to default printer
            $app.DoCmd.OpenReport($ReportName, 0)    # acViewNormal = 0 sends to printer
        } else {
            # Open in preview first to apply filter and set print options
            $app.DoCmd.OpenReport($ReportName, 2, [System.Reflection.Missing]::Value, $whereArg)    # acViewPreview = 2

            # Set printer properties
            $printer = $app.Printer
            if ($Copies -gt 1) {
                $printer.Copies = $Copies
            }

            $printRangeMap = @{ all = 0; pages = 2; selection = 1 }
            $printer.PrintRange = $printRangeMap[$PrintRange]

            if ($PrintRange -eq 'pages') {
                if ($FromPage) { $printer.PageFrom = $FromPage }
                if ($ToPage)   { $printer.PageTo = $ToPage }
            }

            # Send to printer
            $app.DoCmd.RunCommand(340)   # acCmdPrint — prints with current settings

            # Close preview
            $app.DoCmd.Close(3, $ReportName, 2)   # acSaveNo
        }
    } catch {
        try { $app.DoCmd.Close(3, $ReportName, 2) } catch {}
        throw "Error printing report '$ReportName': $_"
    }

    $result = [ordered]@{
        database        = (Split-Path $DbPath -Leaf)
        report          = $ReportName
        action          = 'sent_to_printer'
        copies          = if ($Copies) { $Copies } else { 1 }
        print_range     = $PrintRange
        where_condition = if ($WhereCondition) { $WhereCondition } else { '' }
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}
