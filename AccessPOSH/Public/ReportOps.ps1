# Public/ReportOps.ps1 — Report creation and group level management

function New-AccessReport {
    <#
    .SYNOPSIS
        Create a new blank report in an Access database.
    .DESCRIPTION
        Creates a new report, optionally sets its RecordSource, and renames it
        to the desired name. Uses the same create-then-rename pattern as
        New-AccessForm.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$ReportName,
        [string]$RecordSource,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'New-AccessReport'
    if (-not $ReportName) { throw "New-AccessReport: -ReportName is required." }

    $app = Connect-AccessDB -DbPath $DbPath

    try {
        $rpt = $app.CreateReport()
        $autoName = $rpt.Name          # Access assigns a default name like "Report1"

        if ($RecordSource) {
            $rpt.RecordSource = $RecordSource
        }

        # Save with the auto-generated name first
        $app.DoCmd.Save()
        $app.DoCmd.Close(3, $autoName, $script:AC_SAVE_YES)  # 3 = acReport

        # Rename to the desired name
        $app.DoCmd.Rename($ReportName, 3, $autoName)  # 3 = acReport

        $result = [ordered]@{
            database      = (Split-Path $DbPath -Leaf)
            report        = $ReportName
            action        = 'created'
            record_source = if ($RecordSource) { $RecordSource } else { '' }
        }
    } finally {
        $script:AccessSession.ControlsCache = @{}
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Get-AccessGroupLevel {
    <#
    .SYNOPSIS
        Read grouping/sorting levels from a report.
    .DESCRIPTION
        Opens the report in design view, iterates over its GroupLevel collection,
        and returns each level's properties (control source, sort order, header/footer, etc.).
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$ReportName,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessGroupLevel'
    if (-not $ReportName) { throw "Get-AccessGroupLevel: -ReportName is required." }

    $app = Connect-AccessDB -DbPath $DbPath

    try {
        Open-InDesignView -ObjectType 'report' -ObjectName $ReportName
        $rpt = $app.Screen.ActiveReport

        $levels = @()
        $index = 0
        # GroupLevel collection — iterate until we get an error (no Count property)
        while ($true) {
            try {
                $gl = $rpt.GroupLevel($index)
                $levels += [ordered]@{
                    index          = $index
                    control_source = $gl.ControlSource
                    sort_order     = if ($gl.SortOrder -eq 0) { 'ascending' } else { 'descending' }
                    group_on       = $gl.GroupOn
                    group_interval = $gl.GroupInterval
                    keep_together  = $gl.KeepTogether
                    group_header   = [bool]$gl.GroupHeader
                    group_footer   = [bool]$gl.GroupFooter
                }
                $index++
            } catch {
                break
            }
        }

        $result = [ordered]@{
            database     = (Split-Path $DbPath -Leaf)
            report       = $ReportName
            group_levels = $levels
            count        = $levels.Count
        }
    } finally {
        try { Save-AndCloseDesign -ObjectType 'report' -ObjectName $ReportName } catch {}
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Set-AccessGroupLevel {
    <#
    .SYNOPSIS
        Add a grouping/sorting level to a report.
    .DESCRIPTION
        Opens the report in design view and calls Application.CreateGroupLevel
        to add a new group/sort level with the specified field expression,
        header/footer visibility, and sort order.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$ReportName,
        [string]$Expression,
        [switch]$GroupHeader,
        [switch]$GroupFooter,
        [ValidateSet('ascending','descending')]
        [string]$SortOrder = 'ascending',
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessGroupLevel'
    if (-not $ReportName) { throw "Set-AccessGroupLevel: -ReportName is required." }
    if (-not $Expression) { throw "Set-AccessGroupLevel: -Expression is required." }

    $app = Connect-AccessDB -DbPath $DbPath

    try {
        Open-InDesignView -ObjectType 'report' -ObjectName $ReportName

        # Application.CreateGroupLevel(reportName, expression, header, footer)
        $headerVal = if ($GroupHeader) { -1 } else { 0 }   # True = -1 in VBA
        $footerVal = if ($GroupFooter) { -1 } else { 0 }
        $levelIndex = $app.Application.CreateGroupLevel($ReportName, $Expression, $headerVal, $footerVal)

        # Set sort order if descending
        $rpt = $app.Screen.ActiveReport
        if ($SortOrder -eq 'descending') {
            $rpt.GroupLevel($levelIndex).SortOrder = 1
        }

        $result = [ordered]@{
            database     = (Split-Path $DbPath -Leaf)
            report       = $ReportName
            action       = 'group_level_added'
            expression   = $Expression
            level_index  = $levelIndex
            group_header = [bool]$GroupHeader
            group_footer = [bool]$GroupFooter
            sort_order   = if ($SortOrder) { $SortOrder } else { 'ascending' }
        }
    } finally {
        try { Save-AndCloseDesign -ObjectType 'report' -ObjectName $ReportName } catch {}
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Remove-AccessGroupLevel {
    <#
    .SYNOPSIS
        Remove a grouping/sorting level from a report by index.
    .DESCRIPTION
        Opens the report in design view, locates the group level at the given
        index, and disables its header and footer sections to effectively
        remove the grouping.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$ReportName,
        [int]$LevelIndex,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Remove-AccessGroupLevel'
    if (-not $ReportName) { throw "Remove-AccessGroupLevel: -ReportName is required." }
    if (-not $PSBoundParameters.ContainsKey('LevelIndex')) { throw "Remove-AccessGroupLevel: -LevelIndex is required." }

    $app = Connect-AccessDB -DbPath $DbPath

    try {
        Open-InDesignView -ObjectType 'report' -ObjectName $ReportName

        $rpt = $app.Screen.ActiveReport
        try {
            $gl = $rpt.GroupLevel($LevelIndex)
            $expression = $gl.ControlSource
            $gl.GroupHeader = 0  # False
            $gl.GroupFooter = 0  # False
        } catch {
            throw "Group level at index $LevelIndex not found on report '$ReportName': $_"
        }

        $result = [ordered]@{
            database    = (Split-Path $DbPath -Leaf)
            report      = $ReportName
            action      = 'group_level_removed'
            level_index = $LevelIndex
            expression  = $expression
        }
    } finally {
        try { Save-AndCloseDesign -ObjectType 'report' -ObjectName $ReportName } catch {}
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}
