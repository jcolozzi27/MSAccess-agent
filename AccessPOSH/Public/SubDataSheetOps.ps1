# Public/SubDataSheetOps.ps1 — Table subdatasheet configuration

function Get-AccessSubDataSheet {
    <#
    .SYNOPSIS
        Read subdatasheet configuration for a table.
    .DESCRIPTION
        Returns the subdatasheet properties (name, link fields, height, expanded)
        for the specified table in an Access database.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$TableName,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessSubDataSheet'
    if (-not $TableName) { throw "Get-AccessSubDataSheet: -TableName is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()

    try {
        $td = $db.TableDefs($TableName)
    } catch {
        throw "Table '$TableName' not found: $_"
    }

    # Read DAO TableDef properties — these may not exist if never set
    $subdatasheet = '[Auto]'
    $linkChild = ''
    $linkMaster = ''
    $height = 0
    $expanded = $false

    try { $subdatasheet = $td.Properties.Item('SubdatasheetName').Value } catch {}
    try { $linkChild = $td.Properties.Item('LinkChildFields').Value } catch {}
    try { $linkMaster = $td.Properties.Item('LinkMasterFields').Value } catch {}
    try { $height = $td.Properties.Item('SubdatasheetHeight').Value } catch {}
    try { $expanded = [bool]$td.Properties.Item('SubdatasheetExpanded').Value } catch {}

    $result = [ordered]@{
        database              = (Split-Path $DbPath -Leaf)
        table                 = $TableName
        subdatasheet_name     = $subdatasheet
        link_child_fields     = $linkChild
        link_master_fields    = $linkMaster
        subdatasheet_height   = $height
        subdatasheet_expanded = $expanded
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Set-AccessSubDataSheet {
    <#
    .SYNOPSIS
        Set or clear subdatasheet configuration for a table.
    .DESCRIPTION
        Configures the subdatasheet properties on the specified table.
        Use -SubDataSheetName '[None]' to remove a subdatasheet.
        DAO properties are created if they do not already exist.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$TableName,
        [string]$SubDataSheetName,
        [string]$LinkChildFields,
        [string]$LinkMasterFields,
        [int]$Height,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessSubDataSheet'
    if (-not $TableName) { throw "Set-AccessSubDataSheet: -TableName is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()

    try {
        $td = $db.TableDefs($TableName)
    } catch {
        throw "Table '$TableName' not found: $_"
    }

    # Helper to set or create a DAO property on a TableDef
    # DAO properties may not exist until explicitly created
    # dbText = 10, dbLong = 4, dbBoolean = 1

    $changes = @()

    if ($PSBoundParameters.ContainsKey('SubDataSheetName')) {
        try {
            $td.Properties.Item('SubdatasheetName').Value = $SubDataSheetName
        } catch {
            # Property doesn't exist, create it  (dbText = 10)
            $prop = $td.CreateProperty('SubdatasheetName', 10, $SubDataSheetName)
            $td.Properties.Append($prop)
        }
        $changes += "subdatasheet_name=$SubDataSheetName"
    }

    if ($PSBoundParameters.ContainsKey('LinkChildFields')) {
        try {
            $td.Properties.Item('LinkChildFields').Value = $LinkChildFields
        } catch {
            $prop = $td.CreateProperty('LinkChildFields', 10, $LinkChildFields)
            $td.Properties.Append($prop)
        }
        $changes += "link_child_fields=$LinkChildFields"
    }

    if ($PSBoundParameters.ContainsKey('LinkMasterFields')) {
        try {
            $td.Properties.Item('LinkMasterFields').Value = $LinkMasterFields
        } catch {
            $prop = $td.CreateProperty('LinkMasterFields', 10, $LinkMasterFields)
            $td.Properties.Append($prop)
        }
        $changes += "link_master_fields=$LinkMasterFields"
    }

    if ($PSBoundParameters.ContainsKey('Height')) {
        try {
            $td.Properties.Item('SubdatasheetHeight').Value = $Height
        } catch {
            # dbLong = 4
            $prop = $td.CreateProperty('SubdatasheetHeight', 4, $Height)
            $td.Properties.Append($prop)
        }
        $changes += "subdatasheet_height=$Height"
    }

    if ($changes.Count -eq 0) {
        throw 'No subdatasheet parameters specified. Use -SubDataSheetName, -LinkChildFields, -LinkMasterFields, or -Height.'
    }

    $result = [ordered]@{
        database = (Split-Path $DbPath -Leaf)
        table    = $TableName
        action   = 'subdatasheet_updated'
        changes  = $changes -join '; '
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}
