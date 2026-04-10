# Public/ImportOps.ps1 — Import data from external sources (Excel, CSV, text, XML, other databases)

function Import-AccessFromExcel {
    <#
    .SYNOPSIS
        Import an Excel worksheet into an Access table.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$ExcelPath,
        [string]$TableName,
        [string]$SheetName,
        [switch]$HasFieldNames,
        [ValidateSet('xlsx','xls')]
        [string]$SpreadsheetType = 'xlsx',
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Import-AccessFromExcel'
    if (-not $ExcelPath) { throw "Import-AccessFromExcel: -ExcelPath is required." }
    if (-not $TableName) { throw "Import-AccessFromExcel: -TableName is required." }
    $app = Connect-AccessDB -DbPath $DbPath
    if (-not (Test-Path $ExcelPath)) { throw "Excel file not found: $ExcelPath" }
    $ExcelPath = (Resolve-Path $ExcelPath).Path
    $typeMap = @{ xlsx = 10; xls = 8 }
    $acType = $typeMap[$SpreadsheetType]
    $rangeArg = if ($SheetName) { $SheetName } else { [System.Reflection.Missing]::Value }
    $app.DoCmd.TransferSpreadsheet(
        0,
        $acType,
        $TableName,
        $ExcelPath,
        [bool]$HasFieldNames,
        $rangeArg
    )
    $result = [ordered]@{ action = 'imported'; source = $ExcelPath; table = $TableName; format = $SpreadsheetType }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Import-AccessFromCSV {
    <#
    .SYNOPSIS
        Import a CSV or delimited text file into an Access table.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$FilePath,
        [string]$TableName,
        [switch]$HasFieldNames,
        [string]$SpecificationName,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Import-AccessFromCSV'
    if (-not $FilePath) { throw "Import-AccessFromCSV: -FilePath is required." }
    if (-not $TableName) { throw "Import-AccessFromCSV: -TableName is required." }
    $app = Connect-AccessDB -DbPath $DbPath
    if (-not (Test-Path $FilePath)) { throw "File not found: $FilePath" }
    $FilePath = (Resolve-Path $FilePath).Path
    $specArg = if ($SpecificationName) { $SpecificationName } else { [System.Reflection.Missing]::Value }
    $app.DoCmd.TransferText(
        0,
        $specArg,
        $TableName,
        $FilePath,
        [bool]$HasFieldNames
    )
    $result = [ordered]@{ action = 'imported'; source = $FilePath; table = $TableName; format = 'csv' }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Import-AccessFromXML {
    <#
    .SYNOPSIS
        Import an XML file into the Access database.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$XmlPath,
        [ValidateSet('structureonly','dataonly','structureanddata')]
        [string]$ImportOptions = 'structureanddata',
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Import-AccessFromXML'
    if (-not $XmlPath) { throw "Import-AccessFromXML: -XmlPath is required." }
    $app = Connect-AccessDB -DbPath $DbPath
    if (-not (Test-Path $XmlPath)) { throw "XML file not found: $XmlPath" }
    $XmlPath = (Resolve-Path $XmlPath).Path
    $optMap = @{ structureonly = 0; structureanddata = 1; dataonly = 2 }
    $app.ImportXML($XmlPath, $optMap[$ImportOptions])
    $result = [ordered]@{ action = 'imported'; source = $XmlPath; import_option = $ImportOptions }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Import-AccessFromDatabase {
    <#
    .SYNOPSIS
        Import a table or query from another Access database.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$SourceDbPath,
        [string]$SourceObject,
        [string]$DestinationTable,
        [ValidateSet('table','query')]
        [string]$ObjectType = 'table',
        [switch]$StructureOnly,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Import-AccessFromDatabase'
    if (-not $SourceDbPath) { throw "Import-AccessFromDatabase: -SourceDbPath is required." }
    if (-not $SourceObject) { throw "Import-AccessFromDatabase: -SourceObject is required." }
    $app = Connect-AccessDB -DbPath $DbPath
    if (-not (Test-Path $SourceDbPath)) { throw "Source database not found: $SourceDbPath" }
    $SourceDbPath = (Resolve-Path $SourceDbPath).Path
    $destName = if ($DestinationTable) { $DestinationTable } else { $SourceObject }
    $objectTypeConst = if ($ObjectType -eq 'query') { 1 } else { 0 }
    $dataOnly = if ($StructureOnly) { $true } else { $false }
    $app.DoCmd.TransferDatabase(
        0,
        'Microsoft Access',
        $SourceDbPath,
        $objectTypeConst,
        $SourceObject,
        $destName,
        $dataOnly
    )
    $result = [ordered]@{ action = 'imported'; source_db = $SourceDbPath; source_object = $SourceObject; destination_table = $destName; object_type = $ObjectType; structure_only = $dataOnly }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Export-AccessToExcel {
    <#
    .SYNOPSIS
        Export an Access table or query to an Excel workbook.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$ObjectName,
        [string]$ExcelPath,
        [string]$SheetName,
        [switch]$HasFieldNames,
        [ValidateSet('xlsx','xls')]
        [string]$SpreadsheetType = 'xlsx',
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Export-AccessToExcel'
    if (-not $ObjectName) { throw "Export-AccessToExcel: -ObjectName is required." }
    if (-not $ExcelPath) { throw "Export-AccessToExcel: -ExcelPath is required." }
    $app = Connect-AccessDB -DbPath $DbPath
    $ExcelPath = [System.IO.Path]::GetFullPath($ExcelPath)
    $typeMap = @{ xlsx = 10; xls = 8 }
    $acType = $typeMap[$SpreadsheetType]
    $rangeArg = if ($SheetName) { $SheetName } else { [System.Reflection.Missing]::Value }
    $app.DoCmd.TransferSpreadsheet(
        1,
        $acType,
        $ObjectName,
        $ExcelPath,
        [bool]$HasFieldNames,
        $rangeArg
    )
    $result = [ordered]@{ action = 'exported'; object = $ObjectName; path = $ExcelPath; format = $SpreadsheetType }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}
