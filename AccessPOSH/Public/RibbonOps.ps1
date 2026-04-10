# Public/RibbonOps.ps1 — Custom ribbon XML management via USysRibbons table

function Get-AccessRibbon {
    <#
    .SYNOPSIS
        Read custom ribbon definitions from USysRibbons.
    .DESCRIPTION
        Returns one or all custom ribbon XML entries stored in the USysRibbons system table.
        If the table does not exist, returns an empty list (or throws when a specific name is requested).
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$RibbonName,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessRibbon'
    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()

    # Check if USysRibbons table exists
    $tableExists = $false
    for ($i = 0; $i -lt $db.TableDefs.Count; $i++) {
        if ($db.TableDefs($i).Name -eq 'USysRibbons') { $tableExists = $true; break }
    }

    if (-not $tableExists) {
        if ($RibbonName) {
            throw "USysRibbons table does not exist - no custom ribbons defined"
        }
        $result = [ordered]@{
            database = (Split-Path $DbPath -Leaf)
            count    = 0
            ribbons  = @()
        }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    }

    if ($RibbonName) {
        $rs = $db.OpenRecordset("SELECT ID, RibbonName, RibbonXml FROM USysRibbons WHERE RibbonName='$($RibbonName -replace "'","''")'", 4)
        if ($rs.EOF) {
            $rs.Close()
            throw "Ribbon '$RibbonName' not found in USysRibbons"
        }
        $result = [ordered]@{
            database    = (Split-Path $DbPath -Leaf)
            id          = $rs.Fields('ID').Value
            ribbon_name = $rs.Fields('RibbonName').Value
            ribbon_xml  = $rs.Fields('RibbonXml').Value
        }
        $rs.Close()
    } else {
        $rs = $db.OpenRecordset("SELECT ID, RibbonName, RibbonXml FROM USysRibbons ORDER BY RibbonName", 4)
        $ribbons = @()
        while (-not $rs.EOF) {
            $ribbons += [ordered]@{
                id          = $rs.Fields('ID').Value
                ribbon_name = $rs.Fields('RibbonName').Value
                ribbon_xml  = $rs.Fields('RibbonXml').Value
            }
            $rs.MoveNext()
        }
        $rs.Close()
        $result = [ordered]@{
            database = (Split-Path $DbPath -Leaf)
            count    = $ribbons.Count
            ribbons  = $ribbons
        }
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Set-AccessRibbon {
    <#
    .SYNOPSIS
        Create or update a custom ribbon XML definition.
    .DESCRIPTION
        Inserts or updates a ribbon entry in the USysRibbons system table.
        Creates the table automatically if it does not exist.
        Optionally sets the ribbon as the database default via the CustomRibbonID property.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$RibbonName,
        [string]$RibbonXml,
        [switch]$SetAsDefault,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessRibbon'
    if (-not $RibbonName) { throw "Set-AccessRibbon: -RibbonName is required." }
    if (-not $RibbonXml) { throw "Set-AccessRibbon: -RibbonXml is required." }
    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()

    # Ensure USysRibbons table exists
    $tableExists = $false
    for ($i = 0; $i -lt $db.TableDefs.Count; $i++) {
        if ($db.TableDefs($i).Name -eq 'USysRibbons') { $tableExists = $true; break }
    }

    if (-not $tableExists) {
        $db.Execute("CREATE TABLE USysRibbons (ID AUTOINCREMENT PRIMARY KEY, RibbonName TEXT(255), RibbonXml MEMO)")
    }

    # Check if ribbon already exists
    $rs = $db.OpenRecordset("SELECT ID FROM USysRibbons WHERE RibbonName='$($RibbonName -replace "'","''")'", 4)
    $exists = -not $rs.EOF
    $rs.Close()

    $safeName = $RibbonName -replace "'", "''"
    $safeXml = $RibbonXml -replace "'", "''"

    if ($exists) {
        $db.Execute("UPDATE USysRibbons SET RibbonXml='$safeXml' WHERE RibbonName='$safeName'")
        $action = 'updated'
    } else {
        $db.Execute("INSERT INTO USysRibbons (RibbonName, RibbonXml) VALUES ('$safeName', '$safeXml')")
        $action = 'created'
    }

    # Optionally set as the default ribbon
    if ($SetAsDefault) {
        try {
            $db.Properties.Item('CustomRibbonID').Value = $RibbonName
        } catch {
            $prop = $db.CreateProperty('CustomRibbonID', 10, $RibbonName)  # dbText = 10
            $db.Properties.Append($prop)
        }
    }

    $result = [ordered]@{
        database       = (Split-Path $DbPath -Leaf)
        ribbon_name    = $RibbonName
        action         = $action
        set_as_default = [bool]$SetAsDefault
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Remove-AccessRibbon {
    <#
    .SYNOPSIS
        Remove a custom ribbon from USysRibbons.
    .DESCRIPTION
        Deletes the specified ribbon entry from the USysRibbons table.
        If the ribbon was the database default, clears the CustomRibbonID property.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$RibbonName,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Remove-AccessRibbon'
    if (-not $RibbonName) { throw "Remove-AccessRibbon: -RibbonName is required." }
    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()

    # Check table exists
    $tableExists = $false
    for ($i = 0; $i -lt $db.TableDefs.Count; $i++) {
        if ($db.TableDefs($i).Name -eq 'USysRibbons') { $tableExists = $true; break }
    }
    if (-not $tableExists) {
        throw "USysRibbons table does not exist — no custom ribbons to remove"
    }

    $safeName = $RibbonName -replace "'", "''"

    # Verify it exists
    $rs = $db.OpenRecordset("SELECT ID FROM USysRibbons WHERE RibbonName='$safeName'", 4)
    if ($rs.EOF) {
        $rs.Close()
        throw "Ribbon '$RibbonName' not found in USysRibbons"
    }
    $rs.Close()

    $db.Execute("DELETE FROM USysRibbons WHERE RibbonName='$safeName'")

    # If this was the default ribbon, clear the property
    try {
        $currentDefault = $db.Properties.Item('CustomRibbonID').Value
        if ($currentDefault -eq $RibbonName) {
            $db.Properties.Item('CustomRibbonID').Value = ''
        }
    } catch {}

    $result = [ordered]@{
        database    = (Split-Path $DbPath -Leaf)
        ribbon_name = $RibbonName
        action      = 'removed'
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}
