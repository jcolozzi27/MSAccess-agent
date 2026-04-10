# Public/TempVarOps.ps1 — TempVars collection management

function Get-AccessTempVar {
    <#
    .SYNOPSIS
        Get one or all TempVars from an Access database.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$Name,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessTempVar'
    $app = Connect-AccessDB -DbPath $DbPath

    if ($Name) {
        try {
            $val = $app.TempVars.Item($Name).Value
        } catch {
            throw "TempVar '$Name' not found"
        }
        $result = [ordered]@{ name = $Name; value = $val }
    } else {
        $tvars = $app.TempVars
        $items = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $tvars.Count; $i++) {
            $tv = $tvars.Item($i)
            $items.Add([ordered]@{ name = $tv.Name; value = $tv.Value })
        }
        $result = [ordered]@{ count = $items.Count; tempvars = @($items) }
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Set-AccessTempVar {
    <#
    .SYNOPSIS
        Create or update a TempVar in an Access database.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$Name,
        [object]$Value,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessTempVar'
    if (-not $Name) { throw "Set-AccessTempVar: -Name is required." }
    if (-not $PSBoundParameters.ContainsKey('Value')) { throw "Set-AccessTempVar: -Value is required." }
    $app = Connect-AccessDB -DbPath $DbPath

    $app.TempVars.Add($Name, $Value)
    $result = [ordered]@{ name = $Name; value = $Value; action = 'set' }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Remove-AccessTempVar {
    <#
    .SYNOPSIS
        Remove one or all TempVars from an Access database.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$Name,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Remove-AccessTempVar'
    $app = Connect-AccessDB -DbPath $DbPath

    if ($Name) {
        $app.TempVars.Remove($Name)
        $result = [ordered]@{ name = $Name; action = 'removed' }
    } else {
        $count = $app.TempVars.Count
        $app.TempVars.RemoveAll()
        $result = [ordered]@{ action = 'removed_all'; count = $count }
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}
