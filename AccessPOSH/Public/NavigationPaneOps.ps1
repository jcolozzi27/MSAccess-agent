# Public/NavigationPaneOps.ps1 — Navigation Pane visibility and lock management

function Show-AccessNavigationPane {
    <#
    .SYNOPSIS
        Show the navigation pane in the Access database.
    .DESCRIPTION
        Makes the Navigation Pane visible using DoCmd.NavigateTo, falling back
        to the CommandBars approach if that fails.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Show-AccessNavigationPane'
    $app = Connect-AccessDB -DbPath $DbPath
    try {
        $app.DoCmd.NavigateTo('acNavigationCategoryObjectType')
    } catch {
        # Fallback: try CommandBars approach
        try {
            $app.CommandBars.Item('Navigation Pane').Enabled = $true
        } catch {
            throw "Failed to show Navigation Pane: $_"
        }
    }
    $result = [ordered]@{
        database = (Split-Path $DbPath -Leaf)
        action   = 'navigation_pane_shown'
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Hide-AccessNavigationPane {
    <#
    .SYNOPSIS
        Hide the navigation pane in the Access database.
    .DESCRIPTION
        Hides the Navigation Pane using DoCmd.SelectObject and RunCommand,
        falling back to the CommandBars approach if that fails.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Hide-AccessNavigationPane'
    $app = Connect-AccessDB -DbPath $DbPath
    try {
        $app.DoCmd.SelectObject(2, '', $true)   # acForm, empty, InNavPane=True selects NavPane
        $app.DoCmd.RunCommand(6)                 # acCmdWindowHide — hides the NavPane
    } catch {
        try {
            $app.CommandBars.Item('Navigation Pane').Enabled = $false
        } catch {
            throw "Failed to hide Navigation Pane: $_"
        }
    }
    $result = [ordered]@{
        database = (Split-Path $DbPath -Leaf)
        action   = 'navigation_pane_hidden'
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Set-AccessNavigationPaneLock {
    <#
    .SYNOPSIS
        Lock or unlock the navigation pane in the Access database.
    .DESCRIPTION
        Sets the 'NavigationPane Locked' database property to prevent or allow
        users from rearranging the Navigation Pane. Creates the property if it
        does not already exist.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [bool]$Locked,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessNavigationPaneLock'
    if (-not $PSBoundParameters.ContainsKey('Locked')) { throw "Set-AccessNavigationPaneLock: -Locked is required." }
    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()

    $propName = 'NavigationPane Locked'
    $lockValue = if ($Locked) { $true } else { $false }

    try {
        $db.Properties.Item($propName).Value = $lockValue
    } catch {
        # Property may not exist yet — create it (dbBoolean = 1)
        try {
            $prop = $db.CreateProperty($propName, 1, $lockValue)
            $db.Properties.Append($prop)
        } catch {
            throw "Failed to set Navigation Pane lock: $_"
        }
    }

    $result = [ordered]@{
        database = (Split-Path $DbPath -Leaf)
        action   = if ($Locked) { 'navigation_pane_locked' } else { 'navigation_pane_unlocked' }
        locked   = $lockValue
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}
