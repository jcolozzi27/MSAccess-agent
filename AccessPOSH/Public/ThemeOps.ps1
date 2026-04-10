# Public/ThemeOps.ps1 — Office theme management for forms and reports

function Get-AccessTheme {
    <#
    .SYNOPSIS
        Read the current theme for a form or report.
    .DESCRIPTION
        Opens the specified form or report in Design View and reads its ThemeName,
        ThemeColorScheme, and ThemeFontScheme properties.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$ObjectName,
        [ValidateSet('form','report')][string]$ObjectType = 'form',
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessTheme'
    if (-not $ObjectName) { throw "Get-AccessTheme: -ObjectName is required." }
    $app = Connect-AccessDB -DbPath $DbPath

    try {
        Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
        $obj = if ($ObjectType -eq 'form') { $app.Screen.ActiveForm } else { $app.Screen.ActiveReport }

        $themeName = ''
        $themeColorScheme = ''
        $themeFontScheme = ''

        try { $themeName = $obj.ThemeName } catch {}
        try { $themeColorScheme = $obj.ThemeColorScheme } catch {}
        try { $themeFontScheme = $obj.ThemeFontScheme } catch {}

        $result = [ordered]@{
            database           = (Split-Path $DbPath -Leaf)
            object_name        = $ObjectName
            object_type        = $ObjectType
            theme_name         = $themeName
            theme_color_scheme = $themeColorScheme
            theme_font_scheme  = $themeFontScheme
        }
    } finally {
        try { Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName } catch {}
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Set-AccessTheme {
    <#
    .SYNOPSIS
        Apply a theme to a form or report.
    .DESCRIPTION
        Opens the specified form or report in Design View and sets its ThemeName property,
        then saves and closes the object.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$ObjectName,
        [string]$ThemeName,
        [ValidateSet('form','report')][string]$ObjectType = 'form',
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessTheme'
    if (-not $ObjectName) { throw "Set-AccessTheme: -ObjectName is required." }
    if (-not $ThemeName) { throw "Set-AccessTheme: -ThemeName is required." }
    $app = Connect-AccessDB -DbPath $DbPath

    try {
        Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
        $obj = if ($ObjectType -eq 'form') { $app.Screen.ActiveForm } else { $app.Screen.ActiveReport }

        try {
            $obj.ThemeName = $ThemeName
        } catch {
            throw "Failed to apply theme '$ThemeName' to $ObjectType '$ObjectName': $_"
        }

        $result = [ordered]@{
            database    = (Split-Path $DbPath -Leaf)
            object_name = $ObjectName
            object_type = $ObjectType
            action      = 'theme_applied'
            theme_name  = $ThemeName
        }
    } finally {
        try { Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName } catch {}
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Get-AccessThemeList {
    <#
    .SYNOPSIS
        List available Office themes from the installation directory.
    .DESCRIPTION
        Searches known Office theme directories for .thmx files and combines them
        with a list of well-known built-in theme names.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessThemeList'
    $app = Connect-AccessDB -DbPath $DbPath

    # Office themes are stored in the Document Themes folder
    # Try multiple known locations
    $themePaths = @(
        "$env:APPDATA\Microsoft\Templates\Document Themes"
        "$env:PROGRAMFILES\Microsoft Office\root\Document Themes*"
        "${env:PROGRAMFILES(x86)}\Microsoft Office\root\Document Themes*"
        "C:\Program Files\Microsoft Office\root\Document Themes 16"
    )

    $themes = @()
    $searchedPaths = @()

    foreach ($pathPattern in $themePaths) {
        $resolvedPaths = @(Resolve-Path $pathPattern -ErrorAction SilentlyContinue)
        foreach ($resolvedPath in $resolvedPaths) {
            $searchedPaths += $resolvedPath.Path
            $thmxFiles = Get-ChildItem -Path $resolvedPath.Path -Filter '*.thmx' -ErrorAction SilentlyContinue
            foreach ($f in $thmxFiles) {
                $themes += [ordered]@{
                    name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                    path = $f.FullName
                }
            }
        }
    }

    # Also add well-known built-in theme names that may not have .thmx files
    $builtInNames = @('Office', 'Office 2007-2010', 'Adjacency', 'Angles', 'Apex', 'Apothecary', 'Aspect',
        'Austin', 'Black Tie', 'Civic', 'Clarity', 'Composite', 'Concourse', 'Couture', 'Damask',
        'Elemental', 'Equity', 'Essential', 'Executive', 'Flow', 'Foundry', 'Grid', 'Hardcover',
        'Horizon', 'Median', 'Metro', 'Module', 'Newsprint', 'Opulent', 'Oriel', 'Origin',
        'Paper', 'Perspective', 'Pushpin', 'Slipstream', 'Solstice', 'Technic', 'Thatch',
        'Trek', 'Urban', 'Verve', 'Waveform')

    # Deduplicate by name
    $existingNames = @($themes | ForEach-Object { $_.name })
    foreach ($name in $builtInNames) {
        if ($name -notin $existingNames) {
            $themes += [ordered]@{
                name = $name
                path = '(built-in)'
            }
        }
    }

    $result = [ordered]@{
        database       = (Split-Path $DbPath -Leaf)
        count          = $themes.Count
        searched_paths = $searchedPaths
        themes         = $themes
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}
