# Private/DesignView.ps1 — Form/report design view helpers, control parsing

function Open-InDesignView {
    <#
    .SYNOPSIS
        Open a form or report in Design view (internal helper).
        Uses $script:AccessSession.App directly for COM reliability.
    #>
    param(
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName
    )
    if (-not $ObjectType) { throw "Open-InDesignView: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "Open-InDesignView: -ObjectName is required." }
    try {
        if ($ObjectType -eq 'form') {
            $script:AccessSession.App.DoCmd.OpenForm($ObjectName, $script:AC_DESIGN)
        } else {
            $script:AccessSession.App.DoCmd.OpenReport($ObjectName, $script:AC_DESIGN)
        }
    } catch {
        throw "Cannot open '$ObjectName' in Design view. If it is open in another view, close it first. Error: $_"
    }
}

function Get-DesignObject {
    <#
    .SYNOPSIS
        Return the COM Form/Report object currently open in Design view (internal helper).
        Uses Screen.ActiveForm/ActiveReport — the Forms/Reports collection
        cannot be accessed reliably from dot-sourced functions due to a
        PowerShell COM marshaling issue.
    #>
    param(
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName
    )
    if (-not $ObjectType) { throw "Get-DesignObject: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "Get-DesignObject: -ObjectName is required." }
    $sessionApp = $script:AccessSession.App
    if ($ObjectType -eq 'form') {
        $result = $sessionApp.Screen.ActiveForm
    } else {
        $result = $sessionApp.Screen.ActiveReport
    }
    if ($null -eq $result -or $result.Name -ne $ObjectName) {
        throw "Cannot get '$ObjectName' ($ObjectType) - is it open in Design view?"
    }
    $result
}

function Save-AndCloseDesign {
    <#
    .SYNOPSIS
        Save and close a form/report open in Design view, invalidate caches (internal helper).
        Uses $script:AccessSession.App directly for COM reliability.
    #>
    param(
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName
    )
    if (-not $ObjectType) { throw "Save-AndCloseDesign: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "Save-AndCloseDesign: -ObjectName is required." }
    $acType = if ($ObjectType -eq 'form') { $script:AC_TYPE['form'] } else { $script:AC_TYPE['report'] }
    try {
        $script:AccessSession.App.DoCmd.Close($acType, $ObjectName, $script:AC_SAVE_YES)
    } catch {
        Write-Warning "Error closing '$ObjectName': $_"
    }
    # Invalidate caches
    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.ControlsCache.Remove($cacheKey) | Out-Null
    $script:AccessSession.CmCache.Remove($cacheKey) | Out-Null
    $script:AccessSession.VbeCodeCache.Remove($cacheKey) | Out-Null
}

function ConvertFrom-ControlBlock {
    <#
    .SYNOPSIS
        Parse the SaveAsText export of a form/report and extract control blocks (internal helper).
    .DESCRIPTION
        Returns a hashtable with:
          controls       — array of controls with properties and line positions
          form_indent    — indentation of the Begin Form/Report line
          ctrl_indent    — indent of the first control found (legacy compat)
          form_begin_idx — 0-based line index of Begin Form/Report
          form_end_idx   — 0-based line index of the closing End
    #>
    param(
        [string]$FormText
    )
    if (-not $FormText) { throw "ConvertFrom-ControlBlock: -FormText is required." }

    $lines = $FormText -split "`r?`n"
    $result = [ordered]@{
        controls       = @()
        form_indent    = ''
        ctrl_indent    = ''
        form_begin_idx = -1
        form_end_idx   = -1
    }

    # Known control type names for fast lookup
    $ctrlTypeNames = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]($script:CTRL_TYPE.Values),
        [System.StringComparer]::Ordinal
    )

    # 1. Locate "Begin Form" or "Begin Report"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $s = $lines[$i].TrimStart()
        if ($s -match '^Begin\s+(Form|Report)\s*$') {
            $raw = $lines[$i]
            $result['form_indent'] = $raw.Substring(0, $raw.Length - $raw.TrimStart().Length)
            $result['form_begin_idx'] = $i
            break
        }
    }
    if ($result['form_begin_idx'] -eq -1) { return $result }

    $formBegin = $result['form_begin_idx']

    # 2. Find the matching "End" (depth tracking, including "Property = Begin" blocks)
    $depth = 0
    for ($i = $formBegin; $i -lt $lines.Count; $i++) {
        $s = $lines[$i].TrimStart()
        if ($s -match '^Begin\b' -or $s -match '^\w+\s*=\s*Begin\s*$') {
            $depth++
        } elseif ($s -eq 'End') {
            $depth--
            if ($depth -eq 0) {
                $result['form_end_idx'] = $i
                break
            }
        }
    }
    if ($result['form_end_idx'] -eq -1) { return $result }

    # Container types whose children get a "parent" field
    $containerTypes = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@('Page', 'OptionGroup'),
        [System.StringComparer]::Ordinal
    )

    # 3. Scan all "Begin <TypeName>" blocks where TypeName is a known control type
    $controls = [System.Collections.Generic.List[object]]::new()
    $containerStack = [System.Collections.Generic.List[object]]::new()  # @{name; end_idx}
    $currentSection = ''
    $i = $formBegin + 1
    while ($i -lt $result['form_end_idx']) {
        # Clean up containers we've passed
        while ($containerStack.Count -gt 0 -and $i -gt $containerStack[$containerStack.Count - 1].end_idx) {
            $containerStack.RemoveAt($containerStack.Count - 1)
        }

        $raw = $lines[$i]
        $s = $raw.TrimStart()
        $indent = $raw.Substring(0, $raw.Length - $s.Length)

        # Skip ClassModule — contains VBA, not controls
        if ($s -match '^Begin\s+ClassModule\s*$') { break }

        # Track current section (Detail, FormHeader, FormFooter, PageHeader, PageFooter, etc.)
        if ($s -match '^Begin\s+Section\s*$') {
            # Look ahead for the section Name property
            $secName = ''
            $secDepth = 1
            for ($si = $i + 1; $si -lt $lines.Count; $si++) {
                $secLine = $lines[$si].TrimStart()
                if ($secLine -match '^Begin\b') { $secDepth++ }
                elseif ($secLine -eq 'End') {
                    $secDepth--
                    if ($secDepth -eq 0) { break }
                }
                if ($secDepth -eq 1) {
                    $secMatch = [regex]::Match($secLine, '^Name\s*=\s*"?([^"]*)"?\s*$')
                    if ($secMatch.Success) { $secName = $secMatch.Groups[1].Value; break }
                }
            }
            $currentSection = $secName
        }

        # Detect "Begin <TypeName>"
        $mCtrl = [regex]::Match($s, '^Begin\s+(\w+)\s*$')
        if ($mCtrl.Success -and $ctrlTypeNames.Contains($mCtrl.Groups[1].Value)) {
            $ctrlStart = $i
            $block = [System.Collections.Generic.List[string]]::new()
            $block.Add($lines[$i])
            $props = @{}
            $blkDepth = 1
            $ctrlEnd = $i
            $j = $i + 1
            while ($j -lt $lines.Count) {
                $bl = $lines[$j]
                $blS = $bl.TrimStart()
                $block.Add($bl)
                # Parse top-level properties only (depth == 1)
                if ($blkDepth -eq 1) {
                    $mProp = [regex]::Match($blS, '^(\w+)\s*=(.*)')
                    if ($mProp.Success) {
                        $props[$mProp.Groups[1].Value] = $mProp.Groups[2].Value.Trim().Trim('"')
                    }
                }
                if ($blS -match '^Begin\b') {
                    $blkDepth++
                } elseif ($blS -eq 'End') {
                    $blkDepth--
                    if ($blkDepth -eq 0) {
                        $ctrlEnd = $j
                        break
                    }
                }
                $j++
            }

            $name = if ($props.ContainsKey('Name')) { $props['Name'] } elseif ($props.ContainsKey('ControlName')) { $props['ControlName'] } else { '' }
            $ctype = -1
            if ($props.ContainsKey('ControlType')) {
                $parsed = 0
                if ([int]::TryParse($props['ControlType'], [ref]$parsed)) { $ctype = $parsed }
            }

            $rawText = $block -join "`r`n"
            $fmtCount = ($block | Where-Object { $_ -match '^\s+ConditionalFormat\d*\s*=\s*Begin\s*$' }).Count

            if (-not $result['ctrl_indent'] -and $name) {
                $result['ctrl_indent'] = $indent
            }

            $ctrlEntry = [ordered]@{
                name           = $name
                control_type   = $ctype
                type_name      = if ($script:CTRL_TYPE.ContainsKey($ctype)) { $script:CTRL_TYPE[$ctype] } else { $mCtrl.Groups[1].Value }
                caption        = if ($props.ContainsKey('Caption')) { $props['Caption'] } else { '' }
                control_source = if ($props.ContainsKey('ControlSource')) { $props['ControlSource'] } else { '' }
                left           = if ($props.ContainsKey('Left'))   { $props['Left'] }   else { '' }
                top            = if ($props.ContainsKey('Top'))    { $props['Top'] }    else { '' }
                width          = if ($props.ContainsKey('Width'))  { $props['Width'] }  else { '' }
                height         = if ($props.ContainsKey('Height')) { $props['Height'] } else { '' }
                visible        = if ($props.ContainsKey('Visible')){ $props['Visible'] } else { '' }
                section        = $currentSection
                parent         = if ($containerStack.Count -gt 0) { $containerStack[$containerStack.Count - 1].name } else { '' }
                start_line     = $ctrlStart + 1   # 1-based
                end_line       = $ctrlEnd + 1     # 1-based inclusive
                raw_block      = $rawText
            }
            if ($fmtCount -gt 0) { $ctrlEntry['format_conditions'] = $fmtCount }
            $controls.Add([PSCustomObject]$ctrlEntry)

            # Container types: re-scan inside instead of skipping past
            if ($containerTypes.Contains($mCtrl.Groups[1].Value)) {
                $containerStack.Add([PSCustomObject]@{ name = $name; end_idx = $ctrlEnd })
                $i = $ctrlStart + 1  # re-scan inside the container
            } else {
                $i = $ctrlEnd + 1
            }
            continue
        }
        $i++
    }

    $result['controls'] = @($controls)
    return $result
}

function Get-ParsedControls {
    <#
    .SYNOPSIS
        Return parsed controls for a form/report, using the ControlsCache (internal helper).
    #>
    param(
        [string]$DbPath,
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-ParsedControls'
    if (-not $ObjectType) { throw "Get-ParsedControls: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "Get-ParsedControls: -ObjectName is required." }
    $cacheKey = "${ObjectType}:${ObjectName}"
    if (-not $script:AccessSession.ControlsCache.ContainsKey($cacheKey)) {
        $text = Get-AccessCode -DbPath $DbPath -ObjectType $ObjectType -Name $ObjectName
        $script:AccessSession.ControlsCache[$cacheKey] = ConvertFrom-ControlBlock -FormText $text
    }
    return $script:AccessSession.ControlsCache[$cacheKey]
}
