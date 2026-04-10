# Public/VbeOps.ps1 — VBE/VBA operations: read, write, search, compile, execute

function Get-AccessVbeLine {
    <#
    .SYNOPSIS
        Read a range of lines from a VBE CodeModule.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER StartLine
        First line to read (1-based).
    .PARAMETER Count
        Number of lines to read.
    .EXAMPLE
        Get-AccessVbeLine -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -StartLine 1 -Count 10
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('module','form','report')][string]$ObjectType,
        [string]$ObjectName,
        [int]$StartLine,
        [int]$Count
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessVbeLine'
    if (-not $ObjectType) { throw "Get-AccessVbeLine: -ObjectType is required (module, form, report)." }
    if (-not $ObjectName) { throw "Get-AccessVbeLine: -ObjectName is required." }
    if (-not $PSBoundParameters.ContainsKey('StartLine')) { throw "Get-AccessVbeLine: -StartLine is required." }
    if (-not $PSBoundParameters.ContainsKey('Count')) { throw "Get-AccessVbeLine: -Count is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $cacheKey = "${ObjectType}:${ObjectName}"
    $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey
    $allLines = $allCode.Split("`n")
    $total = $allLines.Count

    if ($StartLine -lt 1 -or $StartLine -gt $total) {
        throw "start_line $StartLine out of range (1-$total)"
    }
    $actual = [math]::Min($Count, $total - $StartLine + 1)
    $result = $allLines[($StartLine - 1) .. ($StartLine - 1 + $actual - 1)] -join "`n"
    return $result.TrimEnd("`r")
}

function Get-AccessVbeProc {
    <#
    .SYNOPSIS
        Extract a procedure by name from a VBE module.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER ProcName
        Name of the procedure (Sub/Function/Property).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessVbeProc -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -ProcName "MyFunc"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('module','form','report')][string]$ObjectType,
        [string]$ObjectName,
        [string]$ProcName,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessVbeProc'
    if (-not $ObjectType) { throw "Get-AccessVbeProc: -ObjectType is required (module, form, report)." }
    if (-not $ObjectName) { throw "Get-AccessVbeProc: -ObjectName is required." }
    if (-not $ProcName) { throw "Get-AccessVbeProc: -ProcName is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName

    try {
        $start = $cm.ProcStartLine($ProcName, 0)  # 0 = vbext_pk_Proc
        $body  = $cm.ProcBodyLine($ProcName, 0)
        $count = $cm.ProcCountLines($ProcName, 0)
    } catch {
        throw "Procedure '$ProcName' not found in '$ObjectName': $_"
    }

    $cacheKey = "${ObjectType}:${ObjectName}"
    $allLines = (Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey).Split("`n")
    $total = $allLines.Count
    $count = [math]::Min($count, $total - $start + 1)
    $code = ($allLines[($start - 1) .. ($start - 1 + $count - 1)] -join "`n").TrimEnd("`r")

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        proc_name  = $ProcName
        start_line = $start
        body_line  = $body
        count      = $count
        code       = $code
    })
}

function Get-AccessVbeModuleInfo {
    <#
    .SYNOPSIS
        Enumerate all procedures in a VBE module with their positions.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessVbeModuleInfo -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('module','form','report')][string]$ObjectType,
        [string]$ObjectName,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessVbeModuleInfo'
    if (-not $ObjectType) { throw "Get-AccessVbeModuleInfo: -ObjectType is required (module, form, report)." }
    if (-not $ObjectName) { throw "Get-AccessVbeModuleInfo: -ObjectName is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $cacheKey = "${ObjectType}:${ObjectName}"
    $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey
    $allLines = $allCode.Split("`n")
    $total = $allLines.Count

    $procs = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    for ($i = 0; $i -lt $total; $i++) {
        $line = $allLines[$i].Trim()
        if ($line -match '^(?:Public\s+|Private\s+|Friend\s+)?(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+(\w+)') {
            $pname = $Matches[1]
            if (-not $seen.Add($pname)) { continue }
            try {
                $pstart = $cm.ProcStartLine($pname, 0)
                $pbody  = $cm.ProcBodyLine($pname, 0)
                $pcount = $cm.ProcCountLines($pname, 0)
                $pcount = [math]::Min($pcount, $total - $pstart + 1)
                $procs.Add([PSCustomObject][ordered]@{
                    name       = $pname
                    start_line = $pstart
                    body_line  = $pbody
                    count      = $pcount
                })
            } catch {
                $procs.Add([PSCustomObject][ordered]@{
                    name       = $pname
                    start_line = ($i + 1)
                })
            }
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        total_lines = $total
        procs       = @($procs)
    })
}

function Set-AccessVbeLine {
    <#
    .SYNOPSIS
        Replace lines in a VBE module. count=0 inserts without deleting. Empty NewCode deletes only.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER StartLine
        First line to replace (1-based).
    .PARAMETER Count
        Number of lines to delete before inserting.
    .PARAMETER NewCode
        Code to insert at StartLine (can be multiline).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessVbeLine -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -StartLine 5 -Count 3 -NewCode "' replaced lines"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('module','form','report')][string]$ObjectType,
        [string]$ObjectName,
        [int]$StartLine,
        [int]$Count = 0,
        [string]$NewCode = '',
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessVbeLine'
    if (-not $ObjectType) { throw "Set-AccessVbeLine: -ObjectType is required (module, form, report)." }
    if (-not $ObjectName) { throw "Set-AccessVbeLine: -ObjectName is required." }
    if (-not $PSBoundParameters.ContainsKey('StartLine')) { throw "Set-AccessVbeLine: -StartLine is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $total = $cm.CountOfLines

    if ($StartLine -lt 1 -or $StartLine -gt ($total + 1)) {
        throw "start_line $StartLine out of range (1-$total)"
    }

    $clamped = $false
    if ($Count -gt 0) {
        $maxCount = $total - $StartLine + 1
        if ($Count -gt $maxCount) { $Count = $maxCount; $clamped = $true }
        $cm.DeleteLines($StartLine, $Count)
    }

    $inserted = 0
    if ($NewCode) {
        $normalized = $NewCode -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
        $cm.InsertLines($StartLine, $normalized)
        $inserted = $NewCode.Split("`n").Count
    }

    # Invalidate cache
    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines
    $clampNote = if ($clamped) { ' (count clamped to module boundary)' } else { '' }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status        = "Lines $StartLine replaced ($Count deleted, $inserted inserted)$clampNote -> module now has $newTotal lines"
        deleted       = $Count
        inserted      = $inserted
        new_total     = $newTotal
    })
}

function Set-AccessVbeProc {
    <#
    .SYNOPSIS
        Replace an entire procedure by name. Auto-locates via ProcStartLine/ProcCountLines.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER ProcName
        Name of the procedure to replace.
    .PARAMETER NewCode
        New code for the procedure. Empty string deletes it.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessVbeProc -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -ProcName "OldFunc" -NewCode "Public Sub OldFunc()`r`n  MsgBox ""Hello""`r`nEnd Sub"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('module','form','report')][string]$ObjectType,
        [string]$ObjectName,
        [string]$ProcName,
        [string]$NewCode = '',
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessVbeProc'
    if (-not $ObjectType) { throw "Set-AccessVbeProc: -ObjectType is required (module, form, report)." }
    if (-not $ObjectName) { throw "Set-AccessVbeProc: -ObjectName is required." }
    if (-not $ProcName) { throw "Set-AccessVbeProc: -ProcName is required." }

    $app = Connect-AccessDB -DbPath $DbPath

    # Close form/report if open in design (avoids COM conflicts with VBE)
    if ($ObjectType -in 'form', 'report') {
        $acObjType = if ($ObjectType -eq 'form') { $script:AC_FORM } else { $script:AC_REPORT }
        try { $app.DoCmd.Close($acObjType, $ObjectName, $script:AC_SAVE_YES) } catch {}
    }

    # Invalidate cache in case CodeModule is stale
    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.CmCache.Remove($cacheKey)

    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    try {
        $start = $cm.ProcStartLine($ProcName, 0)
        $count = $cm.ProcCountLines($ProcName, 0)
    } catch {
        throw "Procedure '$ProcName' not found in '$ObjectName': $_"
    }

    $total = $cm.CountOfLines
    $count = [math]::Min($count, $total - $start + 1)

    $cm.DeleteLines($start, $count)

    $inserted = 0
    if ($NewCode) {
        $normalized = $NewCode -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
        $cm.InsertLines($start, $normalized)
        $inserted = $NewCode.Split("`n").Count
    }

    # Invalidate caches
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)
    $script:AccessSession.CmCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines
    $action = if ($NewCode) { 'replaced' } else { 'deleted' }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status   = "Proc '$ProcName' $action ($count deleted, $inserted inserted) -> module now has $newTotal lines"
        action   = $action
        proc     = $ProcName
        deleted  = $count
        inserted = $inserted
        new_total = $newTotal
    })
}

function Update-AccessVbeProc {
    <#
    .SYNOPSIS
        Surgically patch code within a procedure using find/replace with whitespace tolerance.
    .DESCRIPTION
        Applies one or more patches (find/replace pairs) to a procedure without rewriting the entire proc.
        Three-layer matching: (1) exact string, (2) whitespace-normalized fallback, (3) context error reporting.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER ProcName
        Name of the procedure to patch.
    .PARAMETER Patches
        Array of hashtables, each with 'find' and 'replace' keys.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Update-AccessVbeProc -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" `
            -ProcName "MyFunc" -Patches @(@{find='MsgBox "Old"'; replace='MsgBox "New"'})
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('module','form','report')][string]$ObjectType,
        [string]$ObjectName,
        [string]$ProcName,
        [array]$Patches,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Update-AccessVbeProc'
    if (-not $ObjectType) { throw "Update-AccessVbeProc: -ObjectType is required (module, form, report)." }
    if (-not $ObjectName) { throw "Update-AccessVbeProc: -ObjectName is required." }
    if (-not $ProcName) { throw "Update-AccessVbeProc: -ProcName is required." }
    if (-not $Patches) { throw "Update-AccessVbeProc: -Patches is required." }

    $app = Connect-AccessDB -DbPath $DbPath

    # Close form/report if open in design (avoids COM conflicts with VBE)
    if ($ObjectType -in 'form', 'report') {
        $acObjType = if ($ObjectType -eq 'form') { $script:AC_FORM } else { $script:AC_REPORT }
        try { $app.DoCmd.Close($acObjType, $ObjectName, $script:AC_SAVE_YES) } catch {}
    }

    # Invalidate cache
    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.CmCache.Remove($cacheKey)

    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName

    # Locate procedure — try kind=0 (Sub/Function), fallback to kind=3 (Property)
    $kind = 0
    try {
        $start = $cm.ProcStartLine($ProcName, 0)
        $count = $cm.ProcCountLines($ProcName, 0)
    } catch {
        try {
            $start = $cm.ProcStartLine($ProcName, 3)
            $count = $cm.ProcCountLines($ProcName, 3)
            $kind = 3
        } catch {
            throw "Procedure '$ProcName' not found in '$ObjectName': $_"
        }
    }

    $total = $cm.CountOfLines
    $count = [math]::Min($count, $total - $start + 1)

    # Get current proc code
    $procCode = $cm.Lines($start, $count)
    $backupCode = $procCode

    # Apply patches sequentially
    $applied = 0
    $notFound = @()
    $wsFallbackNotes = @()

    for ($pi = 0; $pi -lt $Patches.Count; $pi++) {
        $patch = $Patches[$pi]
        $findText    = "$($patch['find'])"
        $replaceText = if ($patch.ContainsKey('replace')) { "$($patch['replace'])" } else { '' }

        # Layer 1: Exact match
        if ($procCode.Contains($findText)) {
            $idx = $procCode.IndexOf($findText)
            $procCode = $procCode.Substring(0, $idx) + $replaceText + $procCode.Substring($idx + $findText.Length)
            $applied++
        }
        else {
            # Layer 2: Whitespace-normalized fallback
            $wsMatch = Test-WsNormalizedMatch -ProcCode $procCode -FindText $findText
            if ($null -ne $wsMatch) {
                $codeLines = @($procCode -split "`r?`n")
                $replaceNorm = $replaceText
                if ($replaceNorm -and -not $replaceNorm.EndsWith("`n")) {
                    $replaceNorm += "`r`n"
                }
                $before = if ($wsMatch.start -gt 0) { $codeLines[0..($wsMatch.start - 1)] } else { @() }
                $after  = if ($wsMatch.end -lt ($codeLines.Count - 1)) { $codeLines[($wsMatch.end + 1)..($codeLines.Count - 1)] } else { @() }
                $middle = if ($replaceNorm) { @($replaceNorm.TrimEnd("`r`n")) } else { @() }
                $procCode = ($before + $middle + $after) -join "`r`n"
                $applied++
                $wsFallbackNotes += "patch[$pi]: matched via ws-normalized fallback"
            }
            else {
                # Layer 3: Context error reporting
                $ctx = Get-ClosestMatchContext -ProcCode $procCode -FindText $findText -ProcName $ProcName
                $notFound += "patch[$pi]: not found. $ctx"
            }
        }
    }

    if ($applied -eq 0) {
        $errMsg = "NOOP: no patches matched in '$ProcName'. Errors:`n$($notFound -join "`n")"
        return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            status     = $errMsg
            applied    = 0
            total      = $Patches.Count
            not_found  = $notFound
        })
    }

    # Strip Option lines if proc is NOT at the top of the module
    $optionWarnings = @()
    if ($start -gt 5) {
        $optionRe = '^\s*Option\s+(Explicit|Compare\s+\w+)\s*$'
        $cleanLines = @()
        foreach ($line in ($procCode -split "`r?`n")) {
            if ($line -match $optionRe) {
                $optionWarnings += "Stripped misplaced Option line: '$($line.Trim())'"
            } else {
                $cleanLines += $line
            }
        }
        $procCode = $cleanLines -join "`r`n"
    }

    # Replace entire proc with patched code
    try {
        $cm.DeleteLines($start, $count)
        if ($procCode.Trim()) {
            $normalized = $procCode -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
            $cm.InsertLines($start, $normalized)
        }
    } catch {
        # Rollback: restore backup
        try { $cm.InsertLines($start, $backupCode) } catch {}
        throw "Patch failed (rolled back): $_"
    }

    # Invalidate caches
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)
    $script:AccessSession.CmCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines
    $newCount = try { $cm.ProcCountLines($ProcName, $kind) } catch { 0 }

    # Build result
    $resultParts = @("OK: $applied/$($Patches.Count) patches applied in '$ProcName' ($count -> $newCount lines) -> module now has $newTotal lines")
    if ($wsFallbackNotes.Count -gt 0) { $resultParts += "WS-fallback: $($wsFallbackNotes -join '; ')" }
    if ($optionWarnings.Count -gt 0)  { $resultParts += $optionWarnings }
    if ($notFound.Count -gt 0)        { $resultParts += "Not found:"; $resultParts += $notFound }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status          = $resultParts -join "`n"
        applied         = $applied
        total           = $Patches.Count
        proc            = $ProcName
        old_lines       = $count
        new_lines       = $newCount
        module_lines    = $newTotal
        ws_fallback     = $wsFallbackNotes
        not_found       = $notFound
    })
}

function Add-AccessVbeCode {
    <#
    .SYNOPSIS
        Append code at the end of a VBE module.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER Code
        Code to append (can be multiline).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-AccessVbeCode -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -Code "Public Sub NewSub()`r`nEnd Sub"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('module','form','report')][string]$ObjectType,
        [string]$ObjectName,
        [string]$Code,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Add-AccessVbeCode'
    if (-not $ObjectType) { throw "Add-AccessVbeCode: -ObjectType is required (module, form, report)." }
    if (-not $ObjectName) { throw "Add-AccessVbeCode: -ObjectName is required." }
    if (-not $Code) { throw "Add-AccessVbeCode: -Code is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $total = $cm.CountOfLines

    $normalized = $Code -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
    $cm.InsertLines($total + 1, $normalized)
    $inserted = $Code.Split("`n").Count

    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status    = "$inserted lines appended -> module now has $newTotal lines"
        inserted  = $inserted
        new_total = $newTotal
    })
}

function Find-AccessVbeText {
    <#
    .SYNOPSIS
        Search for text (or regex) in a VBE module.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Find-AccessVbeText -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -SearchText "MsgBox"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('module','form','report')][string]$ObjectType,
        [string]$ObjectName,
        [string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Find-AccessVbeText'
    if (-not $ObjectType) { throw "Find-AccessVbeText: -ObjectType is required (module, form, report)." }
    if (-not $ObjectName) { throw "Find-AccessVbeText: -ObjectName is required." }
    if (-not $SearchText) { throw "Find-AccessVbeText: -SearchText is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $cacheKey = "${ObjectType}:${ObjectName}"
    $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey

    if (-not $allCode) {
        return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{ found = $false; match_count = 0; matches = @() })
    }

    $matchList = [System.Collections.Generic.List[object]]::new()
    $lineNum = 0
    foreach ($line in $allCode.Split("`n")) {
        $lineNum++
        if (Test-TextMatch -Needle $SearchText -Haystack $line -MatchCase:$MatchCase -UseRegex:$UseRegex) {
            $matchList.Add([PSCustomObject][ordered]@{
                line    = $lineNum
                content = $line.TrimEnd("`r")
            })
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        found       = ($matchList.Count -gt 0)
        match_count = $matchList.Count
        matches     = @($matchList)
    })
}

function Search-AccessVbe {
    <#
    .SYNOPSIS
        Search for text (or regex) across ALL VBA modules, forms, and reports.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER MaxResults
        Maximum total matches to return (default 100).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Search-AccessVbe -DbPath "C:\db.accdb" -SearchText "DoCmd"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [int]$MaxResults = 100,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Search-AccessVbe'
    if (-not $SearchText) { throw "Search-AccessVbe: -SearchText is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $objects = Get-AccessObject -DbPath $DbPath -ObjectType all
    $results = [System.Collections.Generic.List[object]]::new()
    $total = 0
    $truncated = $false

    foreach ($objType in @('module', 'form', 'report')) {
        if ($truncated) { break }
        $names = @($objects.$objType)
        foreach ($objName in $names) {
            if ($truncated) { break }
            try {
                $cm = Get-CodeModule -App $app -ObjectType $objType -ObjectName $objName
                $cacheKey = "${objType}:${objName}"
                $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey
                if (-not $allCode) { continue }

                $objMatches = [System.Collections.Generic.List[object]]::new()
                $lineNum = 0
                foreach ($line in $allCode.Split("`n")) {
                    $lineNum++
                    if (Test-TextMatch -Needle $SearchText -Haystack $line -MatchCase:$MatchCase -UseRegex:$UseRegex) {
                        $objMatches.Add([PSCustomObject][ordered]@{
                            line    = $lineNum
                            content = $line.TrimEnd("`r")
                        })
                        $total++
                        if ($total -ge $MaxResults) { $truncated = $true; break }
                    }
                }
                if ($objMatches.Count -gt 0) {
                    $results.Add([PSCustomObject][ordered]@{
                        object_type = $objType
                        object_name = $objName
                        matches     = @($objMatches)
                    })
                }
            } catch { continue }
        }
    }

    $out = [ordered]@{ total_matches = $total; results = @($results) }
    if ($truncated) { $out['truncated'] = $true }
    Format-AccessOutput -AsJson:$AsJson -Data $out
}

function Search-AccessQuery {
    <#
    .SYNOPSIS
        Search for text (or regex) in the SQL of all queries.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER MaxResults
        Maximum results to return (default 100).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Search-AccessQuery -DbPath "C:\db.accdb" -SearchText "Users"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [int]$MaxResults = 100,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Search-AccessQuery'
    if (-not $SearchText) { throw "Search-AccessQuery: -SearchText is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()
    $results = [System.Collections.Generic.List[object]]::new()
    $total = 0

    foreach ($qd in $db.QueryDefs) {
        $name = $qd.Name
        if ($name.StartsWith('~')) { continue }
        $sql = $qd.SQL
        if (Test-TextMatch -Needle $SearchText -Haystack $sql -MatchCase:$MatchCase -UseRegex:$UseRegex) {
            $results.Add([PSCustomObject][ordered]@{
                query_name = $name
                sql        = $sql.Trim()
            })
            $total++
            if ($total -ge $MaxResults) { break }
        }
    }

    $out = [ordered]@{ total_matches = $total; results = @($results) }
    if ($total -ge $MaxResults) { $out['truncated'] = $true }
    Format-AccessOutput -AsJson:$AsJson -Data $out
}

function Find-AccessUsage {
    <#
    .SYNOPSIS
        Search for a name across VBA code, query SQL, and form/report control properties.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER MaxResults
        Maximum total matches (default 200).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Find-AccessUsage -DbPath "C:\db.accdb" -SearchText "Users"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [int]$MaxResults = 200,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Find-AccessUsage'
    if (-not $SearchText) { throw "Find-AccessUsage: -SearchText is required." }

    # 1. VBA matches
    $vbaResult = Search-AccessVbe -DbPath $DbPath -SearchText $SearchText -MatchCase:$MatchCase -UseRegex:$UseRegex -MaxResults $MaxResults
    $vbaMatches = [System.Collections.Generic.List[object]]::new()
    foreach ($group in @($vbaResult.results)) {
        foreach ($m in @($group.matches)) {
            $vbaMatches.Add([PSCustomObject][ordered]@{
                object_type = $group.object_type
                object_name = $group.object_name
                line        = $m.line
                content     = $m.content
            })
        }
    }
    $total = $vbaMatches.Count
    $truncated = [bool]$vbaResult.truncated

    # 2. Query matches
    $queryMatches = [System.Collections.Generic.List[object]]::new()
    if (-not $truncated) {
        $remaining = $MaxResults - $total
        $qryResult = Search-AccessQuery -DbPath $DbPath -SearchText $SearchText -MatchCase:$MatchCase -UseRegex:$UseRegex -MaxResults $remaining
        foreach ($q in @($qryResult.results)) { $queryMatches.Add($q) }
        $total += $qryResult.total_matches
        $truncated = [bool]$qryResult.truncated
    }

    # 3. Control property matches — search form/report exports
    $controlMatches = [System.Collections.Generic.List[object]]::new()
    if (-not $truncated) {
        $app = $script:AccessSession.App
        $objects = Get-AccessObject -DbPath $DbPath -ObjectType all
        foreach ($objType in @('form', 'report')) {
            if ($truncated) { break }
            foreach ($objName in @($objects.$objType)) {
                if ($truncated) { break }
                try {
                    $tmp = [System.IO.Path]::GetTempFileName()
                    try {
                        $app.SaveAsText($script:AC_TYPE[$objType], $objName, $tmp)
                        $rawText = (Read-TempFile -Path $tmp).Content
                    } finally {
                        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                    }
                    foreach ($line in $rawText.Split("`n")) {
                        $stripped = $line.Trim()
                        foreach ($prop in $script:CONTROL_SEARCH_PROPS) {
                            if ($stripped.StartsWith("$prop =")) {
                                $valuePart = $stripped.Substring($prop.Length + 2).Trim()
                                if (Test-TextMatch -Needle $SearchText -Haystack $valuePart -MatchCase:$MatchCase -UseRegex:$UseRegex) {
                                    $controlMatches.Add([PSCustomObject][ordered]@{
                                        object_type = $objType
                                        object_name = $objName
                                        property    = $prop
                                        value       = $valuePart
                                    })
                                    $total++
                                    if ($total -ge $MaxResults) { $truncated = $true }
                                    break
                                }
                            }
                        }
                    }
                } catch { continue }
            }
        }
    }

    $out = [ordered]@{
        search_text     = $SearchText
        vba_matches     = @($vbaMatches)
        query_matches   = @($queryMatches)
        control_matches = @($controlMatches)
        total_matches   = $total
    }
    if ($truncated) { $out['truncated'] = $true }
    Format-AccessOutput -AsJson:$AsJson -Data $out
}

function Invoke-AccessMacro {
    <#
    .SYNOPSIS
        Run an Access macro by name.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER MacroName
        Name of the macro to run.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-AccessMacro -DbPath "C:\db.accdb" -MacroName "AutoExec"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$MacroName,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Invoke-AccessMacro'
    if (-not $MacroName) { throw "Invoke-AccessMacro: -MacroName is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    try {
        $app.DoCmd.RunMacro($MacroName)
    } catch {
        throw "Error running macro '$MacroName': $_"
    }

    Format-AccessOutput -AsJson:$AsJson -Data @{
        macro_name = $MacroName
        status     = 'executed'
    }
}

function Invoke-AccessVba {
    <#
    .SYNOPSIS
        Call a VBA Sub/Function via Application.Run or Forms COM access.
    .DESCRIPTION
        Supports two syntaxes:
        - 'ModuleName.ProcName' or 'ProcName' -> Application.Run (standard modules)
        - 'Forms.FormName.Method' -> COM Forms() access (form must be open)
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER Procedure
        Procedure name or qualified path.
    .PARAMETER Arguments
        Arguments to pass to the procedure (max 30 for Application.Run).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-AccessVba -DbPath "C:\db.accdb" -Procedure "Module1.Calculate" -Arguments @(42)
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$Procedure,
        [object[]]$Arguments,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Invoke-AccessVba'
    if (-not $Procedure) { throw "Invoke-AccessVba: -Procedure is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $callArgs = if ($Arguments) { $Arguments } else { @() }

    if ($callArgs.Count -gt 30) {
        throw 'Application.Run supports max 30 arguments.'
    }

    # Forms.FormName.Method -> direct COM access
    if ($Procedure.Contains('.')) {
        $parts = $Procedure.Split('.', 3)
        if ($parts[0] -eq 'Forms' -and $parts.Count -eq 3) {
            $formName   = $parts[1]
            $methodName = $parts[2]
            try {
                $form = $app.Forms($formName)
                if ($callArgs.Count -gt 0) {
                    $result = $form.GetType().InvokeMember(
                        $methodName,
                        [System.Reflection.BindingFlags]::InvokeMethod,
                        $null, $form, $callArgs
                    )
                } else {
                    # Try method call, fall back to property
                    try {
                        $result = $form.GetType().InvokeMember(
                            $methodName,
                            [System.Reflection.BindingFlags]::InvokeMethod,
                            $null, $form, @()
                        )
                    } catch {
                        $result = $form.GetType().InvokeMember(
                            $methodName,
                            [System.Reflection.BindingFlags]::GetProperty,
                            $null, $form, @()
                        )
                    }
                }
            } catch {
                throw "Error calling Forms('$formName').$methodName : $_. Make sure the form is open."
            }

            return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
                procedure = $Procedure
                result    = $result
                status    = 'executed'
            })
        }
    }

    # Standard Application.Run
    try {
        $result = switch ($callArgs.Count) {
            0  { $app.Run($Procedure) }
            1  { $app.Run($Procedure, $callArgs[0]) }
            2  { $app.Run($Procedure, $callArgs[0], $callArgs[1]) }
            3  { $app.Run($Procedure, $callArgs[0], $callArgs[1], $callArgs[2]) }
            4  { $app.Run($Procedure, $callArgs[0], $callArgs[1], $callArgs[2], $callArgs[3]) }
            5  { $app.Run($Procedure, $callArgs[0], $callArgs[1], $callArgs[2], $callArgs[3], $callArgs[4]) }
            default {
                # Build args array for Invoke
                $invokeArgs = @($Procedure) + $callArgs
                $app.GetType().InvokeMember('Run', [System.Reflection.BindingFlags]::InvokeMethod, $null, $app, $invokeArgs)
            }
        }
    } catch {
        throw "Error running '$Procedure': $_"
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        procedure = $Procedure
        result    = $result
        status    = 'executed'
    })
}

function Invoke-AccessEval {
    <#
    .SYNOPSIS
        Evaluate a VBA/Access expression via Application.Eval.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER Expression
        Expression to evaluate (e.g., "Date()", "DLookup(...)").
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-AccessEval -DbPath "C:\db.accdb" -Expression "Date()"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$Expression,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Invoke-AccessEval'
    if (-not $Expression) { throw "Invoke-AccessEval: -Expression is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    try {
        $result = $app.Eval($Expression)
    } catch {
        throw "Error evaluating '$Expression': $_"
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        expression = $Expression
        result     = $result
        status     = 'evaluated'
    })
}

function Test-AccessVbaCompile {
    <#
    .SYNOPSIS
        Compile and save all VBA modules. Returns status and error location if compilation fails.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Test-AccessVbaCompile -DbPath "C:\db.accdb"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Test-AccessVbaCompile'

    $app = Connect-AccessDB -DbPath $DbPath

    try {
        $app.RunCommand($script:AC_CMD_COMPILE)
    } catch {
        # Try to get error location from VBE
        $errLoc = $null
        try {
            $pane = $app.VBE.ActiveCodePane
            if ($null -ne $pane) {
                $cm = $pane.CodeModule
                $startLine = 0; $startCol = 0; $endLine = 0; $endCol = 0
                $pane.GetSelection([ref]$startLine, [ref]$startCol, [ref]$endLine, [ref]$endCol)
                $errCode = ''
                if ($startLine -gt 0 -and $cm.CountOfLines -ge $startLine) {
                    $errCode = $cm.Lines($startLine, 1)
                }
                $errLoc = [ordered]@{
                    component = $cm.Parent.Name
                    line      = $startLine
                    code      = $errCode
                }
            }
        } catch {}

        $result = [ordered]@{
            status       = 'error'
            error_detail = "VBA compilation error: $_"
        }
        if ($errLoc) { $result['error_location'] = $errLoc }
        return Format-AccessOutput -AsJson:$AsJson -Data $result
    }

    # Invalidate caches — compilation may change module state
    $script:AccessSession.VbeCodeCache = @{}
    $script:AccessSession.CmCache = @{}

    Format-AccessOutput -AsJson:$AsJson -Data @{ status = 'compiled' }
}

function Import-AccessVbaFile {
    <#
    .SYNOPSIS
        Import a .bas (standard module) or .cls (class module) file into an Access
        database via VBComponents.Import. Validates ANSI encoding and auto-converts
        if needed. Replaces any existing component with the same name.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER FilePath
        Path to the .bas or .cls file to import.
    .PARAMETER Force
        Auto-convert non-ANSI files to a temp ANSI copy before importing (default).
        Set -Force:$false to error on non-ANSI files instead of converting.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Import-AccessVbaFile -DbPath "C:\db.accdb" -FilePath "C:\modules\clsHelper.cls"
    .EXAMPLE
        Import-AccessVbaFile -DbPath "C:\db.accdb" -FilePath "C:\modules\modUtils.bas" -AsJson
    .EXAMPLE
        # Import multiple files
        Get-ChildItem "C:\vba\*.cls","C:\vba\*.bas" | ForEach-Object {
            Import-AccessVbaFile -DbPath "C:\db.accdb" -FilePath $_.FullName -AsJson
        }
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$FilePath,
        [switch]$Force = $true,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Import-AccessVbaFile'
    if (-not $FilePath) { throw "Import-AccessVbaFile: -FilePath is required." }
    if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
        throw "Import-AccessVbaFile: File not found: $FilePath"
    }

    $FilePath = (Resolve-Path -LiteralPath $FilePath).Path
    $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()
    if ($ext -notin '.bas', '.cls') {
        throw "Import-AccessVbaFile: Only .bas and .cls files are supported. Got '$ext'."
    }

    # Validate encoding
    $encCheck = Test-VbaFileEncoding -Path $FilePath
    $importPath = $FilePath
    $converted = $false
    $tmpPath = $null

    if (-not $encCheck.IsAnsi) {
        if (-not $Force) {
            throw "Import-AccessVbaFile: $($encCheck.Reason) Use -Force to auto-convert."
        }
        Write-Verbose "Non-ANSI encoding detected ($($encCheck.Encoding)). Converting to ANSI temp copy."
        $tmpPath = ConvertTo-AnsiTempFile -SourcePath $FilePath
        $importPath = $tmpPath
        $converted = $true
    }

    try {
        $app = Connect-AccessDB -DbPath $DbPath
        $proj = $app.VBE.ActiveVBProject

        # Derive module name from the file's Attribute VB_Name or from filename
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)

        # Remove existing component if present
        try {
            $existing = $proj.VBComponents.Item($moduleName)
            if ($null -ne $existing) {
                $proj.VBComponents.Remove($existing)
                Write-Verbose "Removed existing component: $moduleName"
            }
        } catch {
            # Component doesn't exist — that's fine
        }

        # Import via VBComponents.Import — correctly handles .cls as class module
        $imported = $proj.VBComponents.Import($importPath)
        $typeName = switch ($imported.Type) {
            1 { 'standard_module' }
            2 { 'class_module' }
            default { "type_$($imported.Type)" }
        }

        # Invalidate VBE caches
        $cacheKey = "module:$($imported.Name)"
        $script:AccessSession.VbeCodeCache.Remove($cacheKey)
        $script:AccessSession.CmCache.Remove($cacheKey)

        $result = [ordered]@{
            status       = 'imported'
            name         = $imported.Name
            module_type  = $typeName
            source_file  = $FilePath
            converted    = $converted
        }
        if ($converted) {
            $result['original_encoding'] = $encCheck.Encoding
        }
        Format-AccessOutput -AsJson:$AsJson -Data $result
    } finally {
        if ($tmpPath -and (Test-Path -LiteralPath $tmpPath)) {
            Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-AccessVbaFileEncoding {
    <#
    .SYNOPSIS
        Check whether a .bas or .cls file has the correct ANSI encoding
        (Windows-1252, no BOM) required by VBComponents.Import.
    .PARAMETER FilePath
        Path to the .bas or .cls file to check.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Test-AccessVbaFileEncoding -FilePath "C:\modules\clsHelper.cls" -AsJson
    .EXAMPLE
        Get-ChildItem "C:\vba\*" -Include *.bas,*.cls | ForEach-Object {
            Test-AccessVbaFileEncoding -FilePath $_.FullName -AsJson
        }
    #>
    [CmdletBinding()]
    param(
        [string]$FilePath,
        [switch]$AsJson
    )

    if (-not $FilePath) { throw "Test-AccessVbaFileEncoding: -FilePath is required." }
    if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
        throw "Test-AccessVbaFileEncoding: File not found: $FilePath"
    }

    $FilePath = (Resolve-Path -LiteralPath $FilePath).Path
    $check = Test-VbaFileEncoding -Path $FilePath

    $result = [ordered]@{
        file     = $FilePath
        is_ansi  = $check.IsAnsi
        encoding = $check.Encoding
    }
    if (-not $check.IsAnsi) {
        $result['reason'] = $check.Reason
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}
