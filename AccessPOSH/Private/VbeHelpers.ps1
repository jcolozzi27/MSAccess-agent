# Private/VbeHelpers.ps1 — VBE CodeModule access, text matching, context helpers

function Test-TextMatch {
    <#
    .SYNOPSIS
        Internal: match needle against haystack (plain substring or regex).
    #>
    param(
        [string]$Needle,
        [string]$Haystack,
        [bool]$MatchCase = $false,
        [bool]$UseRegex = $false
    )

    if ($UseRegex) {
        $opts = [System.Text.RegularExpressions.RegexOptions]::None
        if (-not $MatchCase) { $opts = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
        return [regex]::IsMatch($Haystack, $Needle, $opts)
    }
    if (-not $MatchCase) {
        return $Haystack.IndexOf($Needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
    }
    return $Haystack.Contains($Needle)
}

function Get-CodeModule {
    <#
    .SYNOPSIS
        Internal: Get cached VBE CodeModule COM object for a module/form/report.
    #>
    param(
        $App,
        [string]$ObjectType,
        [string]$ObjectName
    )

    if (-not $script:VBE_PREFIX.ContainsKey($ObjectType)) {
        throw "object_type '$ObjectType' does not support VBE. Use 'module', 'form', or 'report'."
    }

    $cacheKey = "${ObjectType}:${ObjectName}"
    $cm = $script:AccessSession.CmCache[$cacheKey]
    if ($null -ne $cm) { return $cm }

    $compName = $script:VBE_PREFIX[$ObjectType] + $ObjectName
    try {
        $project = $App.VBE.VBProjects(1)
        $component = $project.VBComponents($compName)
        $cm = $component.CodeModule
        $script:AccessSession.CmCache[$cacheKey] = $cm
        return $cm
    } catch {
        $script:AccessSession.CmCache.Remove($cacheKey)
        throw "Cannot access CodeModule '$compName'. Is 'Trust access to the VBA project object model' enabled in Access Trust Center? Error: $_"
    }
}

function Get-AllModuleCode {
    <#
    .SYNOPSIS
        Internal: Get full module text using VbeCodeCache.
    #>
    param(
        $CodeModule,
        [string]$CacheKey
    )

    if (-not $script:AccessSession.VbeCodeCache.ContainsKey($CacheKey)) {
        $total = $CodeModule.CountOfLines
        $text = if ($total -gt 0) { $CodeModule.Lines(1, $total) } else { '' }
        $script:AccessSession.VbeCodeCache[$CacheKey] = $text
    }
    return $script:AccessSession.VbeCodeCache[$CacheKey]
}

function Test-WsNormalizedMatch {
    <#
    .SYNOPSIS
        Whitespace-tolerant matching: strips leading whitespace from each line
        and does a sliding-window search. Returns hashtable with start/end 0-based
        line indices, or $null if no match.
    #>
    param(
        [string]$ProcCode,
        [string]$FindText
    )
    $procLines = $ProcCode -split "`r?`n"
    $findLines = @($FindText -split "`r?`n")
    # Remove empty trailing lines from find text
    while ($findLines.Count -gt 0 -and -not $findLines[-1].Trim()) {
        $findLines = $findLines[0..($findLines.Count - 2)]
    }
    if ($findLines.Count -eq 0) { return $null }

    $procStripped = @($procLines | ForEach-Object { $_.TrimStart() })
    $findStripped = @($findLines | ForEach-Object { $_.TrimStart() })
    $window = $findStripped.Count

    for ($i = 0; $i -le ($procStripped.Count - $window); $i++) {
        $match = $true
        for ($j = 0; $j -lt $window; $j++) {
            if ($procStripped[$i + $j] -cne $findStripped[$j]) {
                $match = $false
                break
            }
        }
        if ($match) {
            return @{ start = $i; end = ($i + $window - 1) }
        }
    }
    return $null
}

function Get-ClosestMatchContext {
    <#
    .SYNOPSIS
        When both exact and ws-normalized match fail, finds the most similar line
        using character-level similarity and returns a contextual snippet.
    #>
    param(
        [string]$ProcCode,
        [string]$FindText,
        [string]$ProcName
    )
    $procLines = $ProcCode -split "`r?`n"
    $findLines = @(($FindText -split "`r?`n") | Where-Object { $_.Trim() })
    if ($findLines.Count -eq 0) { return "Empty find text in proc '$ProcName'" }

    $ref = $findLines[0].Trim()
    $bestRatio = 0.0
    $bestIdx = 0

    for ($i = 0; $i -lt $procLines.Count; $i++) {
        $candidate = $procLines[$i].Trim()
        if (-not $candidate) { continue }
        # Longest Common Subsequence ratio (simplified SequenceMatcher)
        $shorter = if ($ref.Length -lt $candidate.Length) { $ref } else { $candidate }
        $longer  = if ($ref.Length -lt $candidate.Length) { $candidate } else { $ref }
        $matchCount = 0
        $usedIdx = -1
        foreach ($ch in $shorter.ToCharArray()) {
            $pos = $longer.IndexOf($ch, $usedIdx + 1)
            if ($pos -ge 0) { $matchCount++; $usedIdx = $pos }
        }
        $ratio = if (($ref.Length + $candidate.Length) -gt 0) { (2.0 * $matchCount) / ($ref.Length + $candidate.Length) } else { 0 }
        if ($ratio -gt $bestRatio) {
            $bestRatio = $ratio
            $bestIdx = $i
        }
    }

    # Build context: 3 lines around best candidate
    $ctxStart = [math]::Max(0, $bestIdx - 1)
    $ctxEnd   = [math]::Min($procLines.Count - 1, $bestIdx + 1)
    $contextLines = @()
    for ($j = $ctxStart; $j -le $ctxEnd; $j++) {
        $marker = if ($j -eq $bestIdx) { '>>>' } else { '   ' }
        $contextLines += "  $marker L$($j + 1): $($procLines[$j].TrimEnd())"
    }
    $pct = [math]::Round($bestRatio * 100)
    $refSnippet = if ($ref.Length -gt 80) { $ref.Substring(0, 80) } else { $ref }
    return "Best match (${pct}% similar) near line $($bestIdx + 1) of '${ProcName}':`n$($contextLines -join "`n")`n  Looking for: '$refSnippet'"
}
