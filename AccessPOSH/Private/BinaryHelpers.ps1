# Private/BinaryHelpers.ps1 — Binary section handling, code-behind splitting, field property helper

function Remove-BinarySections {
    <#
    .SYNOPSIS
        Strip binary sections (PrtMip, PrtDevMode, NameMap, etc.) from a form/report export.
        Reduces size ~20x without affecting VBA or controls.
        Also removes the Checksum line (Access recalculates on import).
    #>
    param(
        [string]$Text
    )
    if (-not $Text) { throw "Remove-BinarySections: -Text is required." }

    $lines = $Text.Split([string[]]@("`r`n", "`n"), [System.StringSplitOptions]::None)
    $result = [System.Collections.Generic.List[string]]::new($lines.Count)
    $skipDepth  = 0
    $skipIndent = ''

    foreach ($line in $lines) {
        $rstripped = $line.TrimEnd("`r", "`n")
        $stripped  = $rstripped.TrimStart()
        $indent    = $rstripped.Substring(0, $rstripped.Length - $stripped.Length)

        if ($skipDepth -gt 0) {
            if ($stripped -eq 'End' -and $indent -eq $skipIndent) {
                $skipDepth--
            }
            continue
        }

        # Checksum line at root level
        if ($rstripped -match '^\s*Checksum\s*=\s*') {
            continue
        }

        # Does a binary block begin here?
        if ($rstripped -match '^(\s*)(\w+)\s*=\s*Begin\s*$') {
            $blockIndent = $Matches[1]
            $blockName   = $Matches[2]
            if ($script:BINARY_SECTIONS.Contains($blockName)) {
                $skipIndent = $blockIndent
                $skipDepth  = 1
                continue
            }
        }

        $result.Add($line)
    }

    return ($result -join "`r`n")
}

function Get-BinaryBlocks {
    <#
    .SYNOPSIS
        Extract binary Begin...End blocks from the original form/report export.
        Returns a hashtable: { section_name = full_block_text }.
    #>
    param(
        [string]$Text
    )
    if (-not $Text) { throw "Get-BinaryBlocks: -Text is required." }

    $blocks = @{}
    $lines = $Text.Split([string[]]@("`r`n", "`n"), [System.StringSplitOptions]::None)
    $i = 0

    while ($i -lt $lines.Count) {
        $line = $lines[$i]
        $rstripped = $line.TrimEnd("`r", "`n")

        if ($rstripped -match '^(\s*)(\w+)\s*=\s*Begin\s*$') {
            $blockIndent = $Matches[1]
            $blockName   = $Matches[2]
            if ($script:BINARY_SECTIONS.Contains($blockName)) {
                $blockLines = [System.Collections.Generic.List[string]]::new()
                $blockLines.Add($line)
                $j = $i + 1
                while ($j -lt $lines.Count) {
                    $bl = $lines[$j]
                    $blr = $bl.TrimEnd("`r", "`n")
                    $bls = $blr.TrimStart()
                    $blIndent = $blr.Substring(0, $blr.Length - $bls.Length)
                    $blockLines.Add($bl)
                    if ($bls -eq 'End' -and $blIndent -eq $blockIndent) {
                        break
                    }
                    $j++
                }
                $blocks[$blockName] = ($blockLines -join "`r`n")
                $i = $j + 1
                continue
            }
        }
        $i++
    }

    return $blocks
}

function Restore-BinarySections {
    <#
    .SYNOPSIS
        Re-inject binary sections from the current Access object's export,
        before calling LoadFromText with edited code.
        If the object doesn't exist yet, returns the code unmodified.
    #>
    param(
        $App,
        [string]$ObjectType,
        [string]$Name,
        [string]$NewCode
    )
    if (-not $App) { throw "Restore-BinarySections: -App is required." }
    if (-not $ObjectType) { throw "Restore-BinarySections: -ObjectType is required." }
    if (-not $Name) { throw "Restore-BinarySections: -Name is required." }
    if (-not $NewCode) { throw "Restore-BinarySections: -NewCode is required." }

    $tmp = [System.IO.Path]::GetTempFileName()
    try {
        try {
            $App.SaveAsText($script:AC_TYPE[$ObjectType], $Name, $tmp)
        } catch {
            Write-Verbose "Restore-BinarySections: '$Name' doesn't exist yet — importing without binary sections"
            return $NewCode
        }
        $original = (Read-TempFile -Path $tmp).Content
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }

    $blocks = Get-BinaryBlocks -Text $original
    if ($blocks.Count -eq 0) { return $NewCode }

    # Inject blocks just before "End Form" / "End Report"
    $lines = $NewCode.Split([string[]]@("`r`n", "`n"), [System.StringSplitOptions]::None)
    $result = [System.Collections.Generic.List[string]]::new($lines.Count + 50)
    $inTopForm = $false
    $injected  = $false

    foreach ($line in $lines) {
        $stripped = $line.Trim()

        if ($stripped -match '^\s*Begin\s+(Form|Report)\s*$') {
            $inTopForm = $true
        }

        if ($inTopForm -and (-not $injected) -and $stripped -match '^\s*End\s+(Form|Report)\s*$') {
            foreach ($blockText in $blocks.Values) {
                $result.Add($blockText)
            }
            $injected  = $true
            $inTopForm = $false
        }

        $result.Add($line)
    }

    return ($result -join "`r`n")
}

function Split-CodeBehind {
    <#
    .SYNOPSIS
        Separate form/report export text into (form_text, vba_code).
        If CodeBehindForm/CodeBehindReport marker exists, splits there.
        Returns [PSCustomObject]@{ FormText; VbaCode }
    #>
    param(
        [string]$Code
    )
    if (-not $Code) { throw "Split-CodeBehind: -Code is required." }

    foreach ($marker in @('CodeBehindForm', 'CodeBehindReport')) {
        $idx = $Code.IndexOf($marker)
        if ($idx -ge 0) {
            $formPart = $Code.Substring(0, $idx).TrimEnd() + "`n"
            $remainder = $Code.Substring($idx)
            $parts = $remainder -split "`n", 2
            $vbaCode = if ($parts.Count -gt 1) { $parts[1] } else { '' }
            # Strip Attribute VB_ lines (auto-generated)
            $vbaLines = foreach ($line in $vbaCode.Split("`n")) {
                if ($line.Trim() -notmatch '^Attribute VB_') { $line }
            }
            $vbaCode = ($vbaLines -join "`n").Trim()
            return [PSCustomObject]@{ FormText = $formPart; VbaCode = $vbaCode }
        }
    }
    return [PSCustomObject]@{ FormText = $Code; VbaCode = '' }
}

function Set-FieldProperty {
    <#
    .SYNOPSIS
        Set a field-level DAO property, creating it if it doesn't exist.
    #>
    param(
        $Db,
        [string]$TableName,
        [string]$FieldName,
        [string]$PropertyName,
        $Value
    )
    if (-not $Db) { throw "Set-FieldProperty: -Db is required." }
    if (-not $TableName) { throw "Set-FieldProperty: -TableName is required." }
    if (-not $FieldName) { throw "Set-FieldProperty: -FieldName is required." }
    if (-not $PropertyName) { throw "Set-FieldProperty: -PropertyName is required." }

    $fld = $Db.TableDefs($TableName).Fields($FieldName)
    try {
        $fld.Properties($PropertyName).Value = $Value
    } catch {
        $prop = $fld.CreateProperty($PropertyName, 10, $Value)  # 10 = dbText
        $fld.Properties.Append($prop)
    }
}

function Invoke-VbaAfterImport {
    <#
    .SYNOPSIS
        Internal: Inject VBA code into a form/report after LoadFromText import.
        Opens in design, enables HasModule, then injects via VBE CodeModule.
    #>
    param(
        $App,
        [string]$ObjectType,
        [string]$Name,
        [string]$VbaCode
    )
    if (-not $App) { throw "Invoke-VbaAfterImport: -App is required." }
    if (-not $ObjectType) { throw "Invoke-VbaAfterImport: -ObjectType is required." }
    if (-not $Name) { throw "Invoke-VbaAfterImport: -Name is required." }

    if (-not $VbaCode.Trim()) { return }

    # 1. Open in design and enable HasModule
    $acObjType = if ($ObjectType -eq 'form') { $script:AC_FORM } else { $script:AC_REPORT }
    if ($ObjectType -eq 'form') {
        $App.DoCmd.OpenForm($Name, $script:AC_DESIGN)
    } else {
        $App.DoCmd.OpenReport($Name, $script:AC_DESIGN)
    }
    try {
        $obj = if ($ObjectType -eq 'form') { $App.Forms($Name) } else { $App.Reports($Name) }
        $obj.HasModule = $true
    } finally {
        $App.DoCmd.Close($acObjType, $Name, $script:AC_SAVE_YES)
    }

    # 2. Clear VBE cache
    $cacheKey = "${ObjectType}:${Name}"
    $script:AccessSession.CmCache.Remove($cacheKey)
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)

    # 3. Get CodeModule and inject VBA
    $prefix = $script:VBE_PREFIX[$ObjectType]
    $compName = "${prefix}${Name}"
    $cm = $App.VBE.ActiveVBProject.VBComponents($compName).CodeModule

    $total = $cm.CountOfLines
    if ($total -gt 0) {
        $cm.DeleteLines(1, $total)
    }

    # Normalize line endings to CRLF
    $VbaCode = $VbaCode -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
    if (-not $VbaCode.EndsWith("`r`n")) { $VbaCode += "`r`n" }

    $cm.InsertLines(1, $VbaCode)

    # Invalidate caches
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)
    $script:AccessSession.CmCache.Remove($cacheKey)
}
