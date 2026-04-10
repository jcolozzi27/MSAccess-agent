# Public/FormReportOps.ps1 — Form and report creation, properties, controls

function New-AccessForm {
    <#
    .SYNOPSIS
        Create a new blank form.
    .DESCRIPTION
        Uses CreateForm() which auto-names the form (Form1, Form2...).
        Saves, closes, then renames to the desired name.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER FormName
        Desired name for the new form.
    .PARAMETER HasHeader
        Toggle form header/footer sections.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-AccessForm -DbPath "C:\db.accdb" -FormName "frmCustomers"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$FormName,
        [switch]$HasHeader,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'New-AccessForm'
    if (-not $FormName) { throw "New-AccessForm: -FormName is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    $autoName = $null
    try {
        $form = $app.CreateForm()
        $autoName = $form.Name

        if ($HasHeader) {
            try {
                $app.Visible = $true
                $app.RunCommand(36)  # acCmdFormHdrFtr — toggle header/footer
            } catch {
                Write-Warning "Could not toggle header/footer via RunCommand: $_"
            }
        }

        # Save with auto-name (no dialog)
        $app.DoCmd.Save($script:AC_TYPE['form'], $autoName)
        # Close without prompt (already saved)
        $app.DoCmd.Close($script:AC_TYPE['form'], $autoName, 2)  # acSaveNo=2

        # Rename to desired name
        if ($autoName -ne $FormName) {
            $app.DoCmd.Rename($FormName, $script:AC_TYPE['form'], $autoName)
        }

        return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            name         = $FormName
            created_from = $autoName
            has_header   = [bool]$HasHeader
        })
    } catch {
        if ($autoName) {
            try { $app.DoCmd.Close($script:AC_TYPE['form'], $autoName, 2) } catch {}
            try { $app.DoCmd.DeleteObject($script:AC_TYPE['form'], $autoName) } catch {}
        }
        throw "Error creating form '$FormName': $_"
    } finally {
        $script:AccessSession.VbeCodeCache = @{}
        $script:AccessSession.ControlsCache = @{}
        $script:AccessSession.CmCache = @{}
    }
}

function Get-AccessFormProperty {
    <#
    .SYNOPSIS
        Read properties from a form or report by opening it in Design view.
    .DESCRIPTION
        If PropertyNames is omitted, reads all readable properties.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER PropertyNames
        Array of property names to read. If omitted, reads all.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessFormProperty -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -PropertyNames "Caption","RecordSource"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName,
        [string[]]$PropertyNames,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessFormProperty'
    if (-not $ObjectType) { throw "Get-AccessFormProperty: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "Get-AccessFormProperty: -ObjectName is required." }

    $null = Connect-AccessDB -DbPath $DbPath
    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    $properties = [ordered]@{}
    $errors = [ordered]@{}
    try {
        if ($PropertyNames) {
            foreach ($pname in $PropertyNames) {
                try {
                    if ($ObjectType -eq 'form') {
                        $val = $script:AccessSession.App.Screen.ActiveForm.Properties.Item($pname).Value
                    } else {
                        $val = $script:AccessSession.App.Screen.ActiveReport.Properties.Item($pname).Value
                    }
                    $properties[$pname] = ConvertTo-SafeValue -Value $val
                } catch {
                    $errors[$pname] = "$_"
                }
            }
        } else {
            if ($ObjectType -eq 'form') {
                $cnt = $script:AccessSession.App.Screen.ActiveForm.Properties.Count
            } else {
                $cnt = $script:AccessSession.App.Screen.ActiveReport.Properties.Count
            }
            for ($i = 0; $i -lt $cnt; $i++) {
                try {
                    if ($ObjectType -eq 'form') {
                        $pName  = $script:AccessSession.App.Screen.ActiveForm.Properties.Item($i).Name
                        $pValue = $script:AccessSession.App.Screen.ActiveForm.Properties.Item($i).Value
                    } else {
                        $pName  = $script:AccessSession.App.Screen.ActiveReport.Properties.Item($i).Name
                        $pValue = $script:AccessSession.App.Screen.ActiveReport.Properties.Item($i).Value
                    }
                    $properties[$pName] = ConvertTo-SafeValue -Value $pValue
                } catch { }
            }
        }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    $result = [ordered]@{
        object     = $ObjectName
        type       = $ObjectType
        properties = $properties
    }
    if ($errors.Count -gt 0) { $result['errors'] = $errors }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Set-AccessFormProperty {
    <#
    .SYNOPSIS
        Set properties on a form or report by opening it in Design view.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER Properties
        Hashtable of property name/value pairs to set.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessFormProperty -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -Properties @{ Caption = "My Form"; RecordSource = "tblCustomers" }
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName,
        [hashtable]$Properties,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessFormProperty'
    if (-not $ObjectType) { throw "Set-AccessFormProperty: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "Set-AccessFormProperty: -ObjectName is required." }
    if (-not $Properties -or $Properties.Count -eq 0) { throw "Set-AccessFormProperty: -Properties is required." }

    $null = Connect-AccessDB -DbPath $DbPath
    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    $applied = [System.Collections.Generic.List[string]]::new()
    $errors = [ordered]@{}
    try {
        foreach ($key in $Properties.Keys) {
            try {
                $coerced = ConvertTo-CoercedProp -Value $Properties[$key]
                if ($ObjectType -eq 'form') {
                    $script:AccessSession.App.Screen.ActiveForm.$key = $coerced
                } else {
                    $script:AccessSession.App.Screen.ActiveReport.$key = $coerced
                }
                $applied.Add($key)
            } catch {
                $errors[$key] = "$_"
            }
        }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    $result = [ordered]@{
        applied = @($applied)
        errors  = $errors
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Get-AccessControl {
    <#
    .SYNOPSIS
        List controls in a form or report (from cached parsed export).
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessControl -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessControl'
    if (-not $ObjectType) { throw "Get-AccessControl: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "Get-AccessControl: -ObjectName is required." }

    $parsed = Get-ParsedControls -DbPath $DbPath -ObjectType $ObjectType -ObjectName $ObjectName
    $controls = @(
        $parsed.controls | Where-Object { $_.name.Trim() } | ForEach-Object {
            $c = [ordered]@{}
            foreach ($prop in $_.PSObject.Properties) {
                if ($prop.Name -ne 'raw_block') {
                    $c[$prop.Name] = $prop.Value
                }
            }
            [PSCustomObject]$c
        }
    )

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        count    = $controls.Count
        controls = $controls
    })
}

function Get-AccessControlDetail {
    <#
    .SYNOPSIS
        Get the full definition of a single control by name (includes raw_block).
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER ControlName
        Name of the control.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessControlDetail -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -ControlName "txtLastName"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName,
        [string]$ControlName,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessControlDetail'
    if (-not $ObjectType) { throw "Get-AccessControlDetail: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "Get-AccessControlDetail: -ObjectName is required." }
    if (-not $ControlName) { throw "Get-AccessControlDetail: -ControlName is required." }

    $parsed = Get-ParsedControls -DbPath $DbPath -ObjectType $ObjectType -ObjectName $ObjectName
    $found = $parsed.controls | Where-Object { $_.name -ieq $ControlName } | Select-Object -First 1
    if (-not $found) {
        $names = @($parsed.controls | ForEach-Object { $_.name })
        throw "Control '$ControlName' not found in '$ObjectName'. Available controls: $($names -join ', ')"
    }

    Format-AccessOutput -AsJson:$AsJson -Data $found
}

function New-AccessControl {
    <#
    .SYNOPSIS
        Create a new control on a form or report.
    .DESCRIPTION
        Opens the form/report in Design view, calls CreateControl/CreateReportControl,
        sets properties, saves and closes.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER ControlType
        Control type: name ('CommandButton') or number (104).
    .PARAMETER Properties
        Hashtable of properties. Special keys: section, parent, column_name, left, top, width, height.
    .PARAMETER ClassName
        ProgID for ActiveX controls (type 119).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-AccessControl -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -ControlType "CommandButton" -Properties @{ Name = "btnSave"; Caption = "Save" }
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName,
        $ControlType,
        [hashtable]$Properties = @{},
        [string]$ClassName,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'New-AccessControl'
    if (-not $ObjectType) { throw "New-AccessControl: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "New-AccessControl: -ObjectName is required." }
    if (-not $ControlType) { throw "New-AccessControl: -ControlType is required." }

    $app = Connect-AccessDB -DbPath $DbPath

    # Resolve control type
    $ctype = $ControlType
    if ($ctype -is [string]) {
        $key = $ctype.ToLower()
        if ($script:CTRL_TYPE_BY_NAME.ContainsKey($key)) {
            $ctype = $script:CTRL_TYPE_BY_NAME[$key]
        } else {
            $intVal = 0
            if ([int]::TryParse($ctype, [ref]$intVal)) { $ctype = $intVal }
            else { throw "Unknown control type: '$ControlType'" }
        }
    }
    $ctype = [int]$ctype

    # Extract structural params from Properties (don't set as COM properties later)
    $p = @{} + $Properties  # copy
    $section    = 0
    $parent     = ''
    $columnName = ''
    $left       = -1
    $top        = -1
    $width      = -1
    $height     = -1

    if ($p.ContainsKey('section')) {
        $secVal = "$($p['section'])".ToLower()
        if ($script:SECTION_MAP.ContainsKey($secVal)) { $section = $script:SECTION_MAP[$secVal] }
        else { $section = [int](ConvertTo-CoercedProp -Value $p['section']) }
        $p.Remove('section')
    }
    if ($p.ContainsKey('parent'))      { $parent     = "$($p['parent'])";      $p.Remove('parent') }
    if ($p.ContainsKey('column_name')) { $columnName = "$($p['column_name'])"; $p.Remove('column_name') }
    if ($p.ContainsKey('left'))   { $left   = [int](ConvertTo-CoercedProp -Value $p['left']);   $p.Remove('left') }
    if ($p.ContainsKey('top'))    { $top    = [int](ConvertTo-CoercedProp -Value $p['top']);    $p.Remove('top') }
    if ($p.ContainsKey('width'))  { $width  = [int](ConvertTo-CoercedProp -Value $p['width']);  $p.Remove('width') }
    if ($p.ContainsKey('height')) { $height = [int](ConvertTo-CoercedProp -Value $p['height']); $p.Remove('height') }

    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    try {
        try {
            if ($ObjectType -eq 'form') {
                $ctrl = $app.CreateControl($ObjectName, $ctype, $section, $parent, $columnName, $left, $top, $width, $height)
            } else {
                $ctrl = $app.CreateReportControl($ObjectName, $ctype, $section, $parent, $columnName, $left, $top, $width, $height)
            }
        } catch {
            $secNames = @($script:SECTION_MAP.GetEnumerator() | Where-Object { $_.Value -eq $section } | ForEach-Object { $_.Key })
            throw ("Error creating control in section=$section ($($secNames -join ', ')): $_. " +
                   "Valid sections: 0=Detail, 1=Header, 2=Footer, 3=PageHeader, 4=PageFooter")
        }

        # ActiveX: set ProgID via Class property
        if ($ClassName -and $ctype -eq 119) {
            try { $ctrl.Class = $ClassName } catch { Write-Warning "Could not set Class='$ClassName': $_" }
        }

        $propErrors = [ordered]@{}
        foreach ($key in $p.Keys) {
            try {
                $ctrl.$key = ConvertTo-CoercedProp -Value $p[$key]
            } catch {
                $propErrors[$key] = "$_"
            }
        }

        $resultData = [ordered]@{
            name         = $ctrl.Name
            control_type = $ctype
            type_name    = if ($script:CTRL_TYPE.ContainsKey($ctype)) { $script:CTRL_TYPE[$ctype] } else { "Type$ctype" }
        }
        if ($propErrors.Count -gt 0) { $resultData['property_errors'] = $propErrors }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    Format-AccessOutput -AsJson:$AsJson -Data $resultData
}

function Remove-AccessControl {
    <#
    .SYNOPSIS
        Delete a control from a form or report.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER ControlName
        Name of the control to delete.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-AccessControl -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -ControlName "txtOld"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName,
        [string]$ControlName,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Remove-AccessControl'
    if (-not $ObjectType) { throw "Remove-AccessControl: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "Remove-AccessControl: -ObjectName is required." }
    if (-not $ControlName) { throw "Remove-AccessControl: -ControlName is required." }

    $app = Connect-AccessDB -DbPath $DbPath
    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    try {
        if ($ObjectType -eq 'form') {
            $app.DeleteControl($ObjectName, $ControlName)
        } else {
            $app.DeleteReportControl($ObjectName, $ControlName)
        }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status  = "OK: control '$ControlName' deleted from '$ObjectName'"
        control = $ControlName
        object  = $ObjectName
    })
}

function Set-AccessControlProperty {
    <#
    .SYNOPSIS
        Modify properties of an existing control on a form or report.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER ControlName
        Name of the control to modify.
    .PARAMETER Properties
        Hashtable of property name/value pairs to set.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessControlProperty -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -ControlName "txtName" -Properties @{ Caption = "Full Name"; Width = 3000 }
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName,
        [string]$ControlName,
        [hashtable]$Properties,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessControlProperty'
    if (-not $ObjectType) { throw "Set-AccessControlProperty: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "Set-AccessControlProperty: -ObjectName is required." }
    if (-not $ControlName) { throw "Set-AccessControlProperty: -ControlName is required." }
    if (-not $Properties -or $Properties.Count -eq 0) { throw "Set-AccessControlProperty: -Properties is required." }

    $null = Connect-AccessDB -DbPath $DbPath
    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    $applied = [System.Collections.Generic.List[string]]::new()
    $errors = [ordered]@{}
    try {
        foreach ($key in $Properties.Keys) {
            try {
                $coerced = ConvertTo-CoercedProp -Value $Properties[$key]
                if ($ObjectType -eq 'form') {
                    $script:AccessSession.App.Screen.ActiveForm.Controls.Item($ControlName).$key = $coerced
                } else {
                    $script:AccessSession.App.Screen.ActiveReport.Controls.Item($ControlName).$key = $coerced
                }
                $applied.Add($key)
            } catch {
                $errors[$key] = "$_"
            }
        }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        applied = @($applied)
        errors  = $errors
    })
}

function Set-AccessControlBatch {
    <#
    .SYNOPSIS
        Modify properties of multiple controls in a single Design view session.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        'form' or 'report'.
    .PARAMETER ObjectName
        Name of the form or report.
    .PARAMETER Controls
        Array of hashtables, each with 'name' (string) and 'props' (hashtable).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessControlBatch -DbPath "C:\db.accdb" -ObjectType form -ObjectName "frmMain" -Controls @(
            @{ name = "txtFirst"; props = @{ Width = 3000 } },
            @{ name = "txtLast";  props = @{ Width = 3000 } }
        )
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [ValidateSet('form','report')][string]$ObjectType,
        [string]$ObjectName,
        [array]$Controls,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessControlBatch'
    if (-not $ObjectType) { throw "Set-AccessControlBatch: -ObjectType is required (form, report)." }
    if (-not $ObjectName) { throw "Set-AccessControlBatch: -ObjectName is required." }
    if (-not $Controls -or $Controls.Count -eq 0) { throw "Set-AccessControlBatch: -Controls is required." }

    if ($Controls.Count -eq 0) {
        return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{ error = 'No controls provided.' })
    }

    $null = Connect-AccessDB -DbPath $DbPath
    Open-InDesignView -ObjectType $ObjectType -ObjectName $ObjectName
    $results = [System.Collections.Generic.List[object]]::new()
    try {
        foreach ($ctrlSpec in $Controls) {
            $ctrlName  = $ctrlSpec['name']
            $ctrlProps = $ctrlSpec['props']
            if (-not $ctrlProps) { $ctrlProps = @{} }
            $applied = [System.Collections.Generic.List[string]]::new()
            $errors  = [ordered]@{}
            try {
                foreach ($key in $ctrlProps.Keys) {
                    try {
                        $coerced = ConvertTo-CoercedProp -Value $ctrlProps[$key]
                        if ($ObjectType -eq 'form') {
                            $script:AccessSession.App.Screen.ActiveForm.Controls.Item($ctrlName).$key = $coerced
                        } else {
                            $script:AccessSession.App.Screen.ActiveReport.Controls.Item($ctrlName).$key = $coerced
                        }
                        $applied.Add($key)
                    } catch {
                        $errors[$key] = "$_"
                    }
                }
            } catch {
                $errors['_control'] = "Control '$ctrlName' not found: $_"
            }
            $entry = [ordered]@{ name = $ctrlName; applied = @($applied) }
            if ($errors.Count -gt 0) { $entry['errors'] = $errors }
            $results.Add([PSCustomObject]$entry)
        }
    } finally {
        Save-AndCloseDesign -ObjectType $ObjectType -ObjectName $ObjectName
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{ results = @($results) })
}
