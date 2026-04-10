# Public/SecurityOps.ps1 — Database security, passwords, and encryption

function Test-AccessDatabasePassword {
    <#
    .SYNOPSIS
        Check if a database is password-protected.
    .DESCRIPTION
        Uses DAO.DBEngine.120 to attempt opening the database without a password.
        If the open fails, the database is password-protected.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Test-AccessDatabasePassword'
    $app = Connect-AccessDB -DbPath $DbPath

    $hasPassword = $false
    try {
        $engine = New-Object -ComObject 'DAO.DBEngine.120'
        $resolvedPath = (Resolve-Path $DbPath).Path
        $testDb = $engine.OpenDatabase($resolvedPath, $false, $true, '')
        $testDb.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($testDb)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($engine)
        $hasPassword = $false
    } catch {
        $hasPassword = $true
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($engine) } catch {}
    }

    $result = [ordered]@{
        database     = (Split-Path $DbPath -Leaf)
        has_password = $hasPassword
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Set-AccessDatabasePassword {
    <#
    .SYNOPSIS
        Set or change the database password.
    .DESCRIPTION
        Uses DAO NewPassword to set or change the password on an Access database.
        The database must be opened exclusively for the password change to succeed.
        If the database already has a password, supply OldPassword.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$NewPassword,
        [string]$OldPassword,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessDatabasePassword'
    if (-not $NewPassword) { throw "Set-AccessDatabasePassword: -NewPassword is required." }
    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()

    $oldPwd = if ($OldPassword) { $OldPassword } else { '' }
    try {
        $db.NewPassword($oldPwd, $NewPassword)
    } catch {
        throw "Error setting database password: $_"
    }

    $result = [ordered]@{
        database = (Split-Path $DbPath -Leaf)
        action   = 'password_set'
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Remove-AccessDatabasePassword {
    <#
    .SYNOPSIS
        Remove the database password.
    .DESCRIPTION
        Uses DAO NewPassword to clear the password from an Access database.
        The current password must be supplied.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [string]$CurrentPassword,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Remove-AccessDatabasePassword'
    if (-not $CurrentPassword) { throw "Remove-AccessDatabasePassword: -CurrentPassword is required." }
    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()

    try {
        $db.NewPassword($CurrentPassword, '')
    } catch {
        throw "Error removing database password: $_"
    }

    $result = [ordered]@{
        database = (Split-Path $DbPath -Leaf)
        action   = 'password_removed'
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Get-AccessDatabaseEncryption {
    <#
    .SYNOPSIS
        Get encryption status and info for a database.
    .DESCRIPTION
        Reads DAO database properties to determine encryption provider and version,
        and tests whether the database is password-protected.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessDatabaseEncryption'
    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()

    $encrypted = $false
    $encryptionType = 'none'
    $version = $null

    try {
        $props = $db.Properties

        try {
            $encProvider = $props.Item('Encryption Provider').Value
            if ($encProvider) {
                $encrypted = $true
                $encryptionType = $encProvider
            }
        } catch {}

        $version = $db.Version
    } catch {
        Write-Warning "Could not read encryption properties: $_"
    }

    # Also do the password test
    $hasPassword = $false
    try {
        $engine = New-Object -ComObject 'DAO.DBEngine.120'
        $resolvedPath = (Resolve-Path $DbPath).Path
        $testDb = $engine.OpenDatabase($resolvedPath, $false, $true, '')
        $testDb.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($testDb)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($engine)
    } catch {
        $hasPassword = $true
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($engine) } catch {}
    }

    $result = [ordered]@{
        database        = (Split-Path $DbPath -Leaf)
        has_password    = $hasPassword
        encryption_type = $encryptionType
        db_version      = $version
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}
