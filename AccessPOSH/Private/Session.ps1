# Private/Session.ps1 — COM session management helpers

function Test-AccessAlive {
    <#
    .SYNOPSIS
        Best-effort COM liveness check (does not depend on Visible).
    #>
    if ($null -eq $script:AccessSession.App) { return $false }
    $alive = $false
    try {
        $null = Get-AccessHwnd -App $script:AccessSession.App
        $alive = $true
    }
    catch {
        try {
            $null = $script:AccessSession.App.Version
            $alive = $true
        }
        catch {
            $alive = $false
        }
    }
    return $alive
}

function Get-AccessHwnd {
    <#
    .SYNOPSIS
        Get the Access window handle. Handles hWndAccessApp being a property or method.
    #>
    param($App)

    $h = $App.hWndAccessApp
    if ($h -is [System.Management.Automation.PSMethod]) {
        return [long]$h.Invoke(@())
    }
    return [long]$h
}

function Set-AccessVisibleBestEffort {
    <#
    .SYNOPSIS
        Try to set Access visibility. Never fail startup if unsupported.
    #>
    param([bool]$Visible = $true)
    if ($null -eq $script:AccessSession.App) { return }
    try {
        $script:AccessSession.App.Visible = $Visible
    } catch {
        Write-Verbose "Could not set Access.Visible=$Visible (continuing): $_"
    }
}

function Clear-AccessCaches {
    <#
    .SYNOPSIS
        Clear all VBE/control/CodeModule caches.
    #>
    $script:AccessSession.VbeCodeCache  = @{}
    $script:AccessSession.ControlsCache = @{}
    $script:AccessSession.CmCache       = @{}
}

function Resolve-SessionDbPath {
    <#
    .SYNOPSIS
        Resolve -DbPath: use explicit value if given, else fall back to active session.
        Throws a terminating error if neither is available.
    #>
    param(
        [string]$DbPath,
        [string]$CallerName = 'AccessPOSH'
    )
    if ($DbPath) { return $DbPath }
    if ($script:AccessSession.DbPath) { return $script:AccessSession.DbPath }
    throw "${CallerName}: -DbPath is required (no active session). Open a database first."
}

function Connect-AccessDB {
    <#
    .SYNOPSIS
        Internal: ensure Access COM is running and the requested DB is open.
        Returns the COM Application object.
    #>
    param(
        [string]$DbPath
    )
    if (-not $DbPath) { throw "Connect-AccessDB: -DbPath is required." }

    $resolved = [System.IO.Path]::GetFullPath($DbPath)

    # If we have an existing session, check liveness
    if ($null -ne $script:AccessSession.App) {
        if (-not (Test-AccessAlive)) {
            Write-Verbose 'COM session stale — auto-reconnecting...'
            $script:AccessSession.App    = $null
            $script:AccessSession.DbPath = $null
            Clear-AccessCaches
        }
    }

    # Launch Access if needed
    if ($null -eq $script:AccessSession.App) {
        Write-Verbose 'Launching Access.Application...'
        try {
            $script:AccessSession.App = New-Object -ComObject 'Access.Application'
        } catch {
            throw "Failed to create Access.Application COM object. Is Microsoft Access installed? Error: $_"
        }
        # Suppress dialogs for non-interactive automation
        try {
            $script:AccessSession.App.DisplayAlerts = $false
            $script:AccessSession.App.AutomationSecurity = 1  # msoAutomationSecurityForceDisable
        } catch {}
        Set-AccessVisibleBestEffort -Visible $true
        Write-Verbose 'Access launched OK'
    }

    # Switch database if needed
    if ($script:AccessSession.DbPath -ne $resolved) {
        if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
            throw "Database file not found: $resolved"
        }

        # Close previous database
        if ($null -ne $script:AccessSession.DbPath) {
            Write-Verbose "Closing previous DB: $($script:AccessSession.DbPath)"
            try {
                $script:AccessSession.App.CloseCurrentDatabase()
            } catch {
                Write-Verbose "Error closing previous DB: $_"
            }
        }

        # Open new database
        Write-Verbose "Opening DB: $resolved"
        try {
            $script:AccessSession.App.OpenCurrentDatabase($resolved)
        } catch {
            if ($_.Exception.Message -match 'already have the database open') {
                Write-Verbose 'DB was already open — syncing state'
            } else {
                throw
            }
        }

        $script:AccessSession.DbPath = $resolved
        Set-AccessVisibleBestEffort -Visible $true
        Clear-AccessCaches
        Write-Verbose 'DB opened OK'
    }

    return $script:AccessSession.App
}
