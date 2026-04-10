# Public/UIAutomation.ps1 — Win32/GDI screenshot, mouse, keyboard automation

# ── AccessPoshUI native interop (loaded once) ─────────────────────────────
if (-not ([System.Management.Automation.PSTypeName]'AccessPoshUI').Type) {
    Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class AccessPoshUI
{
    [DllImport("user32.dll")]
    public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hwnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hwnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern short VkKeyScanW(char ch);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    public static extern IntPtr GetWindowDC(IntPtr hwnd);

    [DllImport("user32.dll")]
    public static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int width, int height);

    [DllImport("gdi32.dll")]
    public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hObject);

    [DllImport("gdi32.dll")]
    public static extern bool DeleteObject(IntPtr hObject);

    [DllImport("gdi32.dll")]
    public static extern bool DeleteDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    public static extern bool BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest,
                                     IntPtr hdcSrc, int xSrc, int ySrc, uint rop);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    // Constants
    public const uint PW_RENDERFULLCONTENT = 2;
    public const int  SW_RESTORE           = 9;
    public const uint SRCCOPY              = 0x00CC0020;
}
'@
}

function Get-AccessScreenshot {
    <#
    .SYNOPSIS
        Capture a screenshot of the Access window and save as PNG.
    .DESCRIPTION
        Optionally opens a form or report first, captures the Access window
        via PrintWindow (PW_RENDERFULLCONTENT), optionally scales down to
        MaxWidth, and saves as PNG. Returns path, dimensions, and file size.
    .EXAMPLE
        Get-AccessScreenshot -DbPath C:\my.accdb -AsJson
    .EXAMPLE
        Get-AccessScreenshot -DbPath C:\my.accdb -ObjectType form -ObjectName MainMenu -MaxWidth 1024
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,

        [ValidateSet('form','report')]
        [string]$ObjectType,

        [string]$ObjectName,

        [string]$OutputPath,

        [int]$WaitMs = 300,

        [int]$MaxWidth = 1920,

        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessScreenshot'

    $app = Connect-AccessDB -DbPath $DbPath

    $weOpened = $false
    try {
        # Optionally open a form or report
        if ($ObjectType -and $ObjectName) {
            switch ($ObjectType) {
                'form'   { $app.DoCmd.OpenForm($ObjectName)   }
                'report' { $app.DoCmd.OpenReport($ObjectName, 2 <# acViewPreview #>) }
            }
            $weOpened = $true
            Start-Sleep -Milliseconds $WaitMs
        }

        # Get window handle
        $hwnd = Get-AccessHwnd -App $app
        if ($hwnd -eq 0) {
            throw 'Could not obtain Access window handle.'
        }
        $hWndPtr = [IntPtr]::new($hwnd)

        # Restore if minimized
        if ([AccessPoshUI]::IsIconic($hWndPtr)) {
            [AccessPoshUI]::ShowWindow($hWndPtr, [AccessPoshUI]::SW_RESTORE) | Out-Null
            Start-Sleep -Milliseconds 200
        }

        # Bring to foreground for reliable capture
        [AccessPoshUI]::SetForegroundWindow($hWndPtr) | Out-Null
        Start-Sleep -Milliseconds 100

        # Get window dimensions
        $rect = New-Object AccessPoshUI+RECT
        [AccessPoshUI]::GetWindowRect($hWndPtr, [ref]$rect) | Out-Null
        $w = $rect.Right  - $rect.Left
        $h = $rect.Bottom - $rect.Top
        if ($w -le 0 -or $h -le 0) {
            throw "Invalid window dimensions: ${w}x${h}"
        }

        $origW = $w
        $origH = $h

        # Capture via PrintWindow into a System.Drawing.Bitmap
        $bmp = [System.Drawing.Bitmap]::new($w, $h)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $hdc = $g.GetHdc()
        try {
            [AccessPoshUI]::PrintWindow($hWndPtr, $hdc, [AccessPoshUI]::PW_RENDERFULLCONTENT) | Out-Null
        } finally {
            $g.ReleaseHdc($hdc)
        }

        # Resize if wider than MaxWidth
        if ($w -gt $MaxWidth) {
            $ratio  = $MaxWidth / $w
            $newH   = [int]($h * $ratio)
            $resized = [System.Drawing.Bitmap]::new($bmp, $MaxWidth, $newH)
            $bmp.Dispose()
            $bmp = $resized
            $w = $MaxWidth
            $h = $newH
        }

        # Determine output path
        if (-not $OutputPath) {
            $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
            $OutputPath = [System.IO.Path]::Combine(
                [System.IO.Path]::GetTempPath(),
                "access_screenshot_${stamp}.png"
            )
        }

        # Save PNG
        $bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $fileSize = (Get-Item -LiteralPath $OutputPath).Length

        $result = [ordered]@{
            status              = 'captured'
            path                = $OutputPath
            width               = $w
            height              = $h
            original_width      = $origW
            original_height     = $origH
            file_size           = $fileSize
        }
        if ($ObjectType -and $ObjectName) {
            $result['object_type'] = $ObjectType
            $result['object_name'] = $ObjectName
        }

        Format-AccessOutput -AsJson:$AsJson -Data $result
    } catch {
        $err = [ordered]@{ status = 'error'; error = $_.Exception.Message }
        Format-AccessOutput -AsJson:$AsJson -Data $err
    } finally {
        # Clean up GDI resources
        if ($null -ne $bmp) { $bmp.Dispose() }
        if ($null -ne $g)   { $g.Dispose()   }

        # Close the form/report if we opened it
        if ($weOpened -and $ObjectType -and $ObjectName) {
            try {
                switch ($ObjectType) {
                    'form'   { $app.DoCmd.Close(2, $ObjectName, 1 <# acSaveNo #>) }
                    'report' { $app.DoCmd.Close(3, $ObjectName, 1 <# acSaveNo #>) }
                }
            } catch {
                Write-Verbose "Could not close $ObjectType '$ObjectName': $_"
            }
        }
    }
}

function Send-AccessClick {
    <#
    .SYNOPSIS
        Send a mouse click to the Access window at image-relative coordinates.
    .DESCRIPTION
        Scales X/Y from reference image coordinates to actual screen coordinates
        using the Access window rect and ImageWidth, then performs left, double,
        or right click via mouse_event.
    .EXAMPLE
        Send-AccessClick -DbPath C:\my.accdb -X 150 -Y 200 -ImageWidth 1024
    .EXAMPLE
        Send-AccessClick -DbPath C:\my.accdb -X 300 -Y 50 -ImageWidth 1920 -ClickType double
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,

        [int]$X,

        [int]$Y,

        [int]$ImageWidth,

        [ValidateSet('left','double','right')]
        [string]$ClickType = 'left',

        [int]$WaitAfterMs = 200,

        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Send-AccessClick'
    if (-not $PSBoundParameters.ContainsKey('X')) { throw "Send-AccessClick: -X is required." }
    if (-not $PSBoundParameters.ContainsKey('Y')) { throw "Send-AccessClick: -Y is required." }
    if (-not $PSBoundParameters.ContainsKey('ImageWidth')) { throw "Send-AccessClick: -ImageWidth is required." }

    # mouse_event flag constants
    $LEFTDOWN  = [uint32]0x0002
    $LEFTUP    = [uint32]0x0004
    $RIGHTDOWN = [uint32]0x0008
    $RIGHTUP   = [uint32]0x0010

    $app = Connect-AccessDB -DbPath $DbPath

    try {
        $hwnd = Get-AccessHwnd -App $app
        if ($hwnd -eq 0) { throw 'Could not obtain Access window handle.' }
        $hWndPtr = [IntPtr]::new($hwnd)

        # Restore if minimized
        if ([AccessPoshUI]::IsIconic($hWndPtr)) {
            [AccessPoshUI]::ShowWindow($hWndPtr, [AccessPoshUI]::SW_RESTORE) | Out-Null
            Start-Sleep -Milliseconds 200
        }

        [AccessPoshUI]::SetForegroundWindow($hWndPtr) | Out-Null
        Start-Sleep -Milliseconds 50

        # Get window rect and compute scale
        $rect = New-Object AccessPoshUI+RECT
        [AccessPoshUI]::GetWindowRect($hWndPtr, [ref]$rect) | Out-Null
        $winW = $rect.Right - $rect.Left
        if ($winW -le 0) { throw "Invalid window width: $winW" }

        $scale   = $winW / $ImageWidth
        $screenX = [int]($rect.Left + $X * $scale)
        $screenY = [int]($rect.Top  + $Y * $scale)

        # Move cursor
        [AccessPoshUI]::SetCursorPos($screenX, $screenY) | Out-Null
        Start-Sleep -Milliseconds 30

        # Perform click
        switch ($ClickType) {
            'left' {
                [AccessPoshUI]::mouse_event($LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
                [AccessPoshUI]::mouse_event($LEFTUP,   0, 0, 0, [UIntPtr]::Zero)
            }
            'double' {
                [AccessPoshUI]::mouse_event($LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
                [AccessPoshUI]::mouse_event($LEFTUP,   0, 0, 0, [UIntPtr]::Zero)
                Start-Sleep -Milliseconds 50
                [AccessPoshUI]::mouse_event($LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
                [AccessPoshUI]::mouse_event($LEFTUP,   0, 0, 0, [UIntPtr]::Zero)
            }
            'right' {
                [AccessPoshUI]::mouse_event($RIGHTDOWN, 0, 0, 0, [UIntPtr]::Zero)
                [AccessPoshUI]::mouse_event($RIGHTUP,   0, 0, 0, [UIntPtr]::Zero)
            }
        }

        Start-Sleep -Milliseconds $WaitAfterMs

        Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            status     = 'clicked'
            screen_x   = $screenX
            screen_y   = $screenY
            image_x    = $X
            image_y    = $Y
            click_type = $ClickType
            scale      = [math]::Round($scale, 4)
        })
    } catch {
        Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            status = 'error'
            error  = $_.Exception.Message
        })
    }
}

function Send-AccessKeyboard {
    <#
    .SYNOPSIS
        Send keyboard input (text or special keys) to the Access window.
    .DESCRIPTION
        Types text via WM_CHAR SendMessage, or sends special-key combos with
        optional modifiers (ctrl, shift, alt) via keybd_event.
    .EXAMPLE
        Send-AccessKeyboard -DbPath C:\my.accdb -Text "Hello World"
    .EXAMPLE
        Send-AccessKeyboard -DbPath C:\my.accdb -Key enter
    .EXAMPLE
        Send-AccessKeyboard -DbPath C:\my.accdb -Key "s" -Modifiers "ctrl"
    .EXAMPLE
        Send-AccessKeyboard -DbPath C:\my.accdb -Key "a" -Modifiers "ctrl+shift"
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,

        [string]$Text,

        [string]$Key,

        [string]$Modifiers,

        [int]$WaitAfterMs = 100,

        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Send-AccessKeyboard'

    # Virtual-key code map for special keys
    $VK_MAP = @{
        enter     = 0x0D; tab       = 0x09; escape    = 0x1B
        backspace = 0x08; delete    = 0x2E; space     = 0x20
        up        = 0x26; down      = 0x28; left      = 0x25; right = 0x27
        home      = 0x24; 'end'     = 0x23
        pageup    = 0x21; pagedown  = 0x22
        f1  = 0x70; f2  = 0x71; f3  = 0x72;  f4  = 0x73
        f5  = 0x74; f6  = 0x75; f7  = 0x76;  f8  = 0x77
        f9  = 0x78; f10 = 0x79; f11 = 0x7A;  f12 = 0x7B
    }

    # Modifier name -> virtual-key code
    $MOD_MAP = @{
        ctrl  = 0x11
        shift = 0x10
        alt   = 0x12
    }

    $KEYEVENTF_KEYUP = [uint32]2

    if (-not $Text -and -not $Key) {
        throw 'At least one of -Text or -Key must be specified.'
    }

    $app = Connect-AccessDB -DbPath $DbPath

    try {
        $hwnd = Get-AccessHwnd -App $app
        if ($hwnd -eq 0) { throw 'Could not obtain Access window handle.' }
        $hWndPtr = [IntPtr]::new($hwnd)

        # Restore if minimized
        if ([AccessPoshUI]::IsIconic($hWndPtr)) {
            [AccessPoshUI]::ShowWindow($hWndPtr, [AccessPoshUI]::SW_RESTORE) | Out-Null
            Start-Sleep -Milliseconds 200
        }

        [AccessPoshUI]::SetForegroundWindow($hWndPtr) | Out-Null
        Start-Sleep -Milliseconds 50

        $action = ''

        # -- Type text via WM_CHAR --
        if ($Text) {
            foreach ($ch in $Text.ToCharArray()) {
                [AccessPoshUI]::SendMessage($hWndPtr, 0x0102, [IntPtr]::new([int][char]$ch), [IntPtr]::Zero) | Out-Null
            }
            $action = "typed $($Text.Length) character(s)"
        }

        # -- Send special key / key combo --
        if ($Key) {
            $keyLower = $Key.ToLower()

            # Resolve virtual key code
            if ($VK_MAP.ContainsKey($keyLower)) {
                $vk = [byte]$VK_MAP[$keyLower]
            } else {
                # Single character — use VkKeyScanW
                if ($Key.Length -eq 1) {
                    $scan = [AccessPoshUI]::VkKeyScanW([char]$Key)
                    $vk = [byte]($scan -band 0xFF)
                } else {
                    throw "Unknown key name: '$Key'. Use a VK_MAP name or a single character."
                }
            }

            # Parse modifier string (e.g. "ctrl+shift")
            $modVks = @()
            if ($Modifiers) {
                foreach ($m in ($Modifiers.ToLower() -split '\+')) {
                    $m = $m.Trim()
                    if (-not $MOD_MAP.ContainsKey($m)) {
                        throw "Unknown modifier: '$m'. Use ctrl, shift, or alt."
                    }
                    $modVks += [byte]$MOD_MAP[$m]
                }
            }

            # Press modifiers down
            foreach ($mv in $modVks) {
                [AccessPoshUI]::keybd_event($mv, 0, 0, [UIntPtr]::Zero)
            }

            # Press and release the main key
            [AccessPoshUI]::keybd_event($vk, 0, 0, [UIntPtr]::Zero)
            [AccessPoshUI]::keybd_event($vk, 0, $KEYEVENTF_KEYUP, [UIntPtr]::Zero)

            # Release modifiers in reverse order
            for ($i = $modVks.Count - 1; $i -ge 0; $i--) {
                [AccessPoshUI]::keybd_event($modVks[$i], 0, $KEYEVENTF_KEYUP, [UIntPtr]::Zero)
            }

            $keyDesc = if ($Modifiers) { "$Modifiers+$Key" } else { $Key }
            $action = if ($action) { "$action; sent key $keyDesc" } else { "sent key $keyDesc" }
        }

        Start-Sleep -Milliseconds $WaitAfterMs

        $result = [ordered]@{
            status    = 'sent'
            action    = $action
        }
        if ($Modifiers) { $result['modifiers'] = $Modifiers }
        if ($Key)       { $result['key']       = $Key }
        if ($Text)      { $result['text_length'] = $Text.Length }

        Format-AccessOutput -AsJson:$AsJson -Data $result
    } catch {
        Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            status = 'error'
            error  = $_.Exception.Message
        })
    }
}
