<# ::
    cls & @echo off & chcp 437 >nul & title PowerShell LaunchBar

    if /i "%~1"=="/?"       goto :help
    if /i "%~1"=="-?"       goto :help
    if /i "%~1"=="--?"      goto :help
    if /i "%~1"=="/help"    goto :help
    if /i "%~1"=="-help"    goto :help
    if /i "%~1"=="--help"   goto :help

    set "WindowStyle=Hidden"
    for %%A in (%*) do if /I "%%A"=="-showdebug" set "WindowStyle=Normal"

    copy /y "%~f0" "%TEMP%\%~n0.ps1" >NUL && powershell -Nologo -NoProfile -ExecutionPolicy Bypass -WindowStyle %WindowStyle% -File "%TEMP%\%~n0.ps1" %*
    exit /b

    :help
    mode con: cols=128 lines=60
    echo.
    echo.
    echo    =============================================================================
    echo                               PowerShell LaunchBar v1.00
    echo                                           ---
    echo                        Author : Leo Gillet - Freenitial on GitHub
    echo    =============================================================================
    echo.
    echo.
    echo    DESCRIPTION:
    echo       -----------
    echo       Launches a customizable PowerShell toolbar that facilitates
    echo       the quick execution of applications, scripts, or shortcuts.
    echo.
    echo       The toolbar runs as a borderless window positioned at the top or bottom
    echo       of the screen.
    echo.
    echo       It allows you to import and export shortcuts and settings via INI files,
    echo       so you can tailor its behavior and appearance to your needs.
    echo.
    echo       To add shortcuts, drag and drop files on the launchbar
    echo       Right-click on shortcuts to see more options
    echo.
    echo.
    echo    OPTIONAL ARGUMENTS:
    echo       --------------------
    echo       1) "Full\Filepath\of\shortcuts.ini"
    echo          - Full path to the INI file containing shortcuts to add.
    echo.
    echo       2) "Full\Filepath\of\settings.ini"
    echo          - Full path to the INI file used to configure the toolbar.
    echo.
    echo       3) -force
    echo          - Forces the import of the shortcuts.ini file without prompting for confirmation.
    echo.
    echo       4) -showdebug
    echo          - Displays the debug console to show execution messages.
    echo.
    echo       5) -nolog
    echo          - Disables the creation of logs in "%localappdata%\PowerShell_LaunchBar\Logs".
    echo.
    echo.
    echo    USAGE:
    echo       ------
    echo       To launch PowerShell_LaunchBar from another batch or cmd without closing console:
    echo          start "" /d "FOLDER_PATH" PowerShell_LaunchBar
    echo.
    echo       To launch with forced INI file import and debug console display:
    echo          start "" /d "FOLDER_PATH" PowerShell_LaunchBar "FULLPATH\SHORTCUTS.INI" "FULLPATH\SETTINGS.INI" -force -showdebug
    echo.
    echo       Multi-line example:
    echo          start "" /d "FOLDER_PATH"   PowerShell_LaunchBar ^^
    echo                                      "FULLPATH\SHORTCUTS.INI" ^^
    echo                                      "FULLPATH\SETTINGS.INI" ^^
    echo                                      -force -showdebug
    echo.
    echo.
    echo    =============================================================================
    echo.
    pause>nul & exit /b
#>


[CmdletBinding()]
param(
    [Parameter(Position=0, ValueFromRemainingArguments = $true)]
    [string[]]$IniFiles = @(),

    [switch]$force,
    [switch]$showdebug,
    [switch]$nolog
)


Add-Type -AssemblyName System.Drawing, System.Windows.Forms

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;
public static class DPIHelper {
    public static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);
    [DllImport("user32.dll")]
    public static extern bool SetProcessDpiAwarenessContext(IntPtr dpiFlag);
}
public static class NativeMethods {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern uint RegisterWindowMessage(string lpString);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
}
public class AppBar {
    public const int ABM_NEW = 0;
    public const int ABM_REMOVE = 1;
    public const int ABM_SETPOS = 2;
    public const int ABM_QUERYPOS = 3;
    public const int ABE_TOP = 1;
    public const int ABE_BOTTOM = 3;
    public const int SPI_SETWORKAREA = 47;
    public const int SPI_GETWORKAREA = 48;
    public const int SPIF_UPDATEINIFILE = 0x01;
    [StructLayout(LayoutKind.Sequential)]
    public struct APPBARDATA {
        public int cbSize;
        public IntPtr hWnd;
        public int uCallbackMessage;
        public int uEdge;
        public RECT rc;
        public IntPtr lParam;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int left, top, right, bottom; }
    [DllImport("shell32.dll", CallingConvention = CallingConvention.StdCall)]
    public static extern uint SHAppBarMessage(uint dwMessage, ref APPBARDATA pData);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, ref RECT pvParam, uint fWinIni);
}
public class IconExtractor {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct SHFILEINFO {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    };
    public const uint SHGFI_ICON = 0x100;
    public const uint SHGFI_SMALLICON = 0x1;
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);
    public static Icon GetIcon(string fileName) {
        SHFILEINFO shinfo = new SHFILEINFO();
        IntPtr hImg = SHGetFileInfo(fileName, 0, ref shinfo, (uint)System.Runtime.InteropServices.Marshal.SizeOf(shinfo), SHGFI_ICON | SHGFI_SMALLICON);
        if (shinfo.hIcon != IntPtr.Zero) { Icon icon = Icon.FromHandle(shinfo.hIcon); return (Icon)icon.Clone(); }
        return null;
    }
}
"@ -ReferencedAssemblies @("System.Drawing", "System.Runtime.InteropServices")

$WM_MOUSELEAVE = 0x02A3
[DPIHelper]::SetProcessDpiAwarenessContext([DPIHelper]::DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2) | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()


$savePath = Join-Path $env:LOCALAPPDATA "Powershell_Launchbar"
$logsPath  = Join-Path $env:LOCALAPPDATA "Powershell_Launchbar\Logs"
$shortcutsFile = Join-Path $savePath "shortcuts.ini"
$settingsFile  = Join-Path $savePath "settings.ini"
if (-not (Test-Path -LiteralPath $savePath)) { New-Item -ItemType Directory -Path $savePath | Out-Null }
if (-not (Test-Path -LiteralPath $logsPath)) { New-Item -ItemType Directory -Path $logsPath | Out-Null }
$logFile = Join-Path $logsPath ("{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
if (-not (Test-Path -LiteralPath $logFile)) { New-Item -ItemType File -Path $logFile | Out-Null }
$logFiles = Get-ChildItem -Path $logsPath -File | Sort-Object Name -Descending
if ($logFiles.Count -gt 20) { $logFiles | Select-Object -Skip 20 | Remove-Item -Force }

function Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$message
    )
    if ($showdebug.IsPresent) { Write-Host $message }
    if (-not $nolog.IsPresent) {
        $message = "[$('{0:yyyy/MM/dd - HH:mm:ss}' -f (Get-Date))] - $message"
        Add-Content -Path $logFile -Value $message
    }
}

log "Init Global variables and paths"
log "PSCommandPath = $PSCommandPath"

$global:Settings = @{
    ToolbarLocation                = "Top"
    ThicknessMode                  = "Small"
    Theme                          = "Light"
    NewShortcutShowText            = "true"
    NewShortcutOpenAsAdmin         = "false"
    NewShortcutAlignRight          = "false"
    NewShortcutAdminAccessFailAlternative = "Ask"
}
function Read-IniFile($Path){
    $ini=[ordered]@{}
    if(Test-Path -LiteralPath $Path){
        $section=""
        foreach($line in Get-Content $Path){
            $line=$line.Trim()
            if ($line -match '^\[(.+)\]') { $section = $Matches[1]; $ini[$section] = [ordered]@{}; continue }
            if($line -match '^(.*?)=(.*)$'){
                $key=$Matches[1].Trim(); $value=$Matches[2].Trim()
                if ($section) { $ini[$section][$key] = $value } else { $ini[$key] = $value }
            }
        }
    }
    return $ini
}

$global:DebounceActive = $false
$global:BaselineWorkArea = $null
$global:LastDisplaySettingsTime = Get-Date
$global:CurrentTooltipControl = $null
$canRequireAdminExtensions = @(".exe",".bat",".cmd",".ps1",".msc",".msi",".msp",".vbs",".vbe",".js",".jse",".wsf",".wsh",".cpl",".reg", ".lnk")

function Write-IniFile($Path, $Data) {
    log "Write IniFile begin..."
    $lines = @()
    foreach ($section in $Data.Keys) {
        if ($Data[$section] -is [System.Collections.IDictionary]) {
            $lines += "[$section]"
            foreach ($k in $Data[$section].Keys) { 
                $lines += "$k=$($Data[$section][$k])" 
            }
        } else { 
            $lines += "$section=$($Data[$section])" 
        }
    }
    Set-Content -Path $Path -Value $lines
    log "Write IniFile end - OK"
}

function Save-Shortcuts {
    log "Saving shortcuts begin..."
    $data = [ordered]@{}; $index = 0
    foreach ($ctrl in $shortcutsPanel.Controls) {
        if ($ctrl -is [System.Windows.Forms.Button] -and $ctrl.Tag -and $ctrl.Tag.ContainsKey("FilePath")) {
            $data["Shortcut$index"] = [ordered]@{
                Order                        = $ctrl.Tag.Order
                Path                         = $ctrl.Tag.FilePath
                DisplayName                  = $ctrl.Tag.DisplayName
                ShowText                     = $ctrl.Tag.ShowText
                AlignRight                   = $ctrl.Tag.AlignRight
                OpenAsAdmin                  = $ctrl.Tag.OpenAsAdmin
                AdminAccessFailAlternative   = $ctrl.Tag.AdminAccessFailAlternative
            }
            $index++
        }
    }
    Write-IniFile $shortcutsFile $data
    log "Saved shortcuts end - OK"
}

function Save-Settings { log "Saving settings begin..."; Write-IniFile $settingsFile @{ Settings = $global:Settings }; log "Saved settings end - OK" }


log "Create main Windows Forms" 
$form = New-Object System.Windows.Forms.Form
$form.SuspendLayout()
$form.FormBorderStyle = 'None'
$form.TopMost = $true
$form.ShowInTaskbar = $false
$form.StartPosition = 'Manual'
$form.Height = $global:BarThickness

$shortcutsPanel = New-Object System.Windows.Forms.Panel
$shortcutsPanel.Dock = 'Fill'
$shortcutsPanel.BackColor = $form.BackColor
$shortcutsPanel.AllowDrop = $true

$settingsButton = New-Object System.Windows.Forms.Button
$settingsButton.Text = $([char]0x26EF)
$settingsButton.TextAlign = 'MiddleCenter'
$settingsButton.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 11)
$settingsButton.Width = 30
$settingsButton.Dock = 'Right'
$settingsButton.FlatStyle = 'Flat'
$settingsButton.BackColor = $form.BackColor

$form.Controls.Add($shortcutsPanel)
$form.Controls.Add($settingsButton)

$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.ShowAlways = $true
$tooltip.AutoPopDelay = 5000
$tooltip.InitialDelay = 50
$tooltip.ReshowDelay  = 50
$tooltipFont = New-Object System.Drawing.Font("Microsoft Sans Serif", 14)
$tooltip.OwnerDraw = $true
$tooltip.add_Draw({
    param($s, $e)
    $e.Graphics.FillRectangle([System.Drawing.Brushes]::Black, $e.Bounds)
    $e.Graphics.DrawRectangle([System.Drawing.Pens]::White, [System.Drawing.Rectangle]::new($e.Bounds.X, $e.Bounds.Y, $e.Bounds.Width - 1, $e.Bounds.Height - 1))
    [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, $e.ToolTipText, $tooltipFont, $e.Bounds, [System.Drawing.Color]::White)
})
$tooltip.add_Popup({param($s, $e); $e.ToolTipSize = [System.Windows.Forms.TextRenderer]::MeasureText($tooltip.GetToolTip($e.AssociatedControl), $tooltipFont) })

$dragIndicator = New-Object System.Windows.Forms.Panel
$dragIndicator.BackColor = [System.Drawing.Color]::Orange
$dragIndicator.Width = 5
$dragIndicator.Height = $global:BarThickness
$dragIndicator.Visible = $false
$shortcutsPanel.Controls.Add($dragIndicator)
$dragIndicator.BringToFront()

# --- Update AppBar position and working area ---
function Update-AppBarPosition {
    param([string]$position)
    log "Updating AppBarPosition begin..."
    $handle = $form.Handle
    $appBarData = New-Object AppBar+APPBARDATA
    $appBarData.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf($appBarData)
    $appBarData.hWnd = $handle
    $appBarData.uCallbackMessage = [NativeMethods]::RegisterWindowMessage("AppBarMessage")
    [AppBar]::SHAppBarMessage([AppBar]::ABM_REMOVE, [ref]$appBarData) | Out-Null
    $currentWorkArea = New-Object AppBar+RECT
    [AppBar]::SystemParametersInfo([AppBar]::SPI_GETWORKAREA, 0, [ref]$currentWorkArea, 0) | Out-Null
    if (-not $global:BaselineWorkArea -or (($global:BaselineWorkArea.right - $global:BaselineWorkArea.left) -ne ($currentWorkArea.right - $currentWorkArea.left))) {
        $global:BaselineWorkArea = New-Object AppBar+RECT -Property @{ left=$currentWorkArea.left; top=$currentWorkArea.top; right=$currentWorkArea.right; bottom=$currentWorkArea.bottom }
    }
    $form.Left = $global:BaselineWorkArea.left
    $form.Width = $global:BaselineWorkArea.right - $global:BaselineWorkArea.left
    $form.Height = $global:BarThickness
    switch ($position) {
        "Top" {
            $edge = [AppBar]::ABE_TOP
            $form.Top = $global:BaselineWorkArea.top
            $workTop = $global:BaselineWorkArea.top + $global:BarThickness
            $workBottom = $global:BaselineWorkArea.bottom
        }
        "Bottom" {
            $edge = [AppBar]::ABE_BOTTOM
            $form.Top = $global:BaselineWorkArea.bottom - $global:BarThickness
            $workTop = $global:BaselineWorkArea.top
            $workBottom = $global:BaselineWorkArea.bottom - $global:BarThickness
        }
    }
    $appBarData.uEdge = $edge
    $appBarData.rc = New-Object AppBar+RECT -Property @{ left=$form.Left; top=$form.Top; right=$form.Left+$form.Width; bottom=$form.Top+$form.Height }
    [AppBar]::SHAppBarMessage([AppBar]::ABM_NEW, [ref]$appBarData) | Out-Null
    [AppBar]::SHAppBarMessage([AppBar]::ABM_SETPOS, [ref]$appBarData) | Out-Null
    $newWorkArea = New-Object AppBar+RECT -Property @{ left=$global:BaselineWorkArea.left; top=$workTop; right=$global:BaselineWorkArea.right; bottom=$workBottom }
    [AppBar]::SystemParametersInfo([AppBar]::SPI_SETWORKAREA, 0, [ref]$newWorkArea, [AppBar]::SPIF_UPDATEINIFILE) | Out-Null
    Start-Sleep -Milliseconds 200
    log "Updated AppBarPosition end - OK"
}

# --- Update layout and appearance ---
function Update-Layout {
    log "Updating Layout begin..."
    switch ($global:Settings["ThicknessMode"]) {
        "Small"  { $iconSize = 14; }
        "Medium" { $iconSize = 22; }
        "Large"  { $iconSize = 28; }
    }
    $themeColors =  if ($global:Settings["Theme"] -eq "Dark") { @{ BackColor = [System.Drawing.Color]::FromArgb(51,51,51); ForeColor = [System.Drawing.Color]::White; BorderColor = [System.Drawing.Color]::DimGray } }
                    else { @{ BackColor = [System.Drawing.Color]::FromArgb(238,238,238); ForeColor = [System.Drawing.Color]::Black; BorderColor = [System.Drawing.Color]::Gray } }
    $shortcutsPanel.BackColor = $themeColors.BackColor
    $settingsButton.ForeColor = $themeColors.ForeColor
    $settingsButton.FlatAppearance.BorderColor = $themeColors.BorderColor
    $settingsButton.BackColor = $shortcutsPanel.BackColor
    $dragIndicator.Height = $global:BarThickness
    $leftButtons = New-Object 'System.Collections.Generic.List[System.Windows.Forms.Button]'
    $rightButtons = New-Object 'System.Collections.Generic.List[System.Windows.Forms.Button]'
    $g = [System.Drawing.Graphics]::FromImage((New-Object System.Drawing.Bitmap(1,1)))
    foreach ($btn in $shortcutsPanel.Controls) {
        if ($btn -is [System.Windows.Forms.Button] -and $btn.Tag -and $btn.Tag.ContainsKey("FilePath")) {
            switch ($global:Settings["Theme"]) {
                "Dark"  { $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::DimGray ; $btn.ForeColor = [System.Drawing.Color]::White ; $btn.BackColor = [System.Drawing.Color]::DarkSlateGray }
                default { $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray ; $btn.ForeColor = [System.Drawing.Color]::Black ; $btn.BackColor = [System.Drawing.Color]::White }
            }
            if ($btn.Tag["DisplayName"] -ne "") { $displayName = $btn.Tag["DisplayName"] } else { $displayName = $btn.Tag["FilePath"] }
            if ($btn.Tag["ShowText"] -eq "true") { $btn.Text = " $displayName" ; $tooltip.SetToolTip($btn, $null)}
            else { $btn.Text = "" ; $tooltip.SetToolTip($btn, $displayName) }
            $btn.Height = $global:BarThickness
            $btn.AutoEllipsis = $true
            $btn.TextImageRelation = "ImageBeforeText"
            $btn.ImageAlign = "MiddleLeft"
            $btn.TextAlign = "MiddleLeft"
            if ($btn.Tag.ContainsKey("OriginalImage") -and $btn.Tag["OriginalImage"]) { try { $btn.Image = $btn.Tag["OriginalImage"].GetThumbnailImage($iconSize, $iconSize, $null, [IntPtr]::Zero) } catch { }}
            $txtWidth = [Math]::Ceiling($g.MeasureString($btn.Text, $btn.Font).Width)
            if ($btn.Tag["ShowText"] -eq "true") { $padding = 10 } else { $padding = 13 }
            $btn.Width = $iconSize + $txtWidth + $padding
            if ($btn.Tag["AlignRight"] -eq "true") { $rightButtons.Add($btn) } else { $leftButtons.Add($btn) }
        }
    }
    # Overflow handling: proportionally shrink text buttons if total width exceeds panel width
    $spacing = 2
    $panelWidth = $shortcutsPanel.ClientSize.Width
    $buttonList =   if ($leftButtons.Count -gt 0 -and $rightButtons.Count -gt 0) { $leftButtons + $rightButtons } 
                    elseif ($leftButtons.Count -gt 0) { $leftButtons } 
                    elseif ($rightButtons.Count -gt 0) { $rightButtons } 
                    else { @() }
    if ($buttonList.Count) {
        $totalSpacing = if ($leftButtons.Count -gt 0 -and $rightButtons.Count -gt 0) { ($buttonList.Count - 2) * $spacing } else { ($buttonList.Count - 1) * $spacing }
        $fixed = 0; $var = @()
        foreach ($btn in $buttonList) { if ($btn.Tag["ShowText"] -eq "true") { $var += $btn } else { $fixed += $btn.Width } }
        $origVar = ($var | Measure-Object -Property Width -Sum).Sum
        if (($fixed + $origVar + $totalSpacing) -gt $panelWidth -and $origVar) {
            $scale = ($panelWidth - $fixed - $totalSpacing) / $origVar
            foreach ($btn in $var) { $btn.Width = [Math]::Floor($btn.Width * $scale) }
        }
    }
    $cs = $shortcutsPanel.ClientSize; $leftX = 0; $rightX = $cs.Width
    foreach ($btn in $leftButtons) { $btn.Location = New-Object System.Drawing.Point($leftX, 0) ; $leftX += $btn.Width + $spacing }
    $rightSorted = $rightButtons | Sort-Object { [int]$_.Tag.Order } -Descending
    foreach ($btn in $rightSorted) { $rightX -= $btn.Width ; $btn.Location = New-Object System.Drawing.Point($rightX, 0) ; $rightX -= $spacing }
    $g.Dispose()
    Save-Settings
    log "Updating Layout end - OK"
}

function Add-ShortcutButton {
    param(
        [string]$FilePath,
        [switch]$NoSave,
        [string]$DefOpenAsAdmin = $global:Settings["NewShortcutOpenAsAdmin"],
        [string]$DefShowText    = $global:Settings["NewShortcutShowText"],
        [string]$DefAlignRight  = $global:Settings["NewShortcutAlignRight"],
        [string]$DefDisplayName = ""
    )
    <#
    if (-not (Test-Path -LiteralPath $FilePath)) {
        Log "Cannot add shortcut. File or folder '$FilePath' does not exist."
        return
    }
    #>

    log "Adding ShortcutButton begin '$FilePath'..."

    $isFolder = -not ($FilePath -match '\.[^\\]+$')
    if (-not $isFolder) { $ext = [System.IO.Path]::GetExtension($FilePath).ToLower() } else { $ext = "" }
    $canOpenAsAdmin = $canRequireAdminExtensions -contains $ext
    if (-not $canOpenAsAdmin) { $DefOpenAsAdmin = "false" }

    $icon = $null
    try { $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($FilePath) } catch { $icon = $null; log "Icon not found with method 1" }
    if (-not $icon) { try { $icon = [IconExtractor]::GetIcon($FilePath) } catch { log "Icon not found with method 2" } }
    $img = $null
    if ($icon) {
        try {
            $ms = New-Object System.IO.MemoryStream
            $icon.ToBitmap().Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
            $ms.Position = 0
            $bmp = New-Object System.Drawing.Bitmap($ms)
            $img = $bmp
        } catch { }
    }

    $btn = New-Object System.Windows.Forms.Button
    $btn.TabStop = $false
    $btn.FlatStyle = 'Flat'
    $btn.FlatAppearance.BorderSize = 1
    $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::DarkCyan
    $displayName = if ($DefDisplayName -ne "") { $DefDisplayName } else { Split-Path $FilePath -Leaf }
    $btn.Tag = @{
        FilePath                   = $FilePath;
        OpenAsAdmin                = $DefOpenAsAdmin;
        ShowText                   = $DefShowText;
        AlignRight                 = $DefAlignRight;
        DisplayName                = $displayName;
        AdminAccessFailAlternative = $global:Settings["NewShortcutAdminAccessFailAlternative"];
        DragStart                  = $null
    }
    if ($DefAlignRight -eq "true") { $group = $shortcutsPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.ContainsKey("FilePath") -and $_.Tag["AlignRight"] -eq "true"} }
    else { $group = $shortcutsPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.ContainsKey("FilePath") -and $_.Tag["AlignRight"] -ne "true"} }
    $btn.Tag.Order = $group.Count
    if ($img) { $btn.Tag["OriginalImage"] = $img }

    $btn.Add_Click({
        param($s, $e)
        $options = $s.Tag
        if ($options["FilePath"] -and (Test-Path -LiteralPath $options["FilePath"])) {
            if ($options["OpenAsAdmin"] -eq "true") {
                if (-not $script:AdminHelperStarted) { Start-AdminHelper }
                if ($options["AdminAccessFailAlternative"] -eq "Ask") { Invoke-AdminRun $options["FilePath"] } 
                else { Invoke-AdminRun $options["FilePath"] $options["AdminAccessFailAlternative"] }
            } 
            else { Start-Process -FilePath $options["FilePath"] }
        }
        else {
            $originalColor = $s.BackColor
            for ($i = 1; $i -le 6; $i++) {
                $s.BackColor = if ($s.BackColor -eq 'DarkRed') { $originalColor } else { 'DarkRed' }
                [NativeMethods]::SendMessage($s.Handle, $WM_MOUSELEAVE, [IntPtr]::Zero, [IntPtr]::Zero)
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 300
            }
            $s.BackColor = $originalColor
        }
    })

    # Create context menu and link the button in its Tag
    $ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $ctxMenu.Tag = $btn

    $miRemove = New-Object System.Windows.Forms.ToolStripMenuItem("Remove")
    $miRemove.Add_Click({
        param($s, $e)
        $localBtn = $s.Owner.Tag
        $form.SuspendLayout()
        $shortcutsPanel.Controls.Remove($localBtn)
        Save-Shortcuts
        Update-Layout
        $form.ResumeLayout()
    })

    $miRename = New-Object System.Windows.Forms.ToolStripMenuItem("Rename")
    $miRename.Add_Click({
        param($s, $e)
        $localBtn = $s.Owner.Tag
        $currentName = $localBtn.Tag["DisplayName"]
        $frm = New-Object System.Windows.Forms.Form
        $frm.Text = "Rename Shortcut"
        $frm.Size = New-Object System.Drawing.Size(300,150)
        $frm.StartPosition = 'CenterScreen'
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "New Name:"
        $lbl.Location = New-Object System.Drawing.Point(10,20)
        $lbl.AutoSize = $true
        $frm.Controls.Add($lbl)
        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Text = $currentName
        $txt.Location = New-Object System.Drawing.Point(80,18)
        $txt.Width = 180
        $frm.Controls.Add($txt)
        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "OK"
        $btnOK.Location = New-Object System.Drawing.Point(100,60)
        $btnOK.Add_Click({
            $frm.Tag = $txt.Text
            $frm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $frm.Close()
        })
        $frm.Controls.Add($btnOK)
        $frm.AcceptButton = $btnOK
        if ($frm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $form.SuspendLayout()
            $newName = $frm.Tag
            if ($newName.Trim() -ne "") {
                $localBtn.Tag["DisplayName"] = $newName.Trim()
                if ($localBtn.Tag["ShowText"] -eq "true") { $localBtn.Text = " " + $newName.Trim() }
                Save-Shortcuts
                Update-Layout
                $form.ResumeLayout()
            }
        }
        $frm.Dispose()
    })

    # Create dropdown "If admin access fail" with composite tags
    $miAdminFailAlternative = New-Object System.Windows.Forms.ToolStripMenuItem "If admin access fail"
    
    $miAsk = New-Object System.Windows.Forms.ToolStripMenuItem "Ask"
    $miAsk.CheckOnClick = $true
    $miAsk.Tag = @{ Value = "Ask"; Button = $btn }
    
    $miLocalCopy = New-Object System.Windows.Forms.ToolStripMenuItem "Local Copy"
    $miLocalCopy.CheckOnClick = $true
    $miLocalCopy.Tag = @{ Value = "LocalCopy"; Button = $btn }
    
    $miRunAsUser = New-Object System.Windows.Forms.ToolStripMenuItem "Run as User"
    $miRunAsUser.CheckOnClick = $true
    $miRunAsUser.Tag = @{ Value = "AsUser"; Button = $btn }
    
    $miCancel = New-Object System.Windows.Forms.ToolStripMenuItem "Cancel"
    $miCancel.CheckOnClick = $true
    $miCancel.Tag = @{ Value = "Cancel"; Button = $btn }
    
    $dropdownHandler = {
        param($s, $e)
        $s.Checked = $true
        $parent = $s.Owner
        $index = $parent.Items.IndexOf($s)
        for ($i = 0; $i -lt $parent.Items.Count; $i++) { if ($i -ne $index) { $parent.Items[$i].Checked = $false } }
        $menuData = $s.Tag
        if ($menuData -and $menuData.ContainsKey("Button")) {
            $btnLocal = $menuData.Button
            $btnLocal.Tag["AdminAccessFailAlternative"] = $menuData.Value
        }
        Save-Shortcuts
    }
    
    $miAsk.Add_Click($dropdownHandler)
    $miLocalCopy.Add_Click($dropdownHandler)
    $miRunAsUser.Add_Click($dropdownHandler)
    $miCancel.Add_Click($dropdownHandler)
    
    switch ($btn.Tag["AdminAccessFailAlternative"]) {
        "Ask"       { $miAsk.Checked = $true }
        "LocalCopy" { $miLocalCopy.Checked = $true }
        "AsUser"    { $miRunAsUser.Checked = $true }
        "Cancel"    { $miCancel.Checked = $true }
        default     { $miAsk.Checked = $true; $btn.Tag["AdminAccessFailAlternative"] = "Ask" }
    }
    $miAdminFailAlternative.DropDownItems.AddRange(@($miAsk, $miLocalCopy, $miRunAsUser, $miCancel))
    
    $miOpenAdmin = New-Object System.Windows.Forms.ToolStripMenuItem("Open as admin")
    $miOpenAdmin.CheckOnClick = $true
    $miOpenAdmin.Checked = ($DefOpenAsAdmin -eq "true")
    $miOpenAdmin.Tag = @{ Value = "OpenAsAdmin"; Button = $btn }
    if (-not $canOpenAsAdmin) { $miOpenAdmin.Enabled = $false } 
    else {
        $miOpenAdmin.Add_Click({
            param($s, $e)
            $comp = $s.Tag
            if ($comp -and $comp.ContainsKey("Button")) {
                $btnLocal = $comp.Button
                $btnLocal.Tag["OpenAsAdmin"] = if ($s.Checked) { "true" } else { "false" }
            }
            Save-Shortcuts
            $owner = $s.Owner
            if ($owner -and $owner.Items) {
                $adminFailItem = $owner.Items | Where-Object { $_.Text -eq "If admin access fail" }
                if ($null -ne $adminFailItem) { $adminFailItem.Enabled = $s.Checked }
            }
        })
    }
    
    $miShowText = New-Object System.Windows.Forms.ToolStripMenuItem("Show text")
    $miShowText.CheckOnClick = $true
    $miShowText.Checked = ($btn.Tag["ShowText"] -eq "true")
    $miShowText.Add_Click({
        param($s, $e)
        $localBtn = $s.Owner.Tag
        if ($localBtn) {
            $localBtn.Tag["ShowText"] = if ($s.Checked) { "true" } else { "false" }
            $form.SuspendLayout()
            Save-Shortcuts
            Update-Layout
            $form.ResumeLayout()
        }
    })
    
    $miAlignRight = New-Object System.Windows.Forms.ToolStripMenuItem("Align right")
    $miAlignRight.CheckOnClick = $true
    if ($DefAlignRight -eq "true") { $miAlignRight.Checked = $true }
    $miAlignRight.Add_Click({
        param($s, $e)
        $localBtn = $s.Owner.Tag
        if ($localBtn) {
            $localBtn.Tag["AlignRight"] = if ($s.Checked) { "true" } else { "false" }
            $form.SuspendLayout()
            Save-Shortcuts
            Update-Layout
            $form.ResumeLayout()
        }
    })
    
    $miAdminFailAlternative.Enabled = $miOpenAdmin.Checked
    $ctxMenu.Items.AddRange(@($miRemove, $miRename, $miShowText, $miAlignRight, $miOpenAdmin, $miAdminFailAlternative))
    $btn.ContextMenuStrip = $ctxMenu

    $btn.Add_MouseEnter({
        param($s, $e)
        if ($global:CurrentTooltipControl -and ($global:CurrentTooltipControl -ne $s)) { $tooltip.Hide($global:CurrentTooltipControl) }
        $global:CurrentTooltipControl = $s
    })
    $btn.Add_MouseDown({ param($s, $e) $s.Tag["DragStart"] = $e.Location })
    $btn.Add_MouseMove({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left -and $s.Tag["DragStart"]) {
            $start = $s.Tag["DragStart"]
            $deltaX = [Math]::Abs($e.Location.X - $start.X)
            $deltaY = [Math]::Abs($e.Location.Y - $start.Y)
            if ($deltaX -gt 5 -or $deltaY -gt 5) {
                $s.DoDragDrop($s, [System.Windows.Forms.DragDropEffects]::Move)
            }
        }
    })
    $btn.Add_MouseLeave({ param($s, $e) ; $tooltip.Hide($s) })
    
    $shortcutsPanel.Controls.Add($btn)
    log "Added ShortcutButton end '$FilePath' - OK"
}

function Export-INI {
    param([string]$INIFile)
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "INI Files|*.ini|All Files|*.*"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { log "Exporting $INIFile in '$($sfd.FileName)'"; Copy-Item -Path $INIFile -Destination $sfd.FileName -Force; log "Exported $INIFile - OK" }
}

$settingsButton.Add_Click({
    $ctx = New-Object System.Windows.Forms.ContextMenuStrip
    $itmSettings = New-Object System.Windows.Forms.ToolStripMenuItem("Settings")
    $itmSettings.Add_Click({ Show-OptionsWindow })
    $itmRefresh = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh Toolbar")
    $itmRefresh.Add_Click({ $form.SuspendLayout(); Update-AppBarPosition -position $global:Settings["ToolbarLocation"]; Update-Layout; $form.ResumeLayout() })
    $itmClose = New-Object System.Windows.Forms.ToolStripMenuItem("Close Toolbar")
    $itmClose.Add_Click({ $form.Close() })
    $ctx.Items.AddRange(@($itmSettings, $itmRefresh, $itmClose))
    $pt = New-Object System.Drawing.Point(0, $settingsButton.Height)
    $ctx.Show($settingsButton, $pt)
})

# --- Options window ---
function Show-OptionsWindow {
    $optionsForm = New-Object System.Windows.Forms.Form
    $optionsForm.SuspendLayout()
    $optionsForm.Text = "Options"
    $optionsForm.ClientSize = New-Object System.Drawing.Size(320,340)
    $optionsForm.StartPosition = 'CenterScreen'

    # Location GroupBox
    $gbLoc = New-Object System.Windows.Forms.GroupBox; $gbLoc.Text = "Location"; $gbLoc.Location = New-Object System.Drawing.Point(10,10); $gbLoc.Size = New-Object System.Drawing.Size(140,70)
    $radioTop = New-Object System.Windows.Forms.RadioButton; $radioTop.Text = "Top"; $radioTop.Location = New-Object System.Drawing.Point(10,20)
    $radioBottom = New-Object System.Windows.Forms.RadioButton; $radioBottom.Text = "Bottom"; $radioBottom.Location = New-Object System.Drawing.Point(10,40)
    if ($global:Settings["ToolbarLocation"] -eq "Top") { $radioTop.Checked = $true } else { $radioBottom.Checked = $true }
    $gbLoc.Controls.AddRange(@($radioTop, $radioBottom))
    $optionsForm.Controls.Add($gbLoc)

    # Size GroupBox
    $gbThick = New-Object System.Windows.Forms.GroupBox; $gbThick.Text = "Size"; $gbThick.Location = New-Object System.Drawing.Point(160,10); $gbThick.Size = New-Object System.Drawing.Size(140,90)
    $radioSmall = New-Object System.Windows.Forms.RadioButton; $radioSmall.Text = "Small"; $radioSmall.Location = New-Object System.Drawing.Point(10,20)
    $radioMedium = New-Object System.Windows.Forms.RadioButton; $radioMedium.Text = "Medium"; $radioMedium.Location = New-Object System.Drawing.Point(10,40)
    $radioLarge = New-Object System.Windows.Forms.RadioButton; $radioLarge.Text = "Large"; $radioLarge.Location = New-Object System.Drawing.Point(10,60)
    switch ($global:Settings["ThicknessMode"]) {
        "Small" { $radioSmall.Checked = $true }
        "Medium" { $radioMedium.Checked = $true }
        "Large" { $radioLarge.Checked = $true }
    }
    $gbThick.Controls.AddRange(@($radioSmall, $radioMedium, $radioLarge))
    $optionsForm.Controls.Add($gbThick)

    # Theme GroupBox
    $gbTheme = New-Object System.Windows.Forms.GroupBox; $gbTheme.Text = "Theme"; $gbTheme.Location = New-Object System.Drawing.Point(10,90); $gbTheme.Size = New-Object System.Drawing.Size(140,70)
    $rLight = New-Object System.Windows.Forms.RadioButton; $rLight.Text = "Light"; $rLight.Location = New-Object System.Drawing.Point(10,20)
    $rDark = New-Object System.Windows.Forms.RadioButton; $rDark.Text = "Dark"; $rDark.Location = New-Object System.Drawing.Point(10,40)
    if ($global:Settings["Theme"] -eq "Dark") { $rDark.Checked = $true } else { $rLight.Checked = $true }
    $gbTheme.Controls.AddRange(@($rLight, $rDark))
    $optionsForm.Controls.Add($gbTheme)

    # Shortcuts GroupBox
    $gbShortcuts = New-Object System.Windows.Forms.GroupBox; $gbShortcuts.Text = "New Shortcuts"; $gbShortcuts.Location = New-Object System.Drawing.Point(160,110); $gbShortcuts.Size = New-Object System.Drawing.Size(140,220)
    $cbShowText = New-Object System.Windows.Forms.CheckBox; $cbShowText.Text = "Show text"; $cbShowText.Location = New-Object System.Drawing.Point(10,20); $cbShowText.Checked = ($global:Settings["NewShortcutShowText"] -eq "true")
    $cbOpenAdmin = New-Object System.Windows.Forms.CheckBox; $cbOpenAdmin.Text = "As admin"; $cbOpenAdmin.Location = New-Object System.Drawing.Point(10,45); $cbOpenAdmin.Checked = ($global:Settings["NewShortcutOpenAsAdmin"] -eq "true")
    $cbAlignRight = New-Object System.Windows.Forms.CheckBox; $cbAlignRight.Text = "Align right"; $cbAlignRight.Location = New-Object System.Drawing.Point(10,70); $cbAlignRight.Checked = ($global:Settings["NewShortcutAlignRight"] -eq "true")
    $gbShortcuts.Controls.AddRange(@($cbShowText, $cbOpenAdmin, $cbAlignRight))
    
    # Nested group box for admin access fail alternative
    $gbAdminFail = New-Object System.Windows.Forms.GroupBox; $gbAdminFail.Text = "If admin access fail"; $gbAdminFail.Location = New-Object System.Drawing.Point(10,100); $gbAdminFail.Size = New-Object System.Drawing.Size(120,110)
    $rAsk = New-Object System.Windows.Forms.RadioButton; $rAsk.Text = "Ask"; $rAsk.Location = New-Object System.Drawing.Point(10,20)
    $rLocalCopy = New-Object System.Windows.Forms.RadioButton; $rLocalCopy.Text = "Local Copy"; $rLocalCopy.Location = New-Object System.Drawing.Point(10,40)
    $rRunAsUser = New-Object System.Windows.Forms.RadioButton; $rRunAsUser.Text = "Run as User"; $rRunAsUser.Location = New-Object System.Drawing.Point(10,60)
    $rCancel = New-Object System.Windows.Forms.RadioButton; $rCancel.Text = "Cancel"; $rCancel.Location = New-Object System.Drawing.Point(10,80)
    switch ($global:Settings["NewShortcutAdminAccessFailAlternative"]) {
        "Ask"       { $rAsk.Checked = $true }
        "LocalCopy" { $rLocalCopy.Checked = $true }
        "AsUser"    { $rRunAsUser.Checked = $true }
        "Cancel"    { $rCancel.Checked = $true }
    }
    $gbAdminFail.Controls.AddRange(@($rAsk, $rLocalCopy, $rRunAsUser, $rCancel))
    $gbShortcuts.Controls.Add($gbAdminFail)
    $optionsForm.Controls.Add($gbShortcuts)

    # Import/Export buttons
    $btnImport = New-Object System.Windows.Forms.Button; $btnImport.Text = "Import Shortcuts/Settings"; $btnImport.Location = New-Object System.Drawing.Point(10,220); $btnImport.Size = New-Object System.Drawing.Size(140,30)
    $btnImport.Add_Click({ Log "Import Begin"; $form.SuspendLayout(); Import-INI; Update-AppBarPosition -position $global:Settings["ToolbarLocation"]; Update-Layout; $form.ResumeLayout(); Log "Import End" })
    $optionsForm.Controls.Add($btnImport)
    $btnExportShortcuts = New-Object System.Windows.Forms.Button; $btnExportShortcuts.Text = "Export Shortcuts"; $btnExportShortcuts.Location = New-Object System.Drawing.Point(10,260); $btnExportShortcuts.Size = New-Object System.Drawing.Size(140,30)
    $btnExportShortcuts.Add_Click({ Export-INI $shortcutsFile })
    $optionsForm.Controls.Add($btnExportShortcuts)
    $btnExportSettings = New-Object System.Windows.Forms.Button; $btnExportSettings.Text = "Export Settings"; $btnExportSettings.Location = New-Object System.Drawing.Point(10,300); $btnExportSettings.Size = New-Object System.Drawing.Size(140,30)
    $btnExportSettings.Add_Click({ Export-INI $settingsFile })
    $optionsForm.Controls.Add($btnExportSettings)

    # Event handlers for Location
    $radioTop.Add_CheckedChanged({ if ($radioTop.Checked) { $global:Settings["ToolbarLocation"] = "Top"; $form.SuspendLayout(); Update-AppBarPosition -position "Top"; $form.ResumeLayout(); Save-Settings } })
    $radioBottom.Add_CheckedChanged({ if ($radioBottom.Checked) { $global:Settings["ToolbarLocation"] = "Bottom"; $form.SuspendLayout(); Update-AppBarPosition -position "Bottom"; $form.ResumeLayout(); Save-Settings } })

    # Event handlers for Size
    $radioSmall.Add_CheckedChanged({ if ($radioSmall.Checked) { $global:Settings["ThicknessMode"] = "Small"; $global:BarThickness = 25; $form.SuspendLayout(); Update-AppBarPosition -position $global:Settings["ToolbarLocation"]; Update-Layout; $form.ResumeLayout() } })
    $radioMedium.Add_CheckedChanged({ if ($radioMedium.Checked) { $global:Settings["ThicknessMode"] = "Medium"; $global:BarThickness = 32; $form.SuspendLayout(); Update-AppBarPosition -position $global:Settings["ToolbarLocation"]; Update-Layout; $form.ResumeLayout() } })
    $radioLarge.Add_CheckedChanged({ if ($radioLarge.Checked) { $global:Settings["ThicknessMode"] = "Large"; $global:BarThickness = 40; $form.SuspendLayout(); Update-AppBarPosition -position $global:Settings["ToolbarLocation"]; Update-Layout; $form.ResumeLayout() } })

    # Event handlers for Theme
    $rLight.Add_CheckedChanged({ if ($rLight.Checked) { $global:Settings["Theme"] = "Light"; $form.SuspendLayout(); Update-Layout; $form.ResumeLayout() } })
    $rDark.Add_CheckedChanged({ if ($rDark.Checked) { $global:Settings["Theme"] = "Dark"; $form.SuspendLayout(); Update-Layout; $form.ResumeLayout() } })

    # Event handlers for New Shortcuts options
    $cbShowText.Add_CheckedChanged({ $global:Settings["NewShortcutShowText"] = if ($cbShowText.Checked) { "true" } else { "false" }; Save-Settings })
    $cbOpenAdmin.Add_CheckedChanged({ $global:Settings["NewShortcutOpenAsAdmin"] = if ($cbOpenAdmin.Checked) { "true" } else { "false" }; Save-Settings })
    $cbAlignRight.Add_CheckedChanged({ $global:Settings["NewShortcutAlignRight"] = if ($cbAlignRight.Checked) { "true" } else { "false" }; Save-Settings })

    # Event handlers for Admin Access Fail options
    $rAsk.Add_CheckedChanged({ if ($rAsk.Checked) { $global:Settings["NewShortcutAdminAccessFailAlternative"] = "Ask"; Save-Settings } })
    $rLocalCopy.Add_CheckedChanged({ if ($rLocalCopy.Checked) { $global:Settings["NewShortcutAdminAccessFailAlternative"] = "LocalCopy"; Save-Settings } })
    $rRunAsUser.Add_CheckedChanged({ if ($rRunAsUser.Checked) { $global:Settings["NewShortcutAdminAccessFailAlternative"] = "AsUser"; Save-Settings } })
    $rCancel.Add_CheckedChanged({ if ($rCancel.Checked) { $global:Settings["NewShortcutAdminAccessFailAlternative"] = "Cancel"; Save-Settings } })

    $optionsForm.ResumeLayout()
    $optionsForm.ShowDialog() | Out-Null
    $optionsForm.Dispose()
}

# --- EHNANCED TEST-PATH ACCESS ---
function Test-PathAcess {
    param([string]$Path)
    try {
        if (Test-Path -LiteralPath $Path) {
            Get-ChildItem -LiteralPath $Path -ErrorAction Stop | Out-Null
            return $true
        }
    } catch { }
    return $false
}

# --- ADMIN HELPER ---
function Start-AdminHelper {
    if ($script:AdminHelperStarted) { return }
    Log "Starting AdminHelper begin..."
    $script:StandardUserSID = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value

    $helperScript = @"
function Test-PathAcess {
    param([string]`$Path)
    try {
        if (Test-Path -LiteralPath `$Path) {
            Get-ChildItem -LiteralPath `$Path -ErrorAction Stop | Out-Null
            return `$true
        }
    } catch { }
    return `$false
}

`$pipeName = 'PSAdminHelperPipe'
while (`$true) {
    try {
        # Configure pipe security
        `$ps = New-Object System.IO.Pipes.PipeSecurity
        `$elevatedSID = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        `$ps.AddAccessRule((New-Object System.IO.Pipes.PipeAccessRule(`$elevatedSID, [System.IO.Pipes.PipeAccessRights]::ReadWrite, [System.Security.AccessControl.AccessControlType]::Allow)))
        `$stdSID = New-Object System.Security.Principal.SecurityIdentifier('$script:StandardUserSID')
        `$ps.AddAccessRule((New-Object System.IO.Pipes.PipeAccessRule(`$stdSID, [System.IO.Pipes.PipeAccessRights]::ReadWrite, [System.Security.AccessControl.AccessControlType]::Allow)))
        `$pipeStream = New-Object System.IO.Pipes.NamedPipeServerStream(`$pipeName, [System.IO.Pipes.PipeDirection]::InOut, 1, [System.IO.Pipes.PipeTransmissionMode]::Byte, [System.IO.Pipes.PipeOptions]::None, 1024, 1024, `$ps)
        `$pipeStream.WaitForConnection()
        # Read the command sent by the client
        `$reader = New-Object System.IO.StreamReader(`$pipeStream)
        `$cmd = `$reader.ReadLine()
        if (`$cmd) {
            try { `$output = (Invoke-Expression `$cmd) }
            catch { `$output = '[Pipeline Error] ' + `$(`$_.Exception.Message) }
            `$writer = New-Object System.IO.StreamWriter(`$pipeStream)
            `$writer.AutoFlush = `$true
            `$writer.WriteLine(`$output)
            `$writer.Dispose()
        }
    }
    catch {
        Write-Host '[Pipeline UNKNOWN Error]: ' + `$(`$_.Exception.Message)
        Read-Host "Press Enter to continue..."
    }
    finally {
        if (`$reader) { `$reader.Dispose() }
        if (`$pipeStream) {
            if (`$pipeStream.IsConnected) { `$pipeStream.Disconnect() }
            `$pipeStream.Dispose()
        }
        Start-Sleep -Seconds 1
    }
}
"@

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "powershell.exe"
        $psi.Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -NoExit -Command $helperScript"
        $psi.Verb = "runas"
        $psi.UseShellExecute = $true
        $script:AdminHelperProcess = [System.Diagnostics.Process]::Start($psi)
        $script:AdminHelperStarted = $true
        Start-Sleep -Seconds 2
        Log "Started AdminHelper - OK"
    }
    catch { Log "Could not start Admin Helper process" }
}

# --- ADMIN RUN ---
function Invoke-AdminRun {
    param(
        [string]$filePath,
        [ValidateSet("LocalCopy", "AsUser", "Cancel")][string]$adminAccessFailAlternative
    )
    Log "Invoke-AdminRun filePath input is '$filePath'"
    
    # Local function to send a command to the helper via the pipe
    function Send-AdminCommand {
        param([string]$command)
        try {
            $client = New-Object System.IO.Pipes.NamedPipeClientStream(".", "PSAdminHelperPipe", [System.IO.Pipes.PipeDirection]::InOut, [System.IO.Pipes.PipeOptions]::None)
            $client.Connect(15000)
            $writer = New-Object System.IO.StreamWriter($client)
            $writer.AutoFlush = $true
            $reader = New-Object System.IO.StreamReader($client)
            $writer.WriteLine($command)
            $writer.Flush()
            return $reader.ReadLine()
        }
        catch { return $null }
        finally {
            if ($writer) { $writer.Dispose() }
            if ($reader) { $reader.Dispose() }
            if ($client) { $client.Dispose() }
        }
    }
    
    # Resolve shortcuts
    if ($filePath -match "\.lnk$") {
        $wsh = New-Object -ComObject WScript.Shell
        $shortcut = $wsh.CreateShortcut($filePath)
        $target = $shortcut.TargetPath
        $arguments = $shortcut.Arguments
        $workDir = $shortcut.WorkingDirectory
    }
    else {
        $target = $filePath
        $arguments = ""
        $workDir = ""
    }
    
    # Test admin access to the target
    $testAdmin_target = Send-AdminCommand "if (Test-PathAcess '$target') { Write-Output OK }"
    if ($testAdmin_target -eq "OK") {
        Log "Admin can read target"
        $cmd = "Start-Process -FilePath '$target'"
        if ($arguments) { $cmd += " -ArgumentList '$arguments'" }
        if ($workDir) {
            $testAdmin_workingDir = Send-AdminCommand "if (Test-PathAcess '$workDir') { Write-Output OK }"
            if ($testAdmin_workingDir -eq "OK") { $cmd += " -WorkingDirectory '$workDir'" }
            else {
                Log "Admin cannot access workingDirectory '$workDir', fallback to target folder."
                $cmd += " -WorkingDirectory '$(Split-Path $target)'"
            }
        }
        Send-AdminCommand $cmd
    }
    else {
        Log "Admin cannot read target"
        if (-not (Test-PathAcess $target)) { Log "User cannot read target either, cancelling." ; return }
        $item = Get-Item $target -ErrorAction SilentlyContinue
        if (-not $item) { Log "Cannot get item" ; return }
        $size = $item.Length; $sizeMB = [math]::Round($size / 1MB, 2)
        $msg = "File not readable as Admin`nLocal copy before running as admin?`nSize: $sizeMB MB"
        
        if (-not $adminAccessFailAlternative) {
            $chooseForm = New-Object System.Windows.Forms.Form
            $chooseForm.Text = "Admin Run"
            $chooseForm.Size = New-Object System.Drawing.Size(500,250)
            $chooseForm.StartPosition = "CenterScreen"
            $label = New-Object System.Windows.Forms.Label
            $label.Text = $msg
            $label.Location = New-Object System.Drawing.Point(10,10)
            $label.Size = New-Object System.Drawing.Size(480,100)
            $label.AutoSize = $false
            $buttonAdmin = New-Object System.Windows.Forms.Button
            $buttonAdmin.Text = "Local copy, then run as admin"
            $buttonAdmin.Size = New-Object System.Drawing.Size(140,30)
            $buttonAdmin.Location = New-Object System.Drawing.Point(10,150)
            $buttonUser = New-Object System.Windows.Forms.Button
            $buttonUser.Text = "Run as user"
            $buttonUser.Size = New-Object System.Drawing.Size(140,30)
            $buttonUser.Location = New-Object System.Drawing.Point(170,150)
            $buttonCancel = New-Object System.Windows.Forms.Button
            $buttonCancel.Text = "Cancel"
            $buttonCancel.Size = New-Object System.Drawing.Size(140,30)
            $buttonCancel.Location = New-Object System.Drawing.Point(330,150)
            $chooseForm.Controls.Add($label)
            $chooseForm.Controls.Add($buttonAdmin)
            $chooseForm.Controls.Add($buttonUser)
            $chooseForm.Controls.Add($buttonCancel)
            $buttonAdmin.DialogResult = [System.Windows.Forms.DialogResult]::Yes
            $buttonUser.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
            $buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $chooseForm.AcceptButton = $buttonAdmin
            $chooseForm.CancelButton = $buttonCancel
            $result = $chooseForm.ShowDialog()
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { $adminAccessFailAlternative = "LocalCopy" }
            elseif ($result -eq [System.Windows.Forms.DialogResult]::Ignore) { $adminAccessFailAlternative = "AsUser" }
            else { $adminAccessFailAlternative = "Cancel" }
        }
        
        switch ($adminAccessFailAlternative) {
			"LocalCopy" {
				$temp = Join-Path $env:TEMP $($item.BaseName + $item.Extension)
				Log "Copying file to temp: $temp"
				try {
					if (Test-Path $temp) { Remove-Item $temp -Recurse -Force -ErrorAction Stop }
					Copy-Item -Path $target -Destination $temp -ErrorAction Stop
				}
				catch {
					[System.Windows.Forms.MessageBox]::Show("Error while suppress existing or copy in temp : " + $_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
					return
				}
				else {
					$cmd = "Start-Process -FilePath '$temp'"
					if ($arguments) { $cmd += " -ArgumentList '$arguments'" }
					if ($workDir) {
						$testAdmin_workingDir = Send-AdminCommand "if (Test-PathAcess '$workDir') { Write-Output OK }"
						if ($testAdmin_workingDir -eq "OK") { $cmd += " -WorkingDirectory '$workDir'" }
						else {
							Log "Admin cannot access workingDirectory '$workDir', fallback to target folder."
							$cmd += " -WorkingDirectory '$(Split-Path $temp)'"
						}
					}
				}
				Send-AdminCommand $cmd
			}
            "AsUser" {
                # Option: Run as user
                $cmd = "Start-Process -FilePath '$target'"
                if ($arguments) { $cmd += " -ArgumentList '$arguments'" }
                if ($workDir) {
                    if (Test-PathAcess $workDir) { $cmd += " -WorkingDirectory '$workDir'" }
                    else {
                        Log "User cannot access workingDirectory '$workDir', fallback to target folder."
                        $cmd += " -WorkingDirectory '$(Split-Path $target)'"
                    }
                }
                Invoke-Expression $cmd
            }
            "Cancel" {
                Log "adminAccessFailAlternative = Cancel"
                return
            }
        }
    }
}

# --- Events for drag/drop and layout reordering ---
$shortcutsPanel.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.Button])) { $e.Effect = [System.Windows.Forms.DragDropEffects]::Move ; $dragIndicator.Visible = $true }
    elseif ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy }
    else { $e.Effect = [System.Windows.Forms.DragDropEffects]::None }
})

$shortcutsPanel.Add_DragOver({
    param($s, $e)
    function Get-IndicatorX($mouseX, $isRightGroup) {
        $alignedButtons = @($shortcutsPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.ContainsKey("FilePath") -and (($_.Tag["AlignRight"] -eq "true") -eq $isRightGroup) })
        if ($alignedButtons.Count -eq 0) {
            if ($isRightGroup)  { return ($shortcutsPanel.ClientSize.Width - $dragIndicator.Width) }
            else                { return 0 }
        } else {
            $closestButton = $null; $minDistance = [int]::MaxValue
            foreach ($btn in $alignedButtons) { $distance = [Math]::Abs($mouseX - ($btn.Left + $btn.Width / 2)) ; if ($distance -lt $minDistance) { $minDistance = $distance; $closestButton = $btn } }
            if ($mouseX -lt ($closestButton.Left + $closestButton.Width/2)) { return [int]($closestButton.Left - $dragIndicator.Width + 2) } 
            else                                                            { return [int]($closestButton.Left + $closestButton.Width - 1) }
        }
    }
    if ($e.Data.GetDataPresent([System.Windows.Forms.Button])) {
        $dragButton = $e.Data.GetData([System.Windows.Forms.Button])
        if ($dragButton) {
            $isRight = ($e.X -gt ($shortcutsPanel.ClientSize.Width / 2))
            $dragIndicator.Location = New-Object System.Drawing.Point($(Get-IndicatorX $e.X $isRight), 0)
            $dragIndicator.Visible = $true
        }
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
    }
    elseif ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        $isRight = ($e.X -gt ($shortcutsPanel.ClientSize.Width / 2))
        $dragIndicator.Location = New-Object System.Drawing.Point($(Get-IndicatorX $e.X $isRight), 0)
        $dragIndicator.Visible = $true
    }
})

function Update-ButtonPosition($button, $xCoordinate) {
    if (-not $button) { return }
    $shortcutsPanel.Controls.Remove($button)
    $allButtons = $shortcutsPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.ContainsKey("FilePath") }
    $leftButtons = @($allButtons | Where-Object { $_.Tag["AlignRight"] -ne "true" } | Sort-Object Left)
    $rightButtons = @($allButtons | Where-Object { $_.Tag["AlignRight"] -eq "true" } | Sort-Object Left)
    $isRightGroup = $xCoordinate -gt ($shortcutsPanel.ClientSize.Width/2)
    $targetButtons = if ($isRightGroup) { $rightButtons } else { $leftButtons }
    $insertIndex = 0
    foreach ($b in $targetButtons) { if ($xCoordinate -lt ($b.Left + $b.Width/2)) { break }; $insertIndex++ }
    $newGroup = @()
    if ($insertIndex -gt 0) { $newGroup += $targetButtons | Select-Object -First $insertIndex }
    $newGroup += $button
    if ($insertIndex -lt $targetButtons.Count) { $newGroup += $targetButtons | Select-Object -Skip $insertIndex }
    $newList = if ($isRightGroup) { $leftButtons + $newGroup } else { $newGroup + $rightButtons }
    $button.Tag["AlignRight"] = $isRightGroup.ToString().ToLower()
    ($button.ContextMenuStrip.Items | Where-Object { $_.Text -eq "Align right" }).Checked = $isRightGroup
    for ($j = 0; $j -lt $newList.Count; $j++) { $newList[$j].Tag.Order = $j }
    $allButtons | ForEach-Object { $shortcutsPanel.Controls.Remove($_) }
    $newList | ForEach-Object { $shortcutsPanel.Controls.Add($_) }
    Save-Shortcuts; Update-Layout
}

function Update-FileDrop($files) {
    $allButtons = $shortcutsPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.FilePath }
    $lookup = @{}
    $allButtons | ForEach-Object { if ($_.Tag.FilePath -and -not $lookup.ContainsKey($_.Tag.FilePath)) { $lookup[$_.Tag.FilePath] = $_ } }
    foreach ($file in $files) {
        $existingButton = $lookup[$file]
        if ($existingButton) {
            $originalColor = $existingButton.BackColor
            for ($i = 1; $i -le 6; $i++) {
                $existingButton.BackColor = if ($existingButton.BackColor -eq 'Orange') { $originalColor } else { 'Orange' }
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 300
            }
            $existingButton.BackColor = $originalColor
        } else {
            $form.SuspendLayout()
            Add-ShortcutButton -FilePath $file -Position $e.X
            Start-Sleep -Milliseconds 50
            $newButton = $shortcutsPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.FilePath -eq $file } | Select-Object -First 1
            if ($newButton) { Update-ButtonPosition $newButton $e.X }
            $form.ResumeLayout()
        }
    }
    $dragIndicator.Visible = $false
}

$shortcutsPanel.Add_DragDrop({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.Button])) {
        $dragButton = $e.Data.GetData([System.Windows.Forms.Button])
        Update-ButtonPosition $dragButton $e.X
        $dragIndicator.Visible = $false
    }
    elseif ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        log "Files dropped on LaunchBar - $files"
        Update-FileDrop $files
    }
})

[Microsoft.Win32.SystemEvents]::add_DisplaySettingsChanged({
    $global:LastDisplaySettingsTime = Get-Date
    log "DisplaySettingsChanged event begin..."
    if (-not $global:DebounceActive) {
        $global:DebounceActive = $true
        while (((Get-Date) - $global:LastDisplaySettingsTime).TotalSeconds -lt 2) { Start-Sleep -Milliseconds 100 ; [System.Windows.Forms.Application]::DoEvents() }
        Update-AppBarPosition -position $global:Settings["ToolbarLocation"]
        Update-Layout
        $global:DebounceActive = $false
    }
    log "DisplaySettingsChanged event end..."
})

function Import-INI {
    param([string]$FilePath)
    
    # If no file is provided, prompt the user to select one
    if (-not $FilePath) {
        Log "No file specified, prompting user..."
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "INI Files|*.ini|All Files|*.*"
        if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { Log "User cancelled file selection."; return }
        $FilePath = $ofd.FileName
        Log "File selected: $FilePath"
    }
    
    if (-not (Test-Path -LiteralPath $FilePath)) {
        [System.Windows.Forms.MessageBox]::Show("INI not readable.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Log "File not found: $FilePath"
        return
    }
    
    Log "Reading INI file: $FilePath"
    $ini = Read-IniFile $FilePath
    
    if ($ini.Contains("Settings")) {
        Log "Processing Settings section."
        if (-not $ini["Settings"] -or $ini["Settings"].Keys.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("The [Settings] section is empty.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Log "Settings section is empty."
            return
        }
        foreach($k in @($global:Settings.Keys)){
            if ($ini["Settings"].Keys -contains $k) {
                $global:Settings[$k] = $ini["Settings"][$k]
                #Set-Variable -Name $k -Value $ini["Settings"][$k] -Scope Global
            }
        }
        switch ($global:Settings["ThicknessMode"]) {
            "Small"  { $global:BarThickness = 25 }
            "Medium" { $global:BarThickness = 32 }
            "Large"  { $global:BarThickness = 40 }
        }
        return 
    }
    elseif ($ini.Keys | Where-Object { $_ -match '^Shortcut\d+$' }) {
        Log "Processing Shortcuts file."
        $fileContent = Get-Content -LiteralPath $FilePath -Raw
        if (-not $fileContent.Trim()) {
            [System.Windows.Forms.MessageBox]::Show("Error: The file is empty.", "File Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Log "File content is empty."
            return
        }
        $sectionPattern = '^\[(?<section>Shortcut\d+)\]\s*\r?\n(?<content>.*?)(?=^\[Shortcut\d+\]\s*\r?\n|\z)'
        $regexOptions = [System.Text.RegularExpressions.RegexOptions]::Compiled -bor [System.Text.RegularExpressions.RegexOptions]::Multiline -bor [System.Text.RegularExpressions.RegexOptions]::Singleline
        $regex = [regex]::new($sectionPattern, $regexOptions)
        $sections = $regex.Matches($fileContent)
        if ($sections.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Error: The file does not contain any valid shortcut sections.", "Structure Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Log "No valid shortcut sections found."
            return
        }
        $requiredKeys = @("ShowText", "AlignRight", "Path", "DisplayName", "Order", "OpenAsAdmin", "AdminAccessFailAlternative")
        foreach ($section in $sections) {
            $sectionName = $section.Groups["section"].Value
            $content = $section.Groups["content"].Value.Trim()
            $missingKeys = @()
            foreach ($key in $requiredKeys) {
                if ($content -notmatch "(?im)^\s*$key\s*=") { $missingKeys += $key }
            }
            if ($missingKeys.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("Error: Section [$sectionName] is missing the following key(s): " + ($missingKeys -join ", "), "Structure Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Log "Section [$sectionName] is missing keys: $($missingKeys -join ', ')"
                return
            }
        }
        Log "Shortcut file validation passed."
    
        # Process shortcut import options
        if ($force.IsPresent) {
            $choice = [System.Windows.Forms.DialogResult]::Yes
            Log "Silent import forced."
        }
        elseif ((Test-Path -LiteralPath $shortcutsFile) -and ((Read-IniFile $shortcutsFile).Keys.Count -gt 0)) {
            $choice = [System.Windows.Forms.MessageBox]::Show("Add imported shortcuts to the existing ones? (Yes = Add, No = Replace)", "Import Options", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
            Log "User chose import option: $choice"
            if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) { Log "User cancelled shortcut import."; return }
        }
        else {
            $choice = [System.Windows.Forms.DialogResult]::No
            Log "No existing shortcuts found; defaulting to replace."
        }
    
        if ($choice -eq [System.Windows.Forms.DialogResult]::No) {
            $shortcutsPanel.Controls.Clear()
            $existing = $ini
            Log "Replacing existing shortcuts."
        }
        else {
            $existing = if (Test-Path -LiteralPath $shortcutsFile) { Read-IniFile $shortcutsFile } else { @{} }
            $index = $existing.Keys.Count
            foreach ($section in $ini.Keys) {
                $existing["Shortcut$index"] = $ini[$section]
                $index++
            }
            Log "Adding imported shortcuts to existing ones."
        }
    
        # Build lookup for existing shortcut buttons
        $lookup = @{}
        $shortcutsPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.FilePath } | ForEach-Object {
            if (-not $lookup.ContainsKey($_.Tag.FilePath)) { $lookup[$_.Tag.FilePath] = $_ }
        }
        Log "Built lookup for existing shortcut buttons."
    
        $buttonsToBlink = @()
        foreach ($section in $ini.Keys) {
            $entry = $ini[$section]
            if ($lookup.ContainsKey($entry.Path)) {
                if (-not $force.IsPresent) {
                    $btn = $lookup[$entry.Path]
                    if (-not ($buttonsToBlink | Where-Object { $_.Button -eq $btn })) {
                        $buttonsToBlink += [PSCustomObject]@{ Button = $btn; OriginalColor = $btn.BackColor }
                    }
                }
            }
            else {
                Add-ShortcutButton -FilePath $entry.Path -NoSave -DefOpenAsAdmin $entry.OpenAsAdmin -DefShowText $entry.ShowText -DefAlignRight $entry.AlignRight -DefDisplayName $entry.DisplayName
                Log "Added shortcut button for $($entry.Path)"
            }
        }
        Log "Layout updated."
    
        if (-not $force.IsPresent -and $buttonsToBlink.Count -gt 0) {
            Log "Blinking buttons to highlight updated shortcuts."
            for ($i = 1; $i -le 6; $i++) {
                foreach ($item in $buttonsToBlink) {
                    $btn = $item.Button
                    $btn.BackColor = if ($btn.BackColor -eq 'Green') { $item.OriginalColor } else { 'Green' }
                }
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 300
            }
            foreach ($item in $buttonsToBlink) { $item.Button.BackColor = $item.OriginalColor }
            Log "Blinking finished."
        }
        Save-Shortcuts
        Log "Shortcuts saved."
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("The INI file does not correspond to either shortcuts or settings.", "Invalid File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Log "INI file format is invalid."
    }
}

Log "Loading existing shortcuts from '$shortcutsFile'..."
if (Test-Path -LiteralPath $shortcutsFile) {
    $ini = Read-IniFile $shortcutsFile
    foreach ($section in $ini.Keys) {
        $entry = $ini[$section]
        Add-ShortcutButton -FilePath $entry.Path -NoSave -DefOpenAsAdmin $entry.OpenAsAdmin -DefShowText $entry.ShowText -DefAlignRight $entry.AlignRight -DefDisplayName $entry.DisplayName
    }
}
Log "Loaded existing shortcuts - OK"

if ($IniFiles.Count -eq 0) { Log "Loading existing settings from '$settingsFile'..."; Import-INI $settingsFile ; Log "Loaded existing settings - OK" }
else { foreach ($IniFile in $IniFiles) { Log "Importing INI file at launch: '$IniFile'..."; Import-INI $IniFile ; Log "Imported INI file at launch: '$IniFile' - OK"} }

Update-AppBarPosition -position $global:Settings["ToolbarLocation"]
$form.Add_Shown({ $form.ResumeLayout() ; Update-Layout : log "Main Form loaded." })

$form.Add_FormClosing({
    log "Main Form closing."
    if ($global:BaselineWorkArea) { [AppBar]::SystemParametersInfo([AppBar]::SPI_SETWORKAREA, 0, [ref]$global:BaselineWorkArea, [AppBar]::SPIF_UPDATEINIFILE) | Out-Null }
    if ($script:AdminHelperProcess -and -not $script:AdminHelperProcess.HasExited) { log "Killing AdminHelperProcess..."; $script:AdminHelperProcess.Kill(); log "Killed AdminHelperProcess - OK" }
})

[System.Windows.Forms.Application]::Run($form)
log "Self removing from $PSCommandPath"
Remove-Item $PSCommandPath -Force
