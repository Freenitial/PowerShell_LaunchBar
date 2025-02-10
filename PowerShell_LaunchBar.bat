<# ::
    cls & @echo off & chcp 437 >nul & title PowerShell LaunchBar

    REM Author  : Leo Gillet - Freenitial on GitHub
    REM Version : 0.9

    REM Optionnal arguments : 
    REM     1) Filepath of .ini file containing shortcuts
    REM     2) /f to force adding those shortcuts if not already exist
    REM --------------------------------------------------------------------
    REM To start from other batch or cmd without exit, launch like this :
    REM start "" /d "FOLDER\CONTAINING_batchfile" PowerShell_LaunchBar
    REM --------------------------------------------------------------------
    REM To start from other batch or cmd without exit + FORCE IMPORT SHORTCUTS FILE, launch like this :
    REM start "" /d "FOLDER\CONTAINING_batchfile" PowerShell_LaunchBar "FULLPATH\TO_IMPORT\SHORTCUT.INI" /f

    copy /y "%~f0" "%TEMP%\%~n0.ps1" >NUL && powershell -Nologo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%TEMP%\%~n0.ps1" "%~1" "%~2"
    exit /b
#>

param([string]$argIniShortcutsFile, [string]$argSilentImport)

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

[DPIHelper]::SetProcessDpiAwarenessContext([DPIHelper]::DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2) | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()

# Global variables and paths
$savePath = Join-Path $env:LOCALAPPDATA "Powershell_Launchbar"
$shortcutsFile = Join-Path $savePath "shortcuts.ini"
$settingsFile  = Join-Path $savePath "settings.ini"
if (-not (Test-Path -LiteralPath $savePath)) { New-Item -ItemType Directory -Path $savePath | Out-Null }

$global:Settings = @{
    ToolbarLocation        = "Top"
    ThicknessMode          = "Small"
    Theme                  = "Light"
    NewShortcutShowText    = "true"
    NewShortcutOpenAsAdmin = "false"
    NewShortcutAlignRight  = "false"
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
if (Test-Path -LiteralPath $settingsFile) {
    $ini = Read-IniFile $settingsFile
    if($ini.Keys -contains "Settings"){
        foreach($k in @($global:Settings.Keys)){
            if ($ini["Settings"].Keys -contains $k) { $global:Settings[$k] = $ini["Settings"][$k] }
        }
    }
}

$global:ToolbarLocation        = $global:Settings["ToolbarLocation"]
$global:ThicknessMode          = $global:Settings["ThicknessMode"]
$global:Theme                  = $global:Settings["Theme"]
$global:NewShortcutShowText    = $global:Settings["NewShortcutShowText"]
$global:NewShortcutOpenAsAdmin = $global:Settings["NewShortcutOpenAsAdmin"]
$global:NewShortcutAlignRight  = $global:Settings["NewShortcutAlignRight"]
$global:DebounceActive = $false
$global:BaselineWorkArea = $null
$global:PrevToolbarLocation = $null
$global:LastDisplaySettingsTime = Get-Date
$canRequireAdminExtensions = @(".exe",".bat",".cmd",".ps1",".msc",".msi",".msp",".vbs",".vbe",".js",".jse",".wsf",".wsh",".cpl",".reg")

switch ($global:ThicknessMode) {
    "Small"  { $global:BarThickness = 25 }
    "Medium" { $global:BarThickness = 32 }
    "Large"  { $global:BarThickness = 40 }
    default  { $global:BarThickness = 25 }
}

function Write-IniFile($Path,$Data){
    $lines = @()
    foreach($section in $Data.Keys){
        if ($Data[$section] -is [hashtable]) {
            $lines += "[$section]"
            foreach ($k in $Data[$section].Keys) { $lines += "$k=$($Data[$section][$k])" }
        } else { $lines += "$section=$($Data[$section])" }
    }
    Set-Content -Path $Path -Value $lines
}

function Save-Shortcuts {
    $data = [ordered]@{}; $index = 0
    foreach ($ctrl in $shortcutsPanel.Controls) {
        if ($ctrl -is [System.Windows.Forms.Button] -and $ctrl.Tag -and $ctrl.Tag.ContainsKey("FilePath")) {
            $data["Shortcut$index"] = @{
                Path        = $ctrl.Tag.FilePath
                OpenAsAdmin = $ctrl.Tag.OpenAsAdmin
                ShowText    = $ctrl.Tag.ShowText
                AlignRight  = $ctrl.Tag.AlignRight
                DisplayName = $ctrl.Tag.DisplayName
                Order       = $ctrl.Tag.Order
            }
            $index++
        }
    }
    Write-IniFile $shortcutsFile $data
}

function Save-Settings {
    $global:Settings["ToolbarLocation"]        = $global:ToolbarLocation
    $global:Settings["ThicknessMode"]          = $global:ThicknessMode
    $global:Settings["Theme"]                  = $global:Theme
    $global:Settings["NewShortcutShowText"]    = $global:NewShortcutShowText
    $global:Settings["NewShortcutOpenAsAdmin"] = $global:NewShortcutOpenAsAdmin
    $global:Settings["NewShortcutAlignRight"]  = $global:NewShortcutAlignRight
    Write-IniFile $settingsFile @{ Settings = $global:Settings }
}

# Create main window (Windows Forms)
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
    param($sender, $e)
    $e.Graphics.FillRectangle([System.Drawing.Brushes]::Black, $e.Bounds)
    $e.Graphics.DrawRectangle([System.Drawing.Pens]::White, [System.Drawing.Rectangle]::new($e.Bounds.X, $e.Bounds.Y, $e.Bounds.Width - 1, $e.Bounds.Height - 1))
    [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, $e.ToolTipText, $tooltipFont, $e.Bounds, [System.Drawing.Color]::White)
})
$tooltip.add_Popup({param($sender, $e); $e.ToolTipSize = [System.Windows.Forms.TextRenderer]::MeasureText($tooltip.GetToolTip($e.AssociatedControl), $tooltipFont) })

$dragIndicator = New-Object System.Windows.Forms.Panel
$dragIndicator.BackColor = [System.Drawing.Color]::Orange
$dragIndicator.Width = 5
$dragIndicator.Height = $global:BarThickness
$dragIndicator.Visible = $false
$shortcutsPanel.Controls.Add($dragIndicator)
$dragIndicator.BringToFront()

# --- Update Appbar position and working area ---
function Update-AppBarPosition {
    param([string]$position)
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
    $form.Left = $global:BaselineWorkArea.left ; $form.Width = $global:BaselineWorkArea.right - $global:BaselineWorkArea.left ; $form.Height = $global:BarThickness
    switch ($position) {
        "Top" { $edge=[AppBar]::ABE_TOP ; $form.Top=$global:BaselineWorkArea.top ; $workTop=$global:BaselineWorkArea.top+$global:BarThickness ; $workBottom=$global:BaselineWorkArea.bottom}
        "Bottom" { $edge=[AppBar]::ABE_BOTTOM;$form.Top=$global:BaselineWorkArea.bottom-$global:BarThickness;$workTop=$global:BaselineWorkArea.top;$workBottom=$global:BaselineWorkArea.bottom-$global:BarThickness}
    }
    $appBarData.uEdge = $edge
    $appBarData.rc = New-Object AppBar+RECT -Property @{ left=$form.Left; top=$form.Top; right=$form.Left+$form.Width; bottom=$form.Top+$form.Height }
    [AppBar]::SHAppBarMessage([AppBar]::ABM_NEW, [ref]$appBarData) | Out-Null
    [AppBar]::SHAppBarMessage([AppBar]::ABM_SETPOS, [ref]$appBarData) | Out-Null
    $newWorkArea = New-Object AppBar+RECT -Property @{ left=$global:BaselineWorkArea.left; top=$workTop; right=$global:BaselineWorkArea.right; bottom=$workBottom }
    [AppBar]::SystemParametersInfo([AppBar]::SPI_SETWORKAREA, 0, [ref]$newWorkArea, [AppBar]::SPIF_UPDATEINIFILE) | Out-Null
    Start-Sleep -Milliseconds 200
    $global:PrevToolbarLocation = $position
}

# --- Update layout and appearance ---
function Update-Layout {
    switch ($global:ThicknessMode) { "Small" { $iconSize=14; $fontSize=8 } ; "Medium" { $iconSize=22; $fontSize=10 } ; "Large" { $iconSize=28; $fontSize=12 } }
    $themeColors = if ($global:Theme -eq "Dark") { @{ BackColor = [System.Drawing.Color]::FromArgb(51,51,51); ForeColor = [System.Drawing.Color]::White; BorderColor = [System.Drawing.Color]::DimGray } } 
                   else { @{ BackColor = [System.Drawing.Color]::FromArgb(238,238,238); ForeColor = [System.Drawing.Color]::Black; BorderColor = [System.Drawing.Color]::Gray } }
    $shortcutsPanel.BackColor=$themeColors.BackColor
    $settingsButton.ForeColor=$themeColors.ForeColor; $settingsButton.FlatAppearance.BorderColor=$themeColors.BorderColor; $settingsButton.BackColor=$shortcutsPanel.BackColor
    $dragIndicator.Height = $global:BarThickness
    $leftButtons = New-Object 'System.Collections.Generic.List[System.Windows.Forms.Button]' ; $rightButtons = New-Object 'System.Collections.Generic.List[System.Windows.Forms.Button]'
    $g = [System.Drawing.Graphics]::FromImage((New-Object System.Drawing.Bitmap(1,1)))
    foreach ($btn in $shortcutsPanel.Controls) {
        if ($btn -is [System.Windows.Forms.Button] -and $btn.Tag -and $btn.Tag.ContainsKey("FilePath")) {
            switch ($global:Theme) {
                "Dark"  { $btn.FlatAppearance.BorderColor=[System.Drawing.Color]::DimGray; $btn.ForeColor=[System.Drawing.Color]::White; $btn.BackColor=[System.Drawing.Color]::DarkSlateGray }
                default { $btn.FlatAppearance.BorderColor=[System.Drawing.Color]::Gray; $btn.ForeColor=[System.Drawing.Color]::Black; $btn.BackColor=[System.Drawing.Color]::White }
            }
            if ($btn.Tag["DisplayName"] -ne "") {$displayName=$btn.Tag["DisplayName"]} else {$displayName=$btn.Tag["FilePath"]}
            if ($btn.Tag["ShowText"] -eq "true") {$btn.Text = " $displayName" ; $tooltip.SetToolTip($btn, $null) }
            else { $btn.Text = "" ; $tooltip.SetToolTip($btn, $displayName) }
            $btn.Height = $global:BarThickness
            $btn.AutoEllipsis = $true
            $btn.TextImageRelation = "ImageBeforeText"
            $btn.ImageAlign = "MiddleLeft"
            $btn.TextAlign = "MiddleLeft"
            if($btn.Tag.ContainsKey("OriginalImage")-and$btn.Tag["OriginalImage"]) {try{$btn.Image=$btn.Tag["OriginalImage"].GetThumbnailImage($iconSize,$iconSize,$null,[IntPtr]::Zero)}catch{}}
            $txtWidth = [Math]::Ceiling($g.MeasureString($btn.Text, $btn.Font).Width)
            if ($btn.Tag["ShowText"] -eq "true") {$padding=10} else {$padding=13}
            $btn.Width = $iconSize + $txtWidth + $padding
            if ($btn.Tag["AlignRight"] -eq "true") { $rightButtons.Add($btn) } else { $leftButtons.Add($btn) }
        }
    }
    # Overflow handling: proportionally shrink text buttons if total width exceeds panel width
    $spacing=2
    $panelWidth=$shortcutsPanel.ClientSize.Width
    $buttonList=if ($leftButtons.Count-gt0 -and $rightButtons.Count-gt0) {$leftButtons+$rightButtons} elseif ($leftButtons.Count-gt0) {$leftButtons} elseif ($rightButtons.Count-gt0) {$rightButtons} else {@()}
    if ($buttonList.Count) {$totalSpacing = if ($leftButtons.Count -gt 0 -and $rightButtons.Count -gt 0) {($buttonList.Count - 2) * $spacing} else {($buttonList.Count - 1) * $spacing}
        $fixed = 0; $var = @()
        foreach ($btn in $buttonList) {if ($btn.Tag["ShowText"] -eq "true") {$var += $btn} else {$fixed += $btn.Width}}
        $origVar = ($var | Measure-Object -Property Width -Sum).Sum
        if (($fixed + $origVar + $totalSpacing) -gt $panelWidth -and $origVar) {
            $scale = ($panelWidth - $fixed - $totalSpacing) / $origVar
            foreach ($btn in $var) {$btn.Width = [Math]::Floor($btn.Width * $scale)}
        }
    }
    $cs=$shortcutsPanel.ClientSize ; $leftX=0; $rightX=$cs.Width
    foreach ($btn in $leftButtons) { $btn.Location=New-Object System.Drawing.Point($leftX, 0) ; $leftX += $btn.Width + $spacing }
    $rightSorted = $rightButtons | Sort-Object { [int]$_.Tag.Order } -Descending
    foreach ($btn in $rightSorted) { $rightX-=$btn.Width ; $btn.Location=New-Object System.Drawing.Point($rightX, 0) ; $rightX-=$spacing }
    $g.Dispose()
    Save-Settings
}

# --- Add a shortcut button ---
function Add-ShortcutButton {
    param(
        [string]$FilePath,
        [switch]$NoSave,
        [string]$DefOpenAsAdmin = $global:NewShortcutOpenAsAdmin,
        [string]$DefShowText    = $global:NewShortcutShowText,
        [string]$DefAlignRight  = $global:NewShortcutAlignRight,
        [string]$DefDisplayName = ""
    )

    if (-not (Test-Path -LiteralPath $FilePath)) { Write-Warning "File or folder '$FilePath' does not exist."; return }
    $isFolder = (Test-Path -LiteralPath $FilePath -PathType Container)
    if (-not $isFolder) { $ext = [System.IO.Path]::GetExtension($FilePath).ToLower() } else { $ext = "" }
    $canOpenAsAdmin = $isFolder -or ($canRequireAdminExtensions -contains $ext)
    if (-not $canOpenAsAdmin) { $DefOpenAsAdmin = "false" }
    $icon = $null; try { $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($FilePath) } catch { $icon = $null }
    if (-not $icon) { try { $icon = [IconExtractor]::GetIcon($FilePath) } catch { } }
    $img=$null;if($icon){try{$ms=New-Object System.IO.MemoryStream;$icon.ToBitmap().Save($ms,[System.Drawing.Imaging.ImageFormat]::Png);$ms.Position=0;$bmp=New-Object System.Drawing.Bitmap($ms);$img=$bmp}catch{}}
    $btn = New-Object System.Windows.Forms.Button ; $btn.TabStop = $false ; $btn.FlatStyle = 'Flat' ; $btn.FlatAppearance.BorderSize = 1 ; $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::DarkCyan
    $displayName = if ($DefDisplayName -ne "") { $DefDisplayName } else { Split-Path $FilePath -Leaf }
    $btn.Tag = @{ FilePath = $FilePath; OpenAsAdmin = $DefOpenAsAdmin; ShowText = $DefShowText; AlignRight = $DefAlignRight; DisplayName = $displayName; DragStart = $null }
    if ($DefAlignRight -eq "true") {$group=$shortcutsPanel.Controls | Where-Object {$_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.ContainsKey("FilePath") -and $_.Tag["AlignRight"] -eq "true"}}
    else {$group=$shortcutsPanel.Controls | Where-Object {$_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.ContainsKey("FilePath") -and $_.Tag["AlignRight"] -ne "true"}}
    $btn.Tag.Order = $group.Count
    if ($img) { $btn.Tag["OriginalImage"] = $img }

    $btn.Add_Click({
        param($s,$e) ; $opts = $s.Tag
        if ($opts["FilePath"] -and (Test-Path -LiteralPath $opts["FilePath"])) {
            if ($opts["OpenAsAdmin"] -eq "true") {if (-not $script:AdminHelperStarted) { Start-AdminHelper } ; Invoke-AdminCommand $opts["FilePath"]} 
            else { Start-Process -FilePath $opts["FilePath"] }
        }
    })

    # Context menu (right-click)
    $ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip

    $miRemove = New-Object System.Windows.Forms.ToolStripMenuItem("Remove")
    $miRemove.Add_Click({
        param($s,$e)
        $btn = $s.GetCurrentParent().SourceControl -as [System.Windows.Forms.Button]
        if ($btn) { $form.SuspendLayout(); $shortcutsPanel.Controls.Remove($btn); Save-Shortcuts; Update-Layout; $form.ResumeLayout() }
    })

    $miRename = New-Object System.Windows.Forms.ToolStripMenuItem("Rename")
    $miRename.Add_Click({
        param($s,$e)
        $btn = $s.GetCurrentParent().SourceControl -as [System.Windows.Forms.Button]
        if ($btn) {
            $currentName = $btn.Tag["DisplayName"]
            $frm = New-Object System.Windows.Forms.Form
            $frm.Text = "Rename Shortcut"; $frm.Size = New-Object System.Drawing.Size(300,150); $frm.StartPosition = 'CenterScreen'
            $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "New Name:"; $lbl.Location = New-Object System.Drawing.Point(10,20); $lbl.AutoSize = $true; $frm.Controls.Add($lbl)
            $txt = New-Object System.Windows.Forms.TextBox; $txt.Text = $currentName; $txt.Location = New-Object System.Drawing.Point(80,18); $txt.Width = 180; $frm.Controls.Add($txt)
            $btnOK = New-Object System.Windows.Forms.Button; $btnOK.Text = "OK"; $btnOK.Location = New-Object System.Drawing.Point(100,60)
            $btnOK.Add_Click({ $frm.Tag = $txt.Text; $frm.DialogResult = [System.Windows.Forms.DialogResult]::OK; $frm.Close() })
            $frm.Controls.Add($btnOK); $frm.AcceptButton = $btnOK
            if ($frm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $form.SuspendLayout()
                $newName = $frm.Tag
                if ($newName.Trim() -ne "") {
                    $btn.Tag["DisplayName"] = $newName.Trim()
                    if ($btn.Tag["ShowText"] -eq "true") { $btn.Text = " " + $newName.Trim() }
                    Save-Shortcuts; Update-Layout
                    $form.ResumeLayout()
                }
            }
            $frm.Dispose()
        }
    })

    $miOpenAdmin = New-Object System.Windows.Forms.ToolStripMenuItem("Open as admin")
    $miOpenAdmin.CheckOnClick = $true
    $miOpenAdmin.Checked = ($DefOpenAsAdmin -eq "true")
    if (-not $canOpenAsAdmin) { $miOpenAdmin.Enabled = $false }
    else {
        $miOpenAdmin.Add_Click({
            param($s, $e)
            $btn = $s.GetCurrentParent().SourceControl -as [System.Windows.Forms.Button]
            $btn.Tag["OpenAsAdmin"] = if ($s.Checked) { "true" } else { "false" }
            Save-Shortcuts
        })
    }

    $miShowText = New-Object System.Windows.Forms.ToolStripMenuItem("Show text")
    $miShowText.CheckOnClick = $true
    $miShowText.Checked = ($btn.Tag["ShowText"] -eq "true")
    $miShowText.Add_Click({
        param($s,$e)
        $btn = $s.GetCurrentParent().SourceControl -as [System.Windows.Forms.Button]
        $btn.Tag["ShowText"] = if ($s.Checked) { "true" } else { "false" }
        $form.SuspendLayout(); Save-Shortcuts; Update-Layout; $form.ResumeLayout()
    })

    $miAlignRight = New-Object System.Windows.Forms.ToolStripMenuItem("Align right")
    $miAlignRight.CheckOnClick = $true
    if ($DefAlignRight -eq "true") { $miAlignRight.Checked = $true }
    $miAlignRight.Add_Click({
        param($s,$e)
        $btn = $s.GetCurrentParent().SourceControl -as [System.Windows.Forms.Button]
        $btn.Tag["AlignRight"] = if ($s.Checked) { "true" } else { "false" }
        $form.SuspendLayout(); Save-Shortcuts; Update-Layout; $form.ResumeLayout()
    })

    $ctxMenu.Items.AddRange(@($miRemove, $miRename, $miOpenAdmin, $miShowText, $miAlignRight))
    $btn.ContextMenuStrip = $ctxMenu

    # Events
    $btn.Add_MouseDown({ param($s,$e) $s.Tag["DragStart"] = $e.Location })
    $btn.Add_MouseMove({
        param($s,$e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left -and $s.Tag["DragStart"]) {
            $start = $s.Tag["DragStart"] ; $deltaX = [Math]::Abs($e.Location.X - $start.X) ; $deltaY = [Math]::Abs($e.Location.Y - $start.Y)
            if ($deltaX -gt 5 -or $deltaY -gt 5) { $s.DoDragDrop($s, [System.Windows.Forms.DragDropEffects]::Move) }
        }
    })
    $btn.Add_MouseLeave({ param($s,$e) ; $tooltip.Hide($s) })

    $shortcutsPanel.Controls.Add($btn)
}

function Import-Shortcuts {
    param([string]$FilePath)
    if (-not $FilePath) {
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "INI Files|*.ini|All Files|*.*"
        if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $FilePath = $ofd.FileName
    }
    if ($argSilentImport -eq "/f") { $choice = [System.Windows.Forms.DialogResult]::Yes }
    elseif ((Test-Path -LiteralPath $shortcutsFile) -and ((Read-IniFile $shortcutsFile).Keys.Count -gt 0)) {
        $choice = [System.Windows.Forms.MessageBox]::Show("Add imported shortcuts to existing ones? (Yes = Append, No = Replace)", "Import Options", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
        if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
    } else { $choice = [System.Windows.Forms.DialogResult]::No }
    $form.SuspendLayout()
    $importData = Read-IniFile $FilePath
    if ($choice -eq [System.Windows.Forms.DialogResult]::No) { $shortcutsPanel.Controls.Clear() ; $existing = $importData } 
    else {
        $existing = if (Test-Path -LiteralPath $shortcutsFile) { Read-IniFile $shortcutsFile } else { @{} }
        $index = $existing.Keys.Count
        foreach ($section in $importData.Keys) { $existing["Shortcut$index"] = $importData[$section] ; $index++ }
    }
    $allButtons = $shortcutsPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.FilePath }
    $lookup = @{}
    $allButtons | ForEach-Object { if ($_.Tag.FilePath -and -not $lookup.ContainsKey($_.Tag.FilePath)) { $lookup[$_.Tag.FilePath] = $_ } }
    $buttonsToBlink = @()
    foreach ($section in $importData.Keys) {
        $entry = $importData[$section]
        if (Test-Path -LiteralPath $entry.Path) {
            if ($lookup.ContainsKey($entry.Path)) {
                if ($argSilentImport -ne "/f") {
                    $button = $lookup[$entry.Path]
                    if (-not ($buttonsToBlink | Where-Object { $_.Button -eq $button })) { $buttonsToBlink += [PSCustomObject]@{ Button=$button ; OriginalColor=$button.BackColor } }
                }
            } else { Add-ShortcutButton -FilePath $entry.Path -NoSave -DefOpenAsAdmin $entry.OpenAsAdmin -DefShowText $entry.ShowText -DefAlignRight $entry.AlignRight -DefDisplayName $entry.DisplayName }
        }
    }
    Update-Layout
    $form.ResumeLayout()
    if ($argSilentImport -ne "/f" -and $buttonsToBlink.Count -gt 0) {
        for ($i = 1; $i -le 6; $i++) {
            foreach ($item in $buttonsToBlink) { $btn = $item.Button ; $orig = $item.OriginalColor ; $btn.BackColor = if($btn.BackColor -eq 'Orange') { $orig } else { 'Orange' } }
            [System.Windows.Forms.Application]::DoEvents() ; Start-Sleep -Milliseconds 300
        }
        foreach ($item in $buttonsToBlink) { $item.Button.BackColor = $item.OriginalColor }
    }
    Save-Shortcuts
}

function Export-Shortcuts {
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "INI Files|*.ini|All Files|*.*"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Copy-Item -Path $shortcutsFile -Destination $sfd.FileName -Force }
}

# --- Options window ---
$settingsButton.Add_Click({
    $ctx = New-Object System.Windows.Forms.ContextMenuStrip
    $itmSettings = New-Object System.Windows.Forms.ToolStripMenuItem("Settings")
    $itmSettings.Add_Click({ Show-OptionsWindow })
    $itmRefresh = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh Toolbar")
    $itmRefresh.Add_Click({ $form.SuspendLayout(); Update-AppBarPosition -position $global:ToolbarLocation; Update-Layout; $form.ResumeLayout() })
    $itmClose = New-Object System.Windows.Forms.ToolStripMenuItem("Close Toolbar")
    $itmClose.Add_Click({ $form.Close() })
    $ctx.Items.AddRange(@($itmSettings, $itmRefresh, $itmClose))
    $pt = New-Object System.Drawing.Point(0, $settingsButton.Height)
    $ctx.Show($settingsButton, $pt)
})

function Show-OptionsWindow {
    $optForm = New-Object System.Windows.Forms.Form
    $optForm.SuspendLayout()

    $optForm.Text = "Options"
    $optForm.ClientSize = New-Object System.Drawing.Size(280,230)
    $optForm.StartPosition = 'CenterScreen'
    
    $gbLoc = New-Object System.Windows.Forms.GroupBox
    $gbLoc.Text = "Location"
    $gbLoc.Location = New-Object System.Drawing.Point(10,10)
    $gbLoc.Size = New-Object System.Drawing.Size(120,70)
    $radioTop = New-Object System.Windows.Forms.RadioButton
    $radioTop.Text = "Top"
    $radioTop.Location = New-Object System.Drawing.Point(10,20)
    $radioBottom = New-Object System.Windows.Forms.RadioButton
    $radioBottom.Text = "Bottom"
    $radioBottom.Location = New-Object System.Drawing.Point(10,40)
    if ($global:ToolbarLocation -eq "Top") { $radioTop.Checked = $true } else { $radioBottom.Checked = $true }
    $gbLoc.Controls.AddRange(@($radioTop, $radioBottom))
    $optForm.Controls.Add($gbLoc)

    $gbThick = New-Object System.Windows.Forms.GroupBox
    $gbThick.Text = "Size"
    $gbThick.Location = New-Object System.Drawing.Point(150,10)
    $gbThick.Size = New-Object System.Drawing.Size(120,100)
    $rSmall = New-Object System.Windows.Forms.RadioButton; $rSmall.Text = "Small"; $rSmall.Location = New-Object System.Drawing.Point(10,20)
    $rMedium = New-Object System.Windows.Forms.RadioButton; $rMedium.Text = "Medium"; $rMedium.Location = New-Object System.Drawing.Point(10,45)
    $rLarge = New-Object System.Windows.Forms.RadioButton; $rLarge.Text = "Large"; $rLarge.Location = New-Object System.Drawing.Point(10,70)
    switch ($global:ThicknessMode) { "Small" {$rSmall.Checked=$true} ; "Medium" {$rMedium.Checked=$true} ; "Large" {$rLarge.Checked=$true} }
    $gbThick.Controls.AddRange(@($rSmall, $rMedium, $rLarge))
    $optForm.Controls.Add($gbThick)

    $gbTheme = New-Object System.Windows.Forms.GroupBox
    $gbTheme.Text = "Theme"
    $gbTheme.Location = New-Object System.Drawing.Point(10,80)
    $gbTheme.Size = New-Object System.Drawing.Size(120,70)
    $rLight = New-Object System.Windows.Forms.RadioButton; $rLight.Text = "Light"; $rLight.Location = New-Object System.Drawing.Point(10,20)
    $rDark = New-Object System.Windows.Forms.RadioButton; $rDark.Text = "Dark"; $rDark.Location = New-Object System.Drawing.Point(10,40)
    if ($global:Theme -eq "Dark") { $rDark.Checked = $true } else { $rLight.Checked = $true }
    $gbTheme.Controls.AddRange(@($rLight, $rDark))
    $optForm.Controls.Add($gbTheme)

    $lblNew = New-Object System.Windows.Forms.Label; $lblNew.Text = "New shortcuts:"; $lblNew.Location = New-Object System.Drawing.Point(150,120)
    $optForm.Controls.Add($lblNew)
    $cbShowText = New-Object System.Windows.Forms.CheckBox; $cbShowText.Text = "Show text"; $cbShowText.Location = New-Object System.Drawing.Point(150,145)
    $cbShowText.Checked = ($global:NewShortcutShowText -eq "true")
    $optForm.Controls.Add($cbShowText)
    $cbOpenAdmin = New-Object System.Windows.Forms.CheckBox; $cbOpenAdmin.Text = "As admin"; $cbOpenAdmin.Location = New-Object System.Drawing.Point(150,169)
    $cbOpenAdmin.Checked = ($global:NewShortcutOpenAsAdmin -eq "true")
    $optForm.Controls.Add($cbOpenAdmin)
    $cbAlignRight = New-Object System.Windows.Forms.CheckBox; $cbAlignRight.Text = "Align right"; $cbAlignRight.Location = New-Object System.Drawing.Point(150,193)
    $cbAlignRight.Checked = ($global:NewShortcutAlignRight -eq "true")
    $optForm.Controls.Add($cbAlignRight)

    $btnImport = New-Object System.Windows.Forms.Button; $btnImport.Text = "Import shortcuts"; $btnImport.Location = New-Object System.Drawing.Point(10,155); $btnImport.Size = New-Object System.Drawing.Size(120,30)
    $btnImport.Add_Click({ Import-Shortcuts })
    $optForm.Controls.Add($btnImport)
    $btnExport = New-Object System.Windows.Forms.Button; $btnExport.Text = "Export shortcuts"; $btnExport.Location = New-Object System.Drawing.Point(10,185); $btnExport.Size = New-Object System.Drawing.Size(120,30)
    $btnExport.Add_Click({ Export-Shortcuts })
    $optForm.Controls.Add($btnExport)

    $radioTop.Add_CheckedChanged({ if ($radioTop.Checked) { $global:ToolbarLocation = "Top"; $form.SuspendLayout(); Update-AppBarPosition -position "Top"; $form.ResumeLayout() } })
    $radioBottom.Add_CheckedChanged({ if ($radioBottom.Checked) { $global:ToolbarLocation = "Bottom"; $form.SuspendLayout(); Update-AppBarPosition -position "Bottom"; $form.ResumeLayout() } })
    $rSmall.Add_CheckedChanged({ if ($rSmall.Checked) { $global:ThicknessMode = "Small"; $global:BarThickness = 25; $form.SuspendLayout(); Update-AppBarPosition -position $global:ToolbarLocation; Update-Layout; $form.ResumeLayout() } })
    $rMedium.Add_CheckedChanged({ if ($rMedium.Checked) { $global:ThicknessMode = "Medium"; $global:BarThickness = 32; $form.SuspendLayout(); Update-AppBarPosition -position $global:ToolbarLocation; Update-Layout; $form.ResumeLayout() } })
    $rLarge.Add_CheckedChanged({ if ($rLarge.Checked) { $global:ThicknessMode = "Large"; $global:BarThickness = 40; $form.SuspendLayout(); Update-AppBarPosition -position $global:ToolbarLocation; Update-Layout; $form.ResumeLayout() } })
    $rLight.Add_CheckedChanged({ if ($rLight.Checked) { $global:Theme = "Light"; $form.SuspendLayout(); Update-Layout; $form.ResumeLayout() } })
    $rDark.Add_CheckedChanged({ if ($rDark.Checked) { $global:Theme = "Dark"; $form.SuspendLayout(); Update-Layout; $form.ResumeLayout() } })
    $cbShowText.Add_CheckedChanged({ $global:NewShortcutShowText = if ($cbShowText.Checked) { "true" } else { "false" } })
    $cbOpenAdmin.Add_CheckedChanged({ $global:NewShortcutOpenAsAdmin = if ($cbOpenAdmin.Checked) { "true" } else { "false" } })
    $cbAlignRight.Add_CheckedChanged({ $global:NewShortcutAlignRight = if ($cbAlignRight.Checked) { "true" } else { "false" } })

    $optForm.ResumeLayout()
    $optForm.ShowDialog() | Out-Null
    $optForm.Dispose()
}

# --- Admin Helper ---
function Start-AdminHelper {
    if ($script:AdminHelperStarted) { return }
    $script:StandardUserSID = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $helperScript = @"
`$elevatedIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
`$pipeName = 'PSAdminHelperPipe'
while (`$true) {
    try {
        `$ps = New-Object System.IO.Pipes.PipeSecurity
        `$elevatedSID = `$elevatedIdentity.User
        `$parElevated = New-Object System.IO.Pipes.PipeAccessRule(`$elevatedSID, [System.IO.Pipes.PipeAccessRights]::ReadWrite, [System.Security.AccessControl.AccessControlType]::Allow)
        `$ps.AddAccessRule(`$parElevated)
        `$standardSID = New-Object System.Security.Principal.SecurityIdentifier('$script:StandardUserSID')
        `$parStandard = New-Object System.IO.Pipes.PipeAccessRule(`$standardSID, [System.IO.Pipes.PipeAccessRights]::ReadWrite, [System.Security.AccessControl.AccessControlType]::Allow)
        `$ps.AddAccessRule(`$parStandard)
        `$server = New-Object System.IO.Pipes.NamedPipeServerStream(`$pipeName, [System.IO.Pipes.PipeDirection]::In, 1, [System.IO.Pipes.PipeTransmissionMode]::Byte, [System.IO.Pipes.PipeOptions]::None, 1024, 1024, `$ps)
        `$server.WaitForConnection()
        `$reader = New-Object System.IO.StreamReader(`$server)
        `$cmd = `$reader.ReadLine()
        if (`$cmd) { Start-Process -FilePath `$cmd -ErrorAction Stop }
    }
    catch { }
    finally {
        if (`$reader) { `$reader.Dispose() }
        if (`$server) { if (`$server.IsConnected) { `$server.Disconnect() } ; `$server.Dispose() }
        Start-Sleep -Seconds 1
    }
}
"@
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "powershell.exe"
        $psi.Arguments = "-Nologo -Noprofile -ExecutionPolicy Bypass -WindowStyle Hidden -Command $helperScript"
        $psi.Verb = "runas"
        $psi.UseShellExecute = $true
        $script:AdminHelperProcess = [System.Diagnostics.Process]::Start($psi)
        $script:AdminHelperStarted = $true
        Start-Sleep -Seconds 2
    } catch { }
}

function Invoke-AdminCommand {
    param([string]$filePath)
    if (-not $script:AdminHelperStarted) { return }
    try {
        $client = New-Object System.IO.Pipes.NamedPipeClientStream(".", "PSAdminHelperPipe", [System.IO.Pipes.PipeDirection]::Out, [System.IO.Pipes.PipeOptions]::None)
        $client.Connect(15000)
        $writer = New-Object System.IO.StreamWriter($client)
        $writer.WriteLine($filePath)
        $writer.Flush()
    } catch { } finally {if ($writer) { $writer.Dispose() } if ($client) { $client.Dispose() }}
}

# --- Events ---
$shortcutsPanel.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.Button])) {$e.Effect = [System.Windows.Forms.DragDropEffects]::Move ; $dragIndicator.Visible=$true}
    elseif ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {$e.Effect = [System.Windows.Forms.DragDropEffects]::Copy}
    else { $e.Effect = [System.Windows.Forms.DragDropEffects]::None }
})

$shortcutsPanel.Add_DragOver({
    param($s, $e)
    function Get-IndicatorX($mouseX, $isRightGroup) {
        $alignedButtons=@($shortcutsPanel.Controls|Where-Object{$_-is[System.Windows.Forms.Button]-and$_.Tag-and$_.Tag.ContainsKey("FilePath")-and(($_.Tag["AlignRight"]-eq"true")-eq$isRightGroup)})
        if ($alignedButtons.Count -eq 0) { if ($isRightGroup) {return ($shortcutsPanel.ClientSize.Width - $dragIndicator.Width)} else {return 0} } 
        else {
            $closestButton = $null ; $minDistance = [int]::MaxValue
            foreach ($btn in $alignedButtons) {$distance = [Math]::Abs($mouseX - ($btn.Left + $btn.Width / 2)) ; if ($distance -lt $minDistance) { $minDistance=$distance ; $closestButton=$btn}}
            if ($mouseX -lt ($closestButton.Left+$closestButton.Width/2)) {return [int]($closestButton.Left-$dragIndicator.Width+2)} else {return [int]($closestButton.Left+$closestButton.Width-1)}
        }
    }
    if ($e.Data.GetDataPresent([System.Windows.Forms.Button])) {
        $dragButton = $e.Data.GetData([System.Windows.Forms.Button])
        if ($dragButton) {$dragIndicator.Location=New-Object System.Drawing.Point($(Get-IndicatorX $e.X $($e.X -gt ($shortcutsPanel.ClientSize.Width / 2))), 0) ; $dragIndicator.Visible=$true}
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
    }
    elseif ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        $dragIndicator.Location = New-Object System.Drawing.Point($(Get-IndicatorX $e.X $($e.X -gt ($shortcutsPanel.ClientSize.Width / 2))), 0)
        $dragIndicator.Visible  = $true
    }
})

function Update-ButtonPosition($button,$xCoordinate){
    if(-not $button){return}
    $shortcutsPanel.Controls.Remove($button)
    $allButtons=$shortcutsPanel.Controls|Where-Object{$_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.ContainsKey("FilePath")}
    $leftButtons=@($allButtons|Where-Object{$_.Tag["AlignRight"] -ne "true"}|Sort-Object Left)
    $rightButtons=@($allButtons|Where-Object{$_.Tag["AlignRight"] -eq "true"}|Sort-Object Left)
    $isRightGroup=($xCoordinate -gt ($shortcutsPanel.ClientSize.Width/2))
    $targetButtons=if($isRightGroup){,$rightButtons}else{,$leftButtons}
    $insertIndex=0
    foreach($currentButton in $targetButtons){
        if($xCoordinate -lt ($currentButton.Left+($currentButton.Width/2))){break}
        $insertIndex++
    }
    $newGroup=@()
    if($insertIndex -gt 0){$newGroup+=$targetButtons[0..($insertIndex-1)]}
    $newGroup+=$button
    if($insertIndex -lt $targetButtons.Count){$newGroup+=$targetButtons[$insertIndex..($targetButtons.Count-1)]}
    $newList=if($isRightGroup){$leftButtons+$newGroup}else{$newGroup+$rightButtons}
    $button.Tag["AlignRight"]=$isRightGroup.ToString().ToLower()
    ($button.ContextMenuStrip.Items|Where-Object{$_.Text -eq "Align right"}).Checked=$isRightGroup
    for($j=0;$j -lt $newList.Count;$j++){ $newList[$j].Tag.Order=$j }
    $allButtons|ForEach-Object{ $shortcutsPanel.Controls.Remove($_) }
    $newList|ForEach-Object{ $shortcutsPanel.Controls.Add($_) }
    Save-Shortcuts;Update-Layout
}

function Update-FileDrop($files){
    $allButtons=$shortcutsPanel.Controls|Where-Object{$_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.FilePath}
    $lookup=@{}
    $allButtons|ForEach-Object{ if($_.Tag.FilePath -and -not $lookup.ContainsKey($_.Tag.FilePath)){ $lookup[$_.Tag.FilePath]=$_ } }
    foreach($file in $files){
        $existingButton=$lookup[$file]
        if($existingButton){
            $originalColor=$existingButton.BackColor
            for($i=1;$i -le 6;$i++){
                $existingButton.BackColor=if($existingButton.BackColor -eq 'Orange'){$originalColor}else{'Orange'}
                [System.Windows.Forms.Application]::DoEvents() ; Start-Sleep -Milliseconds 300
            }
            $existingButton.BackColor=$originalColor
        } else {
            $form.SuspendLayout()
            Add-ShortcutButton -FilePath $file -Position $e.X ; Start-Sleep -Milliseconds 50
            $newButton=$shortcutsPanel.Controls| Where-Object{$_ -is [System.Windows.Forms.Button] -and $_.Tag -and $_.Tag.FilePath -eq $file} | Select-Object -First 1
            if($newButton){ Update-ButtonPosition $newButton $e.X }
            $form.ResumeLayout()
        }
    }
    $dragIndicator.Visible=$false
}

$shortcutsPanel.Add_DragDrop({
    param($s,$e)
    if($e.Data.GetDataPresent([System.Windows.Forms.Button])){
        $dragButton=$e.Data.GetData([System.Windows.Forms.Button])
        Update-ButtonPosition $dragButton $e.X
        $dragIndicator.Visible=$false
    }
    elseif($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)){
        $files=$e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        Update-FileDrop $files
    }
})

[Microsoft.Win32.SystemEvents]::add_DisplaySettingsChanged({
    $global:LastDisplaySettingsTime = Get-Date
    if (-not $global:DebounceActive) {
        $global:DebounceActive = $true
        while (((Get-Date) - $global:LastDisplaySettingsTime).TotalSeconds -lt 2) { Start-Sleep -Milliseconds 100; [System.Windows.Forms.Application]::DoEvents() }
        Update-AppBarPosition -position $global:ToolbarLocation
        Update-Layout
        $global:DebounceActive = $false
    }
})

# Load existing shortcuts
if (Test-Path -LiteralPath $shortcutsFile) {
    $ini = Read-IniFile $shortcutsFile
    foreach ($section in $ini.Keys) {
        $entry = $ini[$section]
        if (Test-Path -LiteralPath $entry.Path) {
            Add-ShortcutButton -FilePath $entry.Path -NoSave -DefOpenAsAdmin $entry.OpenAsAdmin -DefShowText $entry.ShowText -DefAlignRight $entry.AlignRight -DefDisplayName $entry.DisplayName
        }
    }
}

$form.Add_Shown({ 
    Update-AppBarPosition -position $global:ToolbarLocation
    Update-Layout
    [System.Windows.Forms.Application]::DoEvents()
    # Import shortcuts if script opened with argument path .ini
    if ($argIniShortcutsFile) {
        if (Test-Path -LiteralPath $argIniShortcutsFile -PathType Leaf) {
            $fileContent = Get-Content -LiteralPath $argIniShortcutsFile -Raw
            if (-not $fileContent.Trim()) {
                [System.Windows.Forms.MessageBox]::Show("Error: The file is empty.", "File Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            # Test .ini validity
            $sectionPattern = '^\[(?<section>Shortcut\d+)\]\s*\r?\n(?<content>.*?)(?=^\[Shortcut\d+\]\s*\r?\n|\z)'
            $regexOptions = [System.Text.RegularExpressions.RegexOptions]::Compiled -bor [System.Text.RegularExpressions.RegexOptions]::Multiline -bor [System.Text.RegularExpressions.RegexOptions]::Singleline
            $regex = [regex]::new($sectionPattern, $regexOptions)
            $sections = $regex.Matches($fileContent)
            if ($sections.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Error: The file does not contain any valid shortcut sections.", "Structure Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            $requiredKeys = @("ShowText", "AlignRight", "Path", "DisplayName", "Order", "OpenAsAdmin")
            foreach ($section in $sections) {
                $sectionName = $section.Groups["section"].Value
                $content = $section.Groups["content"].Value.Trim()
                $missingKeys = @()
                foreach ($key in $requiredKeys) { if ($content -notmatch "(?im)^\s*$key\s*=") { $missingKeys += $key } }
                if ($missingKeys.Count -gt 0) {
                    [System.Windows.Forms.MessageBox]::Show("Error: Section [$sectionName] is missing the following key(s): " + ($missingKeys -join ", "), "Structure Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
            }
            Import-Shortcuts -FilePath $argIniShortcutsFile
        }
    }
})

$form.Add_FormClosing({
    if ($global:BaselineWorkArea) { [AppBar]::SystemParametersInfo([AppBar]::SPI_SETWORKAREA, 0, [ref]$global:BaselineWorkArea, [AppBar]::SPIF_UPDATEINIFILE) | Out-Null }
    if ($script:AdminHelperProcess -and -not $script:AdminHelperProcess.HasExited) { $script:AdminHelperProcess.Kill() }
})

$form.ResumeLayout()
[System.Windows.Forms.Application]::Run($form)
