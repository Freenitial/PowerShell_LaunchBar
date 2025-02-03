<# ::

    REM Author  : Leo Gillet - Freenitial on GitHub
    REM Version : 0.7

    cls & @echo off & title PowerShell_LaunchBar
    copy /y "%~f0" "%TEMP%\%~n0.ps1" >NUL && powershell -Nologo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%TEMP%\%~n0.ps1"
    exit /b

#>

# --- DPI-Aware ---
Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool SetProcessDPIAware();
"@ -Name "DpiHelper" -Namespace "Win32"
[Win32.DpiHelper]::SetProcessDPIAware() | Out-Null

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class MonitorUtil {
    [DllImport("user32.dll")]
    public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);
}
"@
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DpiUtil {
    public const int MDT_EFFECTIVE_DPI = 0;
    [DllImport("Shcore.dll")]
    public static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);
}
"@
function Get-ScalingFactor {
    param([IntPtr]$hwnd)
    # MONITOR_DEFAULTTONEAREST = 2
    $hMonitor = [MonitorUtil]::MonitorFromWindow($hwnd, 2)
    $dpiX = 0; $dpiY = 0
    $result = [DpiUtil]::GetDpiForMonitor($hMonitor, [DpiUtil]::MDT_EFFECTIVE_DPI, [ref]$dpiX, [ref]$dpiY)
    if ($result -eq 0) {
         return [double]$dpiX / 96.0
    }
    else {
         return 1.0
    }
}

Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# --- INI Parsing and Writing Functions ---
function Read-IniFile ($Path) {
    $ini = @{}
    if (Test-Path -LiteralPath $Path) {
        $section = ""
        foreach ($line in Get-Content $Path) {
            $line = $line.Trim()
            if ($line -match '^\[(.+)\]') { $section = $Matches[1]; $ini[$section] = @{}; continue }
            if ($line -match '^(.*?)=(.*)$') {
                $key = $Matches[1].Trim(); $value = $Matches[2].Trim()
                if ($section) { $ini[$section][$key] = $value } else { $ini[$key] = $value }
            }
        }
    }
    return $ini
}
function Write-IniFile ($Path, $Data) {
    $lines = @()
    foreach ($section in $Data.Keys) {
        if ($Data[$section] -is [hashtable]) {
            $lines += "[$section]"
            foreach ($k in $Data[$section].Keys) { $lines += "$k=$($Data[$section][$k])" }
        } else { $lines += "$section=$($Data[$section])" }
    }
    Set-Content -Path $Path -Value $lines
}

# --- Native Classes (AppBar, MessageHook, IconExtractor) ---
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class AppBar {
    public const int ABM_NEW = 0, ABM_REMOVE = 1, ABM_SETPOS = 2, ABM_QUERYPOS = 3;
    public const int ABE_TOP = 1, ABE_BOTTOM = 3;
    public const int SPI_SETWORKAREA = 47, SPI_GETWORKAREA = 48;
    public const int SPIF_UPDATEINIFILE = 0x01, SPIF_SENDCHANGE = 0x02;
    public const int WM_SETTINGCHANGE = 0x1A;
    [StructLayout(LayoutKind.Sequential)]
    public struct APPBARDATA {
        public int cbSize;
        public IntPtr hWnd;
        public int uCallbackMessage;
        public int uEdge;
        public Rect rc;
        public IntPtr lParam;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct Rect { public int left, top, right, bottom; }
    [DllImport("shell32.dll")]
    public static extern int SHAppBarMessage(int dwMessage, ref APPBARDATA pData);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SystemParametersInfo(int uiAction, int uiParam, ref Rect pvParam, int fWinIni);
}
"@
Add-Type @"
using System;
using System.Windows.Interop;
public class MessageHook {
    public static void AddHook(IntPtr hwnd, int callbackMessage, System.Action<int, IntPtr, IntPtr> handler) {
        HwndSource source = HwndSource.FromHwnd(hwnd);
        if (source != null) {
            source.AddHook(new HwndSourceHook(delegate(IntPtr hwndHook, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
                if (msg == callbackMessage) { handler(msg, wParam, lParam); handled = true; }
                return IntPtr.Zero;
            }));
        }
    }
}
"@ -ReferencedAssemblies "PresentationCore","PresentationFramework","WindowsBase"
Add-Type -TypeDefinition @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
public class IconExtractor {
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto)]
    public struct SHFILEINFO {
        public IntPtr hIcon; public int iIcon; public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst=260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst=80)]
        public string szTypeName;
    };
    public const uint SHGFI_ICON = 0x100, SHGFI_SMALLICON = 0x1;
    [DllImport("shell32.dll", CharSet=CharSet.Auto)]
    public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);
    public static Icon GetIcon(string fileName) {
        SHFILEINFO shinfo = new SHFILEINFO();
        IntPtr hImg = SHGetFileInfo(fileName, 0, ref shinfo, (uint)System.Runtime.InteropServices.Marshal.SizeOf(shinfo), SHGFI_ICON | SHGFI_SMALLICON);
        if (shinfo.hIcon != IntPtr.Zero) { Icon icon = Icon.FromHandle(shinfo.hIcon); return (Icon)icon.Clone(); }
        return null;
    }
}
"@ -ReferencedAssemblies "System.Drawing"

# --- Global Variables and Paths ---
$savePath = Join-Path $env:LOCALAPPDATA "Powershell_Toolbar"
$shortcutsFile = Join-Path $savePath "shortcuts.ini"
$settingsFile  = Join-Path $savePath "settings.ini"
if (-not (Test-Path -LiteralPath $savePath)) { New-Item -ItemType Directory -Path $savePath | Out-Null }
$global:Settings = @{
    ToolbarLocation = "Top"
    ThicknessMode   = "Small"
    Theme           = "Light"
    NewShortcutShowText = "true"
    NewShortcutOpenAsAdmin = "false"
    NewShortcutAlignRight = "false"
}
if (Test-Path -LiteralPath $settingsFile) {
    $ini = Read-IniFile $settingsFile
    if ($ini.ContainsKey("Settings")) {
        foreach ($k in @($global:Settings.Keys)) {
            if ($ini["Settings"].ContainsKey($k)) { $global:Settings[$k] = $ini["Settings"][$k] }
        }
    }
}
$global:ToolbarLocation = $global:Settings["ToolbarLocation"]
$global:ThicknessMode   = $global:Settings["ThicknessMode"]
$global:Theme           = $global:Settings["Theme"]
$global:NewShortcutShowText = ([string]$global:Settings["NewShortcutShowText"]).ToLower()
$global:NewShortcutOpenAsAdmin = ([string]$global:Settings["NewShortcutOpenAsAdmin"]).ToLower()
$global:NewShortcutAlignRight = ([string]$global:Settings["NewShortcutAlignRight"]).ToLower()
switch ($global:ThicknessMode) {
    "Small"   { $global:BaseThickness = 25 }
    "Medium"  { $global:BaseThickness = 32 }
    "Large"   { $global:BaseThickness = 40 }
    default   { $global:BaseThickness = 25 }
}
$global:BarThickness = $global:BaseThickness  # en DIP ; les ajustements DPI se font lors du calcul de la zone
$global:BaseWorkArea = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

# --- Admin Helper Variables ---
function Start-AdminHelper {
    if ($script:AdminHelperStarted) { return }
    $helperScript = @"
`$pipeName = 'PSAdminHelperPipe'
while (`$true) {
    try {
        `$ps = New-Object System.IO.Pipes.PipeSecurity
        `$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        `$par = New-Object System.IO.Pipes.PipeAccessRule(`$currentUser, [System.IO.Pipes.PipeAccessRights]::ReadWrite, [System.Security.AccessControl.AccessControlType]::Allow)
        `$ps.AddAccessRule(`$par)
        `$server = New-Object System.IO.Pipes.NamedPipeServerStream(`$pipeName, [System.IO.Pipes.PipeDirection]::In, 1, [System.IO.Pipes.PipeTransmissionMode]::Byte, [System.IO.Pipes.PipeOptions]::None, 1024, 1024, `$ps)
        `$server.WaitForConnection()
        `$reader = New-Object System.IO.StreamReader(`$server)
        `$cmd = `$reader.ReadLine()
        if (`$cmd) { Start-Process -FilePath `$cmd -ErrorAction Stop }
    }
    catch { Write-Error `$_.Exception.Message }
    finally {
        if (`$reader) { `$reader.Dispose() }
        if (`$server) { 
            if (`$server.IsConnected) { `$server.Disconnect() }
            `$server.Dispose()
        }
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
    }
    catch { Write-Warning "Error while starting admin helper : $_" }
}
function Invoke-AdminCommand {
    param([string]$filePath)
    if (-not $script:AdminHelperStarted) { Write-Warning "Admin helper not started"; return }
    try {
        $client = New-Object System.IO.Pipes.NamedPipeClientStream(".", "PSAdminHelperPipe", [System.IO.Pipes.PipeDirection]::Out, [System.IO.Pipes.PipeOptions]::None)
        $client.Connect(15000)
        $writer = New-Object System.IO.StreamWriter($client)
        $writer.WriteLine($filePath)
        $writer.Flush()
    }
    catch { Write-Warning "Error executing admin command : $_" }
    finally {
        if ($writer) { $writer.Dispose() }
        if ($client) { $client.Dispose() }
    }
}

# --- Load Main Window XAML ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PowerShell_LaunchBar" WindowStyle="None" ResizeMode="NoResize" Topmost="True"
        WindowStartupLocation="Manual" Background="#FFEEEEEE" ShowInTaskbar="False" AllowDrop="True">
    <Window.Resources>
        <Style x:Key="CustomButtonStyle" TargetType="Button">
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Setter Property="Padding" Value="5,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" 
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="DarkCyan"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <DockPanel LastChildFill="True">
        <Button x:Name="SettingsButton" DockPanel.Dock="Right" Margin="2,0,2,0">
            <Path Data="M9.405,0.5 C9.783,0.5 10.126,0.733 10.263,1.088L10.823,2.77C11.114,2.92 11.393,3.093 11.656,3.286L13.378,2.86C13.738,2.77 14.114,2.911 14.309,3.223L15.809,5.777C16.004,6.089 15.967,6.483 15.715,6.753L14.463,8.139C14.484,8.425 14.484,8.714 14.463,9L15.715,10.386C15.967,10.656 16.004,11.05 15.809,11.362L14.309,13.916C14.114,14.228 13.738,14.369 13.378,14.279L11.656,13.853C11.393,14.046 11.114,14.219 10.823,14.369L10.263,16.051C10.126,16.406 9.783,16.639 9.405,16.639L6.405,16.639C6.027,16.639 5.684,16.406 5.547,16.051L4.987,14.369C4.696,14.219 4.417,14.046 4.154,13.853L2.432,14.279C2.072,14.369 1.696,14.228 1.501,13.916L0.001,11.362C-0.194,11.05 -0.157,10.656 0.095,10.386L1.347,9C1.326,8.714 1.326,8.425 1.347,8.139L0.095,6.753C-0.157,6.483 -0.194,6.089 0.001,5.777L1.501,3.223C1.696,2.911 2.072,2.77 2.432,2.86L4.154,3.286C4.417,3.093 4.696,2.92 4.987,2.77L5.547,1.088C5.684,0.733 6.027,0.5 6.405,0.5L9.405,0.5z M7.905,5.5C6.248,5.5 4.905,6.843 4.905,8.5 4.905,10.157 6.248,11.5 7.905,11.5 9.562,11.5 10.905,10.157 10.905,8.5 10.905,6.843 9.562,5.5 7.905,5.5z"
                  Fill="Gray" Stretch="Uniform" Margin="3"/>
        </Button>
        <ScrollViewer HorizontalScrollBarVisibility="Disabled" VerticalScrollBarVisibility="Disabled" Margin="0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Horizontal" x:Name="LeftShortcutStack" Grid.Column="0" VerticalAlignment="Stretch">
                    <StackPanel.Resources>
                        <Style TargetType="Button" BasedOn="{StaticResource CustomButtonStyle}"/>
                    </StackPanel.Resources>
                </StackPanel>
                <StackPanel Orientation="Horizontal" x:Name="RightShortcutStack" Grid.Column="1" VerticalAlignment="Stretch">
                    <StackPanel.Resources>
                        <Style TargetType="Button" BasedOn="{StaticResource CustomButtonStyle}"/>
                    </StackPanel.Resources>
                </StackPanel>
            </Grid>
        </ScrollViewer>
    </DockPanel>
</Window>
"@
$reader = New-Object System.Xml.XmlNodeReader($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$settingsButton = $window.FindName("SettingsButton")
$leftStack = $window.FindName("LeftShortcutStack")
$rightStack = $window.FindName("RightShortcutStack")

# --- Update Button Tooltip ---
function Update-ButtonToolTip($btn) {
    if ($btn.ToolTip -is [System.Windows.Controls.ToolTip]) {
        $tt = $btn.ToolTip
        $tt.Background = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0,0,0)))
        $tt.Foreground = [System.Windows.Media.Brushes]::White
    }
}

Add-Type @'
using System; 
using System.Runtime.InteropServices;
using System.Drawing;
public class DPI {  
  [DllImport("gdi32.dll")]
  static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
  public enum DeviceCap {
      VERTRES = 10,
      DESKTOPVERTRES = 117
  } 
  public static float scaling() {
      Graphics g = Graphics.FromHwnd(IntPtr.Zero);
      IntPtr desktop = g.GetHdc();
      int LogicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.VERTRES);
      int PhysicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.DESKTOPVERTRES);
      return (float)PhysicalScreenHeight / (float)LogicalScreenHeight;
  }
}
'@ -ReferencedAssemblies 'System.Drawing.dll' -ErrorAction Stop

function Refresh-WorkArea {
    $appBarData = New-Object AppBar+APPBARDATA
    $appBarData.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf($appBarData)
    $appBarData.hWnd = (New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle
    [AppBar]::SHAppBarMessage([AppBar]::ABM_REMOVE, [ref]$appBarData)
    Start-Sleep -Milliseconds 100
    Update-AppBarPosition -position $global:ToolbarLocation
}

# --- AppBar Positioning (DPI-Aware) ---
function Update-AppBarPosition {
    param([string]$position)
    $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle
    $dummy = New-Object AppBar+APPBARDATA
    $dummy.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf($dummy)
    $dummy.hWnd = $hwnd
    [AppBar]::SHAppBarMessage([AppBar]::ABM_REMOVE, [ref]$dummy)

    $currentWorkArea = New-Object AppBar+Rect
    [AppBar]::SystemParametersInfo([AppBar]::SPI_GETWORKAREA, 0, [ref]$currentWorkArea, 0)

    $dpiFactor = Get-ScalingFactor $hwnd
    $baseArea = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

    $winWidth = $baseArea.Width / $dpiFactor
    $window.Width = $winWidth
    $window.Height = $global:BarThickness
    $window.Left = $baseArea.Left / $dpiFactor

    switch ($position) {
        "Top" {
            $window.Top = 0
            $edge = [AppBar]::ABE_TOP
            $offset = [int]($global:BarThickness * $dpiFactor)
            $newWorkArea = New-Object AppBar+Rect
            $newWorkArea.left   = $baseArea.Left
            $newWorkArea.top    = $baseArea.Top + $offset
            $newWorkArea.right  = $baseArea.Right
            $newWorkArea.bottom = $baseArea.Bottom
        }
        "Bottom" {
            $winHeight = $baseArea.Height / $dpiFactor
            $window.Top = $winHeight - $global:BarThickness
            $edge = [AppBar]::ABE_BOTTOM
            $offset = [int]($global:BarThickness * $dpiFactor)
            $newWorkArea = New-Object AppBar+Rect
            $newWorkArea.left   = $baseArea.Left
            $newWorkArea.top    = $baseArea.Top
            $newWorkArea.right  = $baseArea.Right
            $newWorkArea.bottom = $baseArea.Bottom - $offset
        }
    }
    $appBarRect = New-Object AppBar+Rect
    $appBarRect.left   = [int]($window.Left * $dpiFactor)
    $appBarRect.top    = [int]($window.Top * $dpiFactor)
    $appBarRect.right  = [int](($window.Left + $window.Width) * $dpiFactor)
    $appBarRect.bottom = [int](($window.Top + $window.Height) * $dpiFactor)
    $appBarData = New-Object AppBar+APPBARDATA
    $appBarData.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf($appBarData)
    $appBarData.hWnd = $hwnd
    $appBarData.uEdge = $edge
    $appBarData.rc = $appBarRect
    $global:CallbackMessage = [AppBar]::SHAppBarMessage([AppBar]::ABM_NEW, [ref]$appBarData)
    Start-Sleep -Milliseconds 200
    [AppBar]::SHAppBarMessage([AppBar]::ABM_SETPOS, [ref]$appBarData)
    [AppBar]::SystemParametersInfo([AppBar]::SPI_SETWORKAREA, 0, [ref]$newWorkArea, [AppBar]::SPIF_UPDATEINIFILE)
}

# --- Save Settings ---
function Save-Settings {
    $global:Settings["ToolbarLocation"] = $global:ToolbarLocation
    $global:Settings["ThicknessMode"]   = $global:ThicknessMode
    $global:Settings["Theme"]           = $global:Theme
    $global:Settings["NewShortcutShowText"] = $global:NewShortcutShowText
    $global:Settings["NewShortcutOpenAsAdmin"] = $global:NewShortcutOpenAsAdmin
    $global:Settings["NewShortcutAlignRight"] = $global:NewShortcutAlignRight
    Write-IniFile $settingsFile @{ Settings = $global:Settings }
}

# --- Update Theme Appearance ---
function Update-ThemeAppearance {
    if ($global:Theme -eq "Dark") {
        $window.Background = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(51,51,51)))
        $settingsButton.Background = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(51,51,51)))
    } else {
        $window.Background = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(238,238,238)))
        $settingsButton.Background = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(238,238,238)))
    }
    foreach ($stack in @($leftStack, $rightStack)) {
        foreach ($child in $stack.Children) {
            if ($child.Tag -ne "Indicator") {
                if ($global:Theme -eq "Dark") {
                    $child.Background = [System.Windows.Media.Brushes]::DarkSlateGray
                    $child.Foreground = [System.Windows.Media.Brushes]::White
                    $child.BorderBrush = [System.Windows.Media.Brushes]::LightGray
                    if ($child.Content -is [System.Windows.Controls.StackPanel]) {
                        foreach ($c in $child.Content.Children) { if ($c -is [System.Windows.Controls.TextBlock]) { $c.Foreground = [System.Windows.Media.Brushes]::White } }
                    }
                } else {
                    $child.Background = [System.Windows.Media.Brushes]::White
                    $child.Foreground = [System.Windows.Media.Brushes]::Black
                    $child.BorderBrush = [System.Windows.Media.Brushes]::Gray
                    if ($child.Content -is [System.Windows.Controls.StackPanel]) {
                        foreach ($c in $child.Content.Children) { if ($c -is [System.Windows.Controls.TextBlock]) { $c.Foreground = [System.Windows.Media.Brushes]::Black } }
                    }
                }
                if ($child.ToolTip) { Update-ButtonToolTip $child }
            }
        }
    }
}

# --- Update Shortcut Buttons Appearance ---
function Update-ShortcutButtonsAppearance {
    switch ($global:ThicknessMode) {
        "Small"   { $iconSize = 16; $fontSize = 12 }
        "Medium"  { $iconSize = 22; $fontSize = 14 }
        "Large"   { $iconSize = 28; $fontSize = 16 }
        default   { $iconSize = 16; $fontSize = 12 }
    }
    foreach ($stack in @($leftStack, $rightStack)) {
        foreach ($child in $stack.Children) {
            if ($child.Tag -ne "Indicator") {
                $child.Height = $global:BarThickness
                $child.BorderThickness = New-Object System.Windows.Thickness(1)
                $child.BorderBrush = if ($global:Theme -eq "Dark") { [System.Windows.Media.Brushes]::LightGray } else { [System.Windows.Media.Brushes]::Gray }
                if ($child.Content -is [System.Windows.Controls.StackPanel]) {
                    $sp = $child.Content
                    if ($sp.Children.Count -ge 2) {
                        $img = $sp.Children[0]
                        $txt = $sp.Children[1]
                        $img.Width = $iconSize; $img.Height = $iconSize; $txt.FontSize = $fontSize
                    }
                }
            }
        }
    }
}

# --- Insertion Indicator for Drag-and-Drop ---
function Show-InsertionIndicator {
    param([Parameter(Mandatory=$true)] $Stack, [int]$Index, [double]$MouseX, $Target)
    $finalIndex = $Index
    if ($Target -and $MouseX -gt ($Target.ActualWidth/2)) { $finalIndex++ }
    if (-not $global:InsertIndicator) {
        $global:InsertIndicator = New-Object System.Windows.Controls.Border
        $global:InsertIndicator.Width = 3
        $global:InsertIndicator.Background = [System.Windows.Media.Brushes]::Orange
        $global:InsertIndicator.Tag = "Indicator"
    }
    if ($Stack.Children.Contains($global:InsertIndicator)) {
        $cur = $Stack.Children.IndexOf($global:InsertIndicator)
        if ($cur -eq $finalIndex) { return }
        [void]$Stack.Children.Remove($global:InsertIndicator)
    }
    [void]$Stack.Children.Insert($finalIndex, $global:InsertIndicator)
}
function Remove-InsertionIndicator {
    foreach ($stack in @($leftStack, $rightStack)) {
        if ($global:InsertIndicator -and $stack.Children.Contains($global:InsertIndicator)) {
            [void]$stack.Children.Remove($global:InsertIndicator)
        }
    }
}

# --- Add Shortcut Button ---
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
    $displayName = if ($DefDisplayName -ne "") { $DefDisplayName } else { Split-Path $FilePath -Leaf }
    try { $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($FilePath) } catch { $icon = $null }
    if (-not $icon) { $icon = [IconExtractor]::GetIcon($FilePath) }
    if ($icon) {
        $ms = New-Object System.IO.MemoryStream
        $icon.ToBitmap().Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $ms.Position = 0
        $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit(); $bmp.StreamSource = $ms; $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad; $bmp.EndInit()
        $imgSource = $bmp
    }
    $button = New-Object System.Windows.Controls.Button
    $button.Tag = @{
        FilePath = $FilePath;
        OpenAsAdmin = $DefOpenAsAdmin;
        ShowText = $DefShowText;
        AlignRight = $DefAlignRight;
        DisplayName = $displayName
    }
    $button.VerticalAlignment = "Stretch"
    $button.VerticalContentAlignment = "Center"
    $button.HorizontalContentAlignment = "Left"
    $button.Height = $global:BarThickness
    $button.AllowDrop = $true
    $button.BorderThickness = New-Object System.Windows.Thickness(1)
    $button.BorderBrush = if ($global:Theme -eq "Dark") { [System.Windows.Media.Brushes]::LightGray } else { [System.Windows.Media.Brushes]::Gray }
    $contextMenu = New-Object System.Windows.Controls.ContextMenu
    $removeItem = New-Object System.Windows.Controls.MenuItem; $removeItem.Header = "Remove"
    $removeItem.Add_Click({
        param($s, $e)
        $btn = $s.Parent.PlacementTarget
        if ($leftStack.Children.Contains($btn)) { [void]$leftStack.Children.Remove($btn) }
        elseif ($rightStack.Children.Contains($btn)) { [void]$rightStack.Children.Remove($btn) }
        Save-Shortcuts; Update-ShortcutStackMargins
    })
    $renameItem = New-Object System.Windows.Controls.MenuItem; $renameItem.Header = "Rename"
    $renameItem.Add_Click({
        param($s, $e)
        $btn = $s.Parent.PlacementTarget
        $currentName = $btn.Tag.DisplayName
        $inputForm = New-Object System.Windows.Forms.Form
        $inputForm.StartPosition = "CenterScreen"; $inputForm.Size = New-Object System.Drawing.Size(300,150)
        $inputForm.Text = "Rename Shortcut"
        $txtBox = New-Object System.Windows.Forms.TextBox; $txtBox.Text = $currentName
        $txtBox.Location = New-Object System.Drawing.Point(10,20); $txtBox.Size = New-Object System.Drawing.Size(260,20)
        $okButton = New-Object System.Windows.Forms.Button; $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(200,60); $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $inputForm.Controls.AddRange(@($txtBox, $okButton)); $inputForm.AcceptButton = $okButton
        if ($inputForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $newName = $txtBox.Text.Trim()
            if ($newName -ne "") {
                $btn.Tag.DisplayName = $newName
                if ($btn.Content -is [System.Windows.Controls.StackPanel]) {
                    $sp = $btn.Content
                    if ($sp.Children.Count -ge 2) { $sp.Children[1].Text = " $newName" }
                } else { $btn.Content = $newName }
                Save-Shortcuts
            }
        }
        $inputForm.Dispose()
    })
    $openAdminItem = New-Object System.Windows.Controls.MenuItem; $openAdminItem.Header = "Open as admin"
    $openAdminItem.IsCheckable = $true; $openAdminItem.IsChecked = ($button.Tag.OpenAsAdmin -eq "true")
    $openAdminItem.Add_Click({
        param($s, $e)
        $btn = $s.Parent.PlacementTarget
        $btn.Tag["OpenAsAdmin"] = if ($s.IsChecked) { "true" } else { "false" }
        Save-Shortcuts
    })
    $showTextItem = New-Object System.Windows.Controls.MenuItem; $showTextItem.Header = "Show text"
    $showTextItem.IsCheckable = $true; $showTextItem.IsChecked = ($button.Tag.ShowText -eq "true")
    $showTextItem.Add_Click({
        param($s, $e)
        $btn = $s.Parent.PlacementTarget
        if ($btn) {
            if ($s.IsChecked) {
                $btn.Tag["ShowText"] = "true"
                $btn.ToolTip = $null
                if ($btn.Content -is [System.Windows.Controls.StackPanel]) { $sp = $btn.Content ; if ($sp.Children.Count -ge 2) { $sp.Children[1].Visibility = "Visible" } }
            } else {
                $btn.Tag["ShowText"] = "false"
                if ($btn.Content -is [System.Windows.Controls.StackPanel]) {
                    $sp = $btn.Content
                    if ($sp.Children.Count -ge 2) { $sp.Children[1].Visibility = "Collapsed" }
                }
                try {
                    $toolTip = New-Object System.Windows.Controls.ToolTip
                    $toolTip.Content = $btn.Tag.DisplayName
                    $toolTip.Background = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0,0,0)))
                    $toolTip.Foreground = [System.Windows.Media.Brushes]::White
                    $btn.ToolTip = $toolTip
                    [System.Windows.Controls.ToolTipService]::SetInitialShowDelay($btn, 0)
                } catch { Write-Warning "Error while creating ToolTip: $_" }
            }
            Save-Shortcuts
        }
    })
    $alignRightItem = New-Object System.Windows.Controls.MenuItem; $alignRightItem.Header = "Align right"
    $alignRightItem.IsCheckable = $true; $alignRightItem.IsChecked = ($button.Tag.AlignRight -eq "true")
    $alignRightItem.Add_Click({
        param($s, $e)
        $btn = $s.Parent.PlacementTarget
        if ($s.IsChecked) {
            $btn.Tag["AlignRight"] = "true"
            if (-not $rightStack.Children.Contains($btn)) {
                if ($leftStack.Children.Contains($btn)) { [void]$leftStack.Children.Remove($btn) }
                $rightStack.Children.Add($btn) | Out-Null
            }
        } else {
            $btn.Tag["AlignRight"] = "false"
            if (-not $leftStack.Children.Contains($btn)) {
                if ($rightStack.Children.Contains($btn)) { [void]$rightStack.Children.Remove($btn) }
                $leftStack.Children.Add($btn) | Out-Null
            }
        }
        Save-Shortcuts; Update-ShortcutStackMargins
    })
    foreach ($item in @($removeItem, $renameItem, $openAdminItem, $showTextItem, $alignRightItem)) { $contextMenu.Items.Add($item) | Out-Null }
    $button.ContextMenu = $contextMenu
    if ($imgSource) {
        switch ($global:ThicknessMode) {
            "Small" { $iconSize = 16; $fontSize = 12 }
            "Medium" { $iconSize = 22; $fontSize = 14 }
            "Large" { $iconSize = 28; $fontSize = 16 }
            default { $iconSize = 16; $fontSize = 12 }
        }
        $imgCtrl = New-Object System.Windows.Controls.Image; $imgCtrl.Source = $imgSource
        $imgCtrl.IsHitTestVisible = $false; $imgCtrl.Width = $iconSize; $imgCtrl.Height = $iconSize; $imgCtrl.VerticalAlignment = "Center"
        $stack = New-Object System.Windows.Controls.StackPanel; $stack.Orientation = "Horizontal"
        $stack.HorizontalAlignment = "Left"; $stack.VerticalAlignment = "Center"
        $stack.Children.Add($imgCtrl) | Out-Null 
        $txtBlock = New-Object System.Windows.Controls.TextBlock; $txtBlock.Text = " $displayName"
        $txtBlock.VerticalAlignment = "Center"; $txtBlock.HorizontalAlignment = "Left"
        $txtBlock.IsHitTestVisible = $false; $txtBlock.FontSize = $fontSize
        if ($button.Tag.ShowText -ne "true") {
            $txtBlock.Visibility = "Collapsed"
            $toolTip = New-Object System.Windows.Controls.ToolTip; $toolTip.Content = $displayName
            $toolTip.Background = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0,0,0)))
            $toolTip.Foreground = [System.Windows.Media.Brushes]::White
            [System.Windows.Controls.ToolTipService]::SetInitialShowDelay($button, 0); $button.ToolTip = $toolTip
        }
        $stack.Children.Add($txtBlock) | Out-Null ; $button.Content = $stack
    } else {
        $button.Content = $displayName
        if ($button.Tag.ShowText -ne "true") {
            $toolTip = New-Object System.Windows.Controls.ToolTip; $toolTip.Content = $displayName
            $toolTip.Background = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0,0,0)))
            $toolTip.Foreground = [System.Windows.Media.Brushes]::White
            [System.Windows.Controls.ToolTipService]::SetInitialShowDelay($button, 0); $button.ToolTip = $toolTip
        }
    }
    $button.Add_Click({
        param($s, $e)
        $opts = $s.Tag
        if ($opts.FilePath -and (Test-Path -LiteralPath $opts.FilePath)) {
            if ($opts.OpenAsAdmin -eq "true") {
                if (-not $global:AdminHelperStarted) { Start-AdminHelper }
                Invoke-AdminCommand $opts.FilePath
            } else { Start-Process -FilePath $opts.FilePath }
        }
    })
    $script:isDragging = $false; $script:dragStartPoint = $null
    $button.Add_PreviewMouseDown({
        param($s, $e)
        if ($e.LeftButton -eq [System.Windows.Input.MouseButtonState]::Pressed) {
            $script:dragStartPoint = $e.GetPosition($s)
            $script:isDragging = $false
        }
    })
    $button.Add_PreviewMouseMove({
        param($s, $e)
        if ($e.LeftButton -eq [System.Windows.Input.MouseButtonState]::Pressed -and -not $script:isDragging -and $script:dragStartPoint) {
            $currentPoint = $e.GetPosition($s)
            $dragDelta = [System.Windows.Point]::Subtract($currentPoint, $script:dragStartPoint)
            if ([Math]::Abs($dragDelta.X) -gt 5 -or [Math]::Abs($dragDelta.Y) -gt 5) {
                $script:isDragging = $true
                $dragData = New-Object Windows.DataObject
                $dragData.SetData("ButtonSource", $s)
                if (-not ($s.Tag -is [hashtable])) { $s.Tag = @{ FilePath = $s.Tag; OriginalIndex = $s.Parent.Children.IndexOf($s) } }
                [System.Windows.DragDrop]::DoDragDrop($s, $dragData, [System.Windows.DragDropEffects]::Move) | Out-Null
                $script:isDragging = $false; $script:dragStartPoint = $null
            }
        }
    })
    $button.Add_DragOver({
        param($s, $e)
        if ($e.Data.GetDataPresent("ButtonSource")) {
            $parent = $s.Parent
            $mousePos = $e.GetPosition($s)
            $targetIndex = 0
            foreach ($child in $parent.Children) {
                if ($child -ne $s -and $child.Tag -ne "Indicator") { $targetIndex++ }
                elseif ($child -eq $s) { break }
            }
            Show-InsertionIndicator -Stack $parent -Index $targetIndex -MouseX $mousePos.X -Target $s
            $e.Effects = [System.Windows.DragDropEffects]::Move; $e.Handled = $true
        }
    })
    $button.Add_DragLeave({ })
    $button.Add_Drop({
        param($s, $e)
        if ($e.Data.GetDataPresent("ButtonSource")) {
            $sourceButton = $e.Data.GetData("ButtonSource")
            $parent = $s.Parent
            if ($global:InsertIndicator -and $parent.Children.Contains($global:InsertIndicator)) { $targetIndex = $parent.Children.IndexOf($global:InsertIndicator) } 
            else { $targetIndex = $parent.Children.IndexOf($s) ; if ($e.GetPosition($s).X -gt ($s.ActualWidth/2)) { $targetIndex++ } }
            Remove-InsertionIndicator
            $sourceIndex = $sourceButton.Parent.Children.IndexOf($sourceButton)
            if ($sourceIndex -ge 0) { [void]$sourceButton.Parent.Children.RemoveAt($sourceIndex) ; if ($targetIndex -gt $sourceIndex) { $targetIndex-- } }
            [void]$parent.Children.Insert($targetIndex, $sourceButton)
            if ($sourceButton.Tag -is [hashtable] -and $sourceButton.Tag.ContainsKey("OriginalIndex")) {
                $sourceButton.Tag = @{
                    FilePath = $sourceButton.Tag.FilePath;
                    OpenAsAdmin = $sourceButton.Tag.OpenAsAdmin;
                    ShowText = $sourceButton.Tag.ShowText;
                    AlignRight = $sourceButton.Tag.AlignRight;
                    DisplayName = $sourceButton.Tag.DisplayName
                }
            }
            Save-Shortcuts; Update-ShortcutStackMargins; $e.Handled = $true
        }
    })
    if ($global:Theme -eq "Dark") {
        $button.Background = [System.Windows.Media.Brushes]::DarkSlateGray
        $button.Foreground = [System.Windows.Media.Brushes]::White
        $button.BorderBrush = [System.Windows.Media.Brushes]::LightGray
        if ($button.Content -is [System.Windows.Controls.StackPanel]) {
            foreach ($c in $button.Content.Children) { if ($c -is [System.Windows.Controls.TextBlock]) { $c.Foreground = [System.Windows.Media.Brushes]::White } }
        }
    } else {
        $button.Background = [System.Windows.Media.Brushes]::White
        $button.Foreground = [System.Windows.Media.Brushes]::Black
        $button.BorderBrush = [System.Windows.Media.Brushes]::Gray
        if ($button.Content -is [System.Windows.Controls.StackPanel]) {
            foreach ($c in $button.Content.Children) { if ($c -is [System.Windows.Controls.TextBlock]) { $c.Foreground = [System.Windows.Media.Brushes]::Black } }
        }
    }
    if ($button.Tag.AlignRight -eq "true") { $rightStack.Children.Add($button) | Out-Null } else { $leftStack.Children.Add($button) | Out-Null }
    if (-not $NoSave) { Save-Shortcuts }
    Update-ShortcutStackMargins
}

# --- Save Shortcuts to INI ---
function Save-Shortcuts {
    $data = @{}
    $index = 0
    foreach ($stack in @($leftStack, $rightStack)) {
        $align = if ($stack -eq $leftStack) { "false" } else { "true" }
        foreach ($btn in $stack.Children) {
            if ($btn.Tag -eq "Indicator") { continue }
            $data["Shortcut$index"] = @{
                Path        = $btn.Tag.FilePath
                OpenAsAdmin = $btn.Tag.OpenAsAdmin
                ShowText    = $btn.Tag.ShowText
                AlignRight  = $align
                DisplayName = $btn.Tag.DisplayName
            }
            $index++
        }
    }
    Write-IniFile $shortcutsFile $data
}

# --- Update Shortcut Stack Margins ---
function Update-ShortcutStackMargins {
    foreach ($stack in @($leftStack, $rightStack)) {
        $count = $stack.Children.Count
        for ($i = 0; $i -lt $count; $i++) {
            $child = $stack.Children[$i]
            if ($child.Tag -eq "Indicator") { continue }
            if ($i -eq 0) { $child.Margin = if ($count -eq 1) { [System.Windows.Thickness]::new(0,0,0,0) } else { [System.Windows.Thickness]::new(0,0,1,0) } }
            elseif ($i -eq ($count - 1)) { $child.Margin = [System.Windows.Thickness]::new(1,0,0,0) }
            else { $child.Margin = [System.Windows.Thickness]::new(1,0,1,0) }
        }
    }
}

$leftStack.Add_DragOver({ if ($_.Data.GetDataPresent("ButtonSource")) { $_.Effects = [System.Windows.DragDropEffects]::Move; $_.Handled = $true } })
$leftStack.Add_Drop({
    param($s, $e)
    if ($e.Data.GetDataPresent("ButtonSource")) {
         $btn = $e.Data.GetData("ButtonSource")
         $btn.Tag.AlignRight = "false"
         if ($rightStack.Children.Contains($btn)) { [void]$rightStack.Children.Remove($btn) }
         $pos = $e.GetPosition($s); $index = 0
         foreach ($child in $s.Children) {
             if ($child -ne $global:InsertIndicator) {
                 $childPos = $child.TransformToAncestor($s).Transform([System.Windows.Point]::new(0,0))
                 if (($childPos.X + $child.ActualWidth/2) -gt $pos.X) { break }
                 $index++
             }
         }
         [void]$s.Children.Insert($index, $btn)
         Save-Shortcuts; Update-ShortcutStackMargins; $e.Handled = $true
    }
})
$rightStack.Add_DragOver({ if ($_.Data.GetDataPresent("ButtonSource")) { $_.Effects = [System.Windows.DragDropEffects]::Move; $_.Handled = $true } })
$rightStack.Add_Drop({
    param($s, $e)
    if ($e.Data.GetDataPresent("ButtonSource")) {
         $btn = $e.Data.GetData("ButtonSource")
         $btn.Tag.AlignRight = "true"
         if ($leftStack.Children.Contains($btn)) { [void]$leftStack.Children.Remove($btn) }
         $pos = $e.GetPosition($s); $index = 0
         foreach ($child in $s.Children) {
             if ($child -ne $global:InsertIndicator) {
                 $childPos = $child.TransformToAncestor($s).Transform([System.Windows.Point]::new(0,0))
                 if (($childPos.X + $child.ActualWidth/2) -gt $pos.X) { break }
                 $index++
             }
         }
         [void]$s.Children.Insert($index, $btn)
         Save-Shortcuts; Update-ShortcutStackMargins; $e.Handled = $true
    }
})
$leftStack.Add_DragLeave({ Remove-InsertionIndicator })
$rightStack.Add_DragLeave({ Remove-InsertionIndicator })

# --- Load Existing Shortcuts ---
if (Test-Path -LiteralPath $shortcutsFile) {
    $ini = Read-IniFile $shortcutsFile
    foreach ($section in $ini.Keys) {
        $entry = $ini[$section]
        if (Test-Path -LiteralPath $entry.Path) {
            Add-ShortcutButton -FilePath $entry.Path -NoSave `
                -DefOpenAsAdmin $entry.OpenAsAdmin -DefShowText $entry.ShowText -DefAlignRight $entry.AlignRight -DefDisplayName $entry.DisplayName
        } else { Write-Warning "File or folder '$($entry.Path)' no longer exists." }
    }
}

# --- Settings Button Menu ---
$settingsButton.Add_Click({
    $cm = New-Object System.Windows.Controls.ContextMenu
    $itemSettings = New-Object System.Windows.Controls.MenuItem; $itemSettings.Header = "Settings"
    $itemSettings.Add_Click({ Show-OptionsWindow })
    $itemClose = New-Object System.Windows.Controls.MenuItem; $itemClose.Header = "Close toolbar"
    $itemClose.Add_Click({
        # Restauration de la zone de travail par défaut (Bounds) du PrimaryScreen
        $baseWorkArea = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        [AppBar]::SystemParametersInfo([AppBar]::SPI_SETWORKAREA, 0, [ref](New-Object AppBar+Rect -Property @{ left = $baseWorkArea.Left; top = $baseWorkArea.Top; right = $baseWorkArea.Right; bottom = $baseWorkArea.Bottom }), [AppBar]::SPIF_UPDATEINIFILE)
        if ($global:AdminHelperProcess -and -not $global:AdminHelperProcess.HasExited) { $global:AdminHelperProcess.Kill() }
        $window.Close()
    })
    foreach ($item in @($itemSettings, $itemClose)) { $cm.Items.Add($item) }
    $cm.IsOpen = $true
})

# --- Options Window ---
function Show-OptionsWindow {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Options"; $form.Size = New-Object System.Drawing.Size(320,320); $form.StartPosition = "CenterScreen"
    $origToolbar = $global:ToolbarLocation; $origThickness = $global:ThicknessMode; $origTheme = $global:Theme
    $origNewShowText = $global:NewShortcutShowText; $origNewOpenAdmin = $global:NewShortcutOpenAsAdmin; $origNewAlignRight = $global:NewShortcutAlignRight
    $gbLoc = New-Object System.Windows.Forms.GroupBox; $gbLoc.Text = "Toolbar location"
    $gbLoc.Location = New-Object System.Drawing.Point(10,10); $gbLoc.Size = New-Object System.Drawing.Size(120,60)
    $radioTop = New-Object System.Windows.Forms.RadioButton; $radioTop.Text = "Top"; $radioTop.Location = New-Object System.Drawing.Point(10,20); $radioTop.AutoSize = $true
    $radioBottom = New-Object System.Windows.Forms.RadioButton; $radioBottom.Text = "Bottom"; $radioBottom.Location = New-Object System.Drawing.Point(10,40); $radioBottom.AutoSize = $true
    $gbLoc.Controls.AddRange(@($radioTop, $radioBottom)); if ($global:ToolbarLocation -eq "Top") { $radioTop.Checked = $true } else { $radioBottom.Checked = $true }
    $form.Controls.Add($gbLoc)
    $gbThick = New-Object System.Windows.Forms.GroupBox; $gbThick.Text = "Toolbar thickness"
    $gbThick.Location = New-Object System.Drawing.Point(150,10); $gbThick.Size = New-Object System.Drawing.Size(120,100)
    $rSmall = New-Object System.Windows.Forms.RadioButton; $rSmall.Text = "Small"; $rSmall.Location = New-Object System.Drawing.Point(10,20); $rSmall.AutoSize = $true
    $rMedium = New-Object System.Windows.Forms.RadioButton; $rMedium.Text = "Medium"; $rMedium.Location = New-Object System.Drawing.Point(10,45); $rMedium.AutoSize = $true
    $rLarge = New-Object System.Windows.Forms.RadioButton; $rLarge.Text = "Large"; $rLarge.Location = New-Object System.Drawing.Point(10,70); $rLarge.AutoSize = $true
    $gbThick.Controls.AddRange(@($rSmall, $rMedium, $rLarge))
    switch ($global:ThicknessMode) { "Small" { $rSmall.Checked = $true } "Medium" { $rMedium.Checked = $true } "Large" { $rLarge.Checked = $true } }
    $form.Controls.Add($gbThick)
    $gbTheme = New-Object System.Windows.Forms.GroupBox; $gbTheme.Text = "Theme"
    $gbTheme.Location = New-Object System.Drawing.Point(10,80); $gbTheme.Size = New-Object System.Drawing.Size(120,60)
    $rLight = New-Object System.Windows.Forms.RadioButton; $rLight.Text = "Light"; $rLight.Location = New-Object System.Drawing.Point(10,20); $rLight.AutoSize = $true
    $rDark = New-Object System.Windows.Forms.RadioButton; $rDark.Text = "Dark"; $rDark.Location = New-Object System.Drawing.Point(10,40); $rDark.AutoSize = $true
    $gbTheme.Controls.AddRange(@($rLight, $rDark)); if ($global:Theme -eq "Dark") { $rDark.Checked = $true } else { $rLight.Checked = $true }
    $form.Controls.Add($gbTheme)
    $lblNew = New-Object System.Windows.Forms.Label; $lblNew.Text = "New shortcuts:"; $lblNew.Location = New-Object System.Drawing.Point(150,120); $lblNew.AutoSize = $true
    $form.Controls.Add($lblNew)
    $cbShowText = New-Object System.Windows.Forms.CheckBox; $cbShowText.Text = "Show text"; $cbShowText.Location = New-Object System.Drawing.Point(150,140)
    $cbShowText.Checked = ($global:NewShortcutShowText -eq "true"); $form.Controls.Add($cbShowText)
    $cbOpenAdmin = New-Object System.Windows.Forms.CheckBox; $cbOpenAdmin.Text = "Open as admin"; $cbOpenAdmin.Location = New-Object System.Drawing.Point(150,165)
    $cbOpenAdmin.Checked = ($global:NewShortcutOpenAsAdmin -eq "true"); $form.Controls.Add($cbOpenAdmin)
    $cbAlignRight = New-Object System.Windows.Forms.CheckBox; $cbAlignRight.Text = "Align right"; $cbAlignRight.Location = New-Object System.Drawing.Point(150,190)
    $cbAlignRight.Checked = ($global:NewShortcutAlignRight -eq "true"); $form.Controls.Add($cbAlignRight)
    $btnImport = New-Object System.Windows.Forms.Button; $btnImport.Text = "Import shortcuts"; $btnImport.Location = New-Object System.Drawing.Point(10,150)
    $btnImport.Size = New-Object System.Drawing.Size(120,30); $btnImport.Add_Click({ Import-Shortcuts }); $form.Controls.Add($btnImport)
    $btnExport = New-Object System.Windows.Forms.Button; $btnExport.Text = "Export shortcuts"; $btnExport.Location = New-Object System.Drawing.Point(10,190)
    $btnExport.Size = New-Object System.Drawing.Size(120,30); $btnExport.Add_Click({ Export-Shortcuts }); $form.Controls.Add($btnExport)
    $okButton = New-Object System.Windows.Forms.Button; $okButton.Text = "OK"; $okButton.Location = New-Object System.Drawing.Point(200,230)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Controls.Add($okButton); $form.AcceptButton = $okButton
    $radioTop.Add_CheckedChanged({ if ($radioTop.Checked) { $global:ToolbarLocation = "Top"; Update-AppBarPosition -position "Top"; Update-ShortcutStackMargins } })
    $radioBottom.Add_CheckedChanged({ if ($radioBottom.Checked) { $global:ToolbarLocation = "Bottom"; Update-AppBarPosition -position "Bottom"; Update-ShortcutStackMargins } })
    $rSmall.Add_CheckedChanged({ if ($rSmall.Checked) { $global:ThicknessMode = "Small"; $global:BarThickness = $global:BaseThickness = 25; Update-AppBarPosition -position $global:ToolbarLocation; Update-ShortcutButtonsAppearance } })
    $rMedium.Add_CheckedChanged({ if ($rMedium.Checked) { $global:ThicknessMode = "Medium"; $global:BarThickness = $global:BaseThickness = 32; Update-AppBarPosition -position $global:ToolbarLocation; Update-ShortcutButtonsAppearance } })
    $rLarge.Add_CheckedChanged({ if ($rLarge.Checked) { $global:ThicknessMode = "Large"; $global:BarThickness = $global:BaseThickness = 40; Update-AppBarPosition -position $global:ToolbarLocation; Update-ShortcutButtonsAppearance } })
    $rLight.Add_CheckedChanged({ if ($rLight.Checked) { $global:Theme = "Light"; Update-ThemeAppearance } })
    $rDark.Add_CheckedChanged({ if ($rDark.Checked) { $global:Theme = "Dark"; Update-ThemeAppearance } })
    $cbShowText.Add_CheckedChanged({ if ($cbShowText.Checked) { $global:NewShortcutShowText = "true" } else { $global:NewShortcutShowText = "false" } })
    $cbOpenAdmin.Add_CheckedChanged({ if ($cbOpenAdmin.Checked) { $global:NewShortcutOpenAsAdmin = "true" } else { $global:NewShortcutOpenAsAdmin = "false" } })
    $cbAlignRight.Add_CheckedChanged({ if ($cbAlignRight.Checked) { $global:NewShortcutAlignRight = "true" } else { $global:NewShortcutAlignRight = "false" } })
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Save-Settings }
    else {
        $global:ToolbarLocation = $origToolbar; $global:ThicknessMode = $origThickness; $global:Theme = $origTheme
        $global:NewShortcutShowText = $origNewShowText; $global:NewShortcutOpenAsAdmin = $origNewOpenAdmin; $global:NewShortcutAlignRight = $origNewAlignRight
        switch ($global:ThicknessMode) { "Small" { $global:BarThickness = 25 } "Medium" { $global:BarThickness = 32 } "Large" { $global:BarThickness = 40 } default { $global:BarThickness = 25 } }
        Update-AppBarPosition -position $global:ToolbarLocation; Update-ShortcutButtonsAppearance; Update-ThemeAppearance
    }
    $form.Dispose()
}

# --- Import/Export Shortcuts ---
function Import-Shortcuts {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "INI Files|*.ini|All Files|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $importFile = $openFileDialog.FileName
        $choiceForm = New-Object System.Windows.Forms.Form
        $choiceForm.Text = "Import Options"; $choiceForm.Size = New-Object System.Drawing.Size(300,150); $choiceForm.StartPosition = "CenterScreen"
        $rAdd = New-Object System.Windows.Forms.RadioButton; $rAdd.Text = "Add"; $rAdd.Location = New-Object System.Drawing.Point(10,20); $rAdd.Checked = $true
        $rReplace = New-Object System.Windows.Forms.RadioButton; $rReplace.Text = "Replace"; $rReplace.Location = New-Object System.Drawing.Point(10,50)
        $okBtn = New-Object System.Windows.Forms.Button; $okBtn.Text = "OK"; $okBtn.Location = New-Object System.Drawing.Point(200,70); $okBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $choiceForm.Controls.AddRange(@($rAdd, $rReplace, $okBtn)); $choiceForm.AcceptButton = $okBtn
        if ($choiceForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $choice = if ($rReplace.Checked) { "Replace" } else { "Add" }
            $importData = Read-IniFile $importFile
            if ($choice -eq "Replace") {
                Write-IniFile $shortcutsFile $importData
                $leftStack.Children.Clear(); $rightStack.Children.Clear()
                foreach ($section in $importData.Keys) {
                    $entry = $importData[$section]
                    if (Test-Path -LiteralPath $entry.Path) {
                        Add-ShortcutButton -FilePath $entry.Path -NoSave `
                            -DefOpenAsAdmin $entry.OpenAsAdmin -DefShowText $entry.ShowText -DefAlignRight $entry.AlignRight -DefDisplayName $entry.DisplayName
                    }
                }
            } else {
                $existing = @{}
                if (Test-Path -LiteralPath $shortcutsFile) { $existing = Read-IniFile $shortcutsFile }
                $index = $existing.Keys.Count
                foreach ($section in $importData.Keys) { $existing["Shortcut$index"] = $importData[$section]; $index++ }
                Write-IniFile $shortcutsFile $existing
                foreach ($section in $importData.Keys) {
                    $entry = $importData[$section]
                    if (Test-Path -LiteralPath $entry.Path) {
                        Add-ShortcutButton -FilePath $entry.Path -NoSave `
                            -DefOpenAsAdmin $entry.OpenAsAdmin -DefShowText $entry.ShowText -DefAlignRight $entry.AlignRight -DefDisplayName $entry.DisplayName
                    }
                }
            }
            Save-Shortcuts; Update-ShortcutStackMargins
        }
        $choiceForm.Dispose()
    }
}
function Export-Shortcuts {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "INI Files|*.ini|All Files|*.*"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Copy-Item -Path $shortcutsFile -Destination $saveFileDialog.FileName -Force }
}

# --- Drag-and-Drop on Main Window ---
$window.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) { $_.Effects = [System.Windows.DragDropEffects]::Copy }
    else { $_.Effects = [System.Windows.DragDropEffects]::None }
    $_.Handled = $true
})
$window.Add_Drop({
    $files = $_.Data.GetData([Windows.DataFormats]::FileDrop)
    foreach ($f in $files) { Add-ShortcutButton $f }
    $_.Handled = $true
})

$window.Add_Loaded({
    Update-AppBarPosition -position $global:ToolbarLocation
    Update-ThemeAppearance
    $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle
    if ($hwnd -eq [IntPtr]::Zero) { 
        Write-Error "Invalid window handle."
        return 
    }
    $handler = { param($msg, $wParam, $lParam) }
    try { 
        [MessageHook]::AddHook($hwnd, $global:CallbackMessage, $handler) 
    }
    catch { 
        Write-Error "Error adding message hook: $_" 
    }
    [Microsoft.Win32.SystemEvents]::add_DisplaySettingsChanged({
        Start-Sleep -Milliseconds 500  # Attendre que Windows termine ses modifications
        Refresh-WorkArea
        Update-ShortcutButtonsAppearance
    })
})

# --- Show Main Window ---
$window.ShowDialog() | Out-Null
