Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# DPI
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DpiHelper {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
}
"@ -ErrorAction SilentlyContinue
    [void][DpiHelper]::SetProcessDPIAware()
} catch { }

# Taskbar
try {
    Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
public class TBProg {
    [ComImport][Guid("ea1afb91-9e28-4b86-90e9-9e9f8a5eefaf")][InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITB3 { void HrInit(); void AddTab(IntPtr h); void DeleteTab(IntPtr h); void ActivateTab(IntPtr h); void SetActiveAlt(IntPtr h); void MarkFullscreenWindow(IntPtr h, bool f); void SetProgressValue(IntPtr h, UInt64 c, UInt64 t); void SetProgressState(IntPtr h, int s); }
    [ComImport][Guid("56fdf344-fd6d-11d0-958a-006097c9a090")][ClassInterface(ClassInterfaceType.None)] public class TBI { }
    static ITB3 _tb; public static ITB3 TB { get { if(_tb==null)_tb=(ITB3)new TBI(); return _tb; } }
    public static void SetState(IntPtr h,int s){try{TB.SetProgressState(h,s);}catch{}}
    public static void SetVal(IntPtr h,double v,double m){try{TB.SetProgressValue(h,(ulong)v,(ulong)m);}catch{}}
}
"@ -ErrorAction SilentlyContinue
} catch { }

# Custom Progress Bar Control with percentage text
Add-Type -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

public class ModernProgressBar : Control {
    private int val = 0;
    private int max = 100;
    private Color barColor = Color.FromArgb(200, 160, 120);
    private Color bgColor = Color.FromArgb(240, 230, 220);
    private Color textColor = Color.FromArgb(100, 70, 50);
    private bool isRTL = true;
    private bool showText = true;
    
    public int Value { 
        get { return val; } 
        set { val = Math.Min(Math.Max(value, 0), max); Invalidate(); } 
    }
    public int Maximum { 
        get { return max; } 
        set { max = value; Invalidate(); } 
    }
    public Color BarColor { 
        get { return barColor; } 
        set { barColor = value; Invalidate(); } 
    }
    public Color BgColor { 
        get { return bgColor; } 
        set { bgColor = value; Invalidate(); } 
    }
    public Color TextColor { 
        get { return textColor; } 
        set { textColor = value; Invalidate(); } 
    }
    public bool IsRTL { 
        get { return isRTL; } 
        set { isRTL = value; Invalidate(); } 
    }
    public bool ShowText { 
        get { return showText; } 
        set { showText = value; Invalidate(); } 
    }
    
    public ModernProgressBar() {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        Height = 24;
        Font = new Font("Segoe UI", 9, FontStyle.Bold);
    }
    
    protected override void OnPaint(PaintEventArgs e) {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
        
        // Background with rounded corners
        using (GraphicsPath bgPath = RoundedRect(new Rectangle(0, 0, Width, Height), 8)) {
            using (SolidBrush bgBrush = new SolidBrush(bgColor)) {
                e.Graphics.FillPath(bgBrush, bgPath);
            }
        }
        
        // Progress bar
        if (val > 0 && max > 0) {
            int progressWidth = (int)((float)val / max * Width);
            if (progressWidth > 0) {
                Rectangle progressRect;
                if (isRTL) {
                    progressRect = new Rectangle(Width - progressWidth, 0, progressWidth, Height);
                } else {
                    progressRect = new Rectangle(0, 0, progressWidth, Height);
                }
                using (GraphicsPath progressPath = RoundedRect(progressRect, 8)) {
                    using (SolidBrush progressBrush = new SolidBrush(barColor)) {
                        e.Graphics.FillPath(progressBrush, progressPath);
                    }
                }
            }
        }
        
        // Draw percentage text
        if (showText && val > 0) {
            string text = val.ToString() + "%";
            using (SolidBrush textBrush = new SolidBrush(textColor)) {
                StringFormat sf = new StringFormat();
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                e.Graphics.DrawString(text, Font, textBrush, new RectangleF(0, 0, Width, Height), sf);
            }
        }
    }
    
    private GraphicsPath RoundedRect(Rectangle rect, int radius) {
        GraphicsPath path = new GraphicsPath();
        int d = radius * 2;
        path.AddArc(rect.X, rect.Y, d, d, 180, 90);
        path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
        path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }
}
"@ -ReferencedAssemblies System.Windows.Forms,System.Drawing -ErrorAction SilentlyContinue

# Global Variables
$script:DarkMode = $false
$script:StableRelease = $null
$script:PreRelease = $null
$script:StableChangelog = ""
$script:PreChangelog = ""
$script:SelectedRelease = $null
$script:TempFile = ""
$script:FinalFile = ""
$script:TotalSize = 0
$script:IsDownloading = $false
$script:IsPaused = $false
$script:DownloadCompleted = $false
$script:DownloadJob = $null
$script:SilentInstall = $false
$script:DownloadedRelease = $null

$script:LibTempFile = ""
$script:LibFinalFile = ""
$script:LibFileToExtract = $null
$script:LibTotalSize = 0
$script:LibIsDownloading = $false
$script:LibIsPaused = $false
$script:LibDownloadCompleted = $false
$script:LibDownloadJob = $null
$script:LibUrl = "https://github.com/Y-PLONI/otzaria-library/releases/latest/download/otzaria_latest.zip"
$script:LibVer = ""

$script:InstallPath = $null
$script:InstalledVersion = "לא נמצא"
$script:InstalledLibVersion = "לא נמצא"
$script:FullRelease = $null
$script:InstallType = "EXE"  # EXE or MSIX
$script:MsixInstalled = $false
$script:MsixVersion = $null
$script:StableMsixRelease = $null
$script:PreMsixRelease = $null

# Color Themes
$script:LightTheme = @{
    BgColor = [System.Drawing.Color]::FromArgb(255,248,244)
    HeaderColor = [System.Drawing.Color]::FromArgb(249,236,223)
    CardColor = [System.Drawing.Color]::White
    TextColor = [System.Drawing.Color]::FromArgb(100,70,50)
    TextLight = [System.Drawing.Color]::FromArgb(150,120,100)
    Divider = [System.Drawing.Color]::FromArgb(240,230,220)
    BtnBg = [System.Drawing.Color]::FromArgb(255,241,229)
    BtnBorder = [System.Drawing.Color]::FromArgb(235,215,195)
    ProgBg = [System.Drawing.Color]::FromArgb(240,230,220)
    ProgBar = [System.Drawing.Color]::FromArgb(200,160,120)
}

$script:DarkTheme = @{
    BgColor = [System.Drawing.Color]::FromArgb(30,30,35)
    HeaderColor = [System.Drawing.Color]::FromArgb(40,40,48)
    CardColor = [System.Drawing.Color]::FromArgb(45,45,52)
    TextColor = [System.Drawing.Color]::FromArgb(230,225,220)
    TextLight = [System.Drawing.Color]::FromArgb(160,155,150)
    Divider = [System.Drawing.Color]::FromArgb(60,60,68)
    BtnBg = [System.Drawing.Color]::FromArgb(55,55,65)
    BtnBorder = [System.Drawing.Color]::FromArgb(80,80,95)
    ProgBg = [System.Drawing.Color]::FromArgb(35,35,40)
    ProgBar = [System.Drawing.Color]::FromArgb(90,90,100)
}

function Get-Theme { if ($script:DarkMode) { $script:DarkTheme } else { $script:LightTheme } }

# RTL MessageBox function
function Show-RTLMessageBox {
    param(
        [string]$Message,
        [string]$Title,
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    $rtlOption = [System.Windows.Forms.MessageBoxOptions]::RtlReading -bor [System.Windows.Forms.MessageBoxOptions]::RightAlign
    return [System.Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons, $Icon, [System.Windows.Forms.MessageBoxDefaultButton]::Button1, $rtlOption)
}

# Functions
function Find-OtzariaInstallPath {
    # Try Registry first (most reliable)
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($regPath in $regPaths) {
        try {
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*otzaria*" -or $_.DisplayName -like "*אוצריא*" }
            foreach ($app in $apps) {
                $installLoc = $app.InstallLocation
                if ($installLoc -and (Test-Path (Join-Path $installLoc "otzaria.exe"))) {
                    return $installLoc.TrimEnd('\')
                }
                # Try UninstallString to extract path
                $uninstall = $app.UninstallString
                if ($uninstall -and $uninstall -match '^"?([^"]+)\\[^\\]+\.exe') {
                    $dir = $matches[1]
                    if (Test-Path (Join-Path $dir "otzaria.exe")) {
                        return $dir
                    }
                }
            }
        } catch { }
    }
    
    # Fallback: scan drives for אוצריא folder
    foreach ($d in Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match '^[A-Z]:\\$' }) {
        $exePath = Join-Path $d.Root "אוצריא\otzaria.exe"
        if (Test-Path $exePath) { return Join-Path $d.Root "אוצריא" }
    }
    
    return $null
}

function Get-MsixInstallation {
    try {
        $msixApp = Get-AppxPackage | Where-Object { $_.Name -like "*otzaria*" -or $_.PackageFullName -like "*sivan22.Otzaria*" } | Select-Object -First 1
        if ($msixApp) {
            return @{
                Installed = $true
                Version = $msixApp.Version
                InstallLocation = $msixApp.InstallLocation
                PackageFullName = $msixApp.PackageFullName
            }
        }
    } catch { }
    return @{ Installed = $false; Version = $null; InstallLocation = $null; PackageFullName = $null }
}

function Get-InstalledVersion {
    param($installPath)
    if ($installPath) {
        $exePath = Join-Path $installPath "otzaria.exe"
        if (Test-Path $exePath) {
            try { $ver = (Get-Item $exePath).VersionInfo.FileVersion; if ($ver) { return $ver } } catch { }
        }
    }
    return "לא נמצא"
}

function Get-InstalledLibVersion {
    param($installPath)
    if ($installPath) {
        $libVerPath = Join-Path $installPath "אוצריא\אודות התוכנה\גירסת ספריה.txt"
        if (Test-Path $libVerPath) {
            try { $ver = Get-Content $libVerPath -Encoding UTF8 -ErrorAction SilentlyContinue | Select-Object -First 1; if ($ver) { return $ver.Trim() } } catch { }
        }
    }
    return "לא נמצא"
}

function Get-Release([bool]$Pre) {
    try {
        $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
        $rel = $r | Where-Object { $_.prerelease -eq $Pre } | Select-Object -First 1
        if (-not $rel) { return @{Release=$null; Changelog=""} }
        $a = $rel.assets | Where-Object { $_.name -like "otzaria-*-windows.exe" -and $_.name -notlike "*-full.exe" } | Select-Object -First 1
        if (-not $a) { return @{Release=$null; Changelog=""} }
        $changelog = if ($rel.body) { $rel.body } else { "אין מידע על שינויים" }
        if ($rel.tag_name -match '^([\d\.]+)\+(\d+)$') {
            return @{
                Release = [pscustomobject]@{ FullVersion = "$($matches[1]).$($matches[2])"; File = "otzaria-$($matches[1]).$($matches[2])-windows.exe"; Url = $a.browser_download_url }
                Changelog = $changelog
            }
        }
    } catch { }
    return @{Release=$null; Changelog=""}
}

function Get-FullRelease {
    try {
        $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
        foreach ($rel in $r) {
            $a = $rel.assets | Where-Object { $_.name -like "otzaria-*-windows-full.exe" } | Select-Object -First 1
            if ($a) {
                # Get version from tag (includes build number like 0.9.74+425)
                if ($rel.tag_name -match '^([\d\.]+)\+(\d+)$') {
                    $fullVer = "$($matches[1]).$($matches[2])"
                    return [pscustomobject]@{ 
                        FullVersion = $fullVer
                        File = "otzaria-$fullVer-windows-full.exe"
                        Url = $a.browser_download_url 
                    }
                }
                # Fallback to filename version if tag doesn't match
                elseif ($a.name -match 'otzaria-([\d\.]+)-windows-full\.exe') {
                    return [pscustomobject]@{ 
                        FullVersion = $matches[1]
                        File = $a.name
                        Url = $a.browser_download_url 
                    }
                }
            }
        }
    } catch { }
    return $null
}

function Get-MsixRelease([bool]$Pre) {
    try {
        $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
        $rel = $r | Where-Object { $_.prerelease -eq $Pre } | Select-Object -First 1
        if (-not $rel) { return $null }
        $a = $rel.assets | Where-Object { $_.name -like "*.msix" } | Select-Object -First 1
        if (-not $a) { return $null }
        if ($rel.tag_name -match '^([\d\.]+)\+(\d+)$') {
            $fullVer = "$($matches[1]).$($matches[2])"
            return [pscustomobject]@{ 
                FullVersion = $fullVer
                File = "otzaria-$fullVer.msix"
                Url = $a.browser_download_url 
            }
        }
    } catch { }
    return $null
}

function Get-LibVer {
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Encoding = [System.Text.Encoding]::UTF8
        $wc.Headers.Add("User-Agent","PS")
        return $wc.DownloadString("https://github.com/Otzaria/otzaria-library/raw/refs/heads/main/MoreBooks/%D7%A1%D7%A4%D7%A8%D7%99%D7%9D/%D7%90%D7%95%D7%A6%D7%A8%D7%99%D7%90/%D7%90%D7%95%D7%93%D7%95%D7%AA%20%D7%94%D7%AA%D7%95%D7%9B%D7%A0%D7%94/%D7%92%D7%99%D7%A8%D7%A1%D7%AA%20%D7%A1%D7%A4%D7%A8%D7%99%D7%94.txt").Trim()
    } catch { return "?" }
}

function Get-CurVer {
    $f = Get-ChildItem "." -Filter "otzaria-*-windows.exe" -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
    if ($f -and $f.Name -match 'otzaria-([\d\.]+)-windows\.exe') { return $matches[1] }
    return "לא נמצא"
}

function Find-Otzaria {
    # Use the main function that checks registry first
    return Find-OtzariaInstallPath
}

function Get-Size($u) {
    try {
        $r = [System.Net.WebRequest]::Create($u)
        $r.Method = "HEAD"; $r.Timeout = 10000
        $resp = $r.GetResponse(); $s = $resp.ContentLength; $resp.Close()
        return $s
    } catch { return 0 }
}

function Get-ZipLibVersion {
    param($zipPath)
    try {
        if (-not (Test-Path $zipPath)) { return $null }
        
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        
        $encoding = [System.Text.Encoding]::GetEncoding(862)
        $zipArchive = [System.IO.Compression.ZipFile]::Open($zipPath, 'Read', $encoding)
        
        # Look for the version file inside the ZIP
        $versionEntry = $zipArchive.Entries | Where-Object { 
            $_.FullName -like "*אודות התוכנה/גירסת ספריה.txt" -or 
            $_.FullName -like "*אודות התוכנה\גירסת ספריה.txt"
        } | Select-Object -First 1
        
        if ($versionEntry) {
            $stream = $versionEntry.Open()
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            $version = $reader.ReadLine()
            $reader.Close()
            $stream.Close()
            $zipArchive.Dispose()
            return $version.Trim()
        }
        
        $zipArchive.Dispose()
        return $null
    } catch {
        return $null
    }
}

function Get-LibChangelog {
    try {
        $url = "https://raw.githubusercontent.com/Otzaria/otzaria-library/main/MoreBooks/%D7%A1%D7%A4%D7%A8%D7%99%D7%9D/%D7%90%D7%95%D7%A6%D7%A8%D7%99%D7%90/%D7%90%D7%95%D7%93%D7%95%D7%AA%20%D7%94%D7%AA%D7%95%D7%9B%D7%A0%D7%94/%D7%A2%D7%93%D7%9B%D7%95%D7%A0%D7%99%20%D7%A1%D7%A4%D7%A8%D7%99%D7%94.md"
        $wc = New-Object System.Net.WebClient
        $wc.Encoding = [System.Text.Encoding]::UTF8
        return $wc.DownloadString($url)
    } catch {
        return "לא ניתן לטעון את רשימת העדכונים"
    }
}

function Get-BetaChangelog {
    try {
        $url = "https://raw.githubusercontent.com/Otzaria/otzaria/dev/assets/%D7%99%D7%95%D7%9E%D7%9F%20%D7%A9%D7%99%D7%A0%D7%95%D7%99%D7%99%D7%9D.md"
        $wc = New-Object System.Net.WebClient
        $wc.Encoding = [System.Text.Encoding]::UTF8
        return $wc.DownloadString($url)
    } catch {
        return "לא ניתן לטעון את רשימת העדכונים"
    }
}

function Convert-MarkdownToText {
    param($md)
    if (-not $md) { return "" }
    $text = $md
    # Remove headers markers but keep text
    $text = $text -replace '(?m)^#{1,6}\s*', ''
    # Convert bold **text** or __text__ to text
    $text = $text -replace '\*\*([^\*]+)\*\*', '$1'
    $text = $text -replace '__([^_]+)__', '$1'
    # Convert italic *text* or _text_ to text
    $text = $text -replace '(?<!\*)\*([^\*]+)\*(?!\*)', '$1'
    $text = $text -replace '(?<!_)_([^_]+)_(?!_)', '$1'
    # Convert list items - to •
    $text = $text -replace '(?m)^\s*[-\*]\s+', '• '
    # Remove code blocks markers
    $text = $text -replace '```[^\n]*\n?', ''
    # Remove inline code markers
    $text = $text -replace '`([^`]+)`', '$1'
    # Remove link markdown [text](url) -> text
    $text = $text -replace '\[([^\]]+)\]\([^\)]+\)', '$1'
    # Remove extra blank lines
    $text = $text -replace '(\r?\n){3,}', "`n`n"
    return $text.Trim()
}

function Show-Changelog {
    param($title, $changelog)
    $t = Get-Theme
    $clForm = New-Object System.Windows.Forms.Form
    $clForm.Text = $title
    $clForm.Size = New-Object System.Drawing.Size(550,500)
    $clForm.StartPosition = "CenterParent"
    $clForm.RightToLeft = "Yes"
    $clForm.RightToLeftLayout = $true
    $clForm.BackColor = $t.BgColor
    $clForm.FormBorderStyle = "FixedDialog"
    $clForm.MaximizeBox = $false
    $clForm.MinimizeBox = $false
    
    # Convert Markdown to readable text
    $displayText = Convert-MarkdownToText $changelog
    
    # Panel with scroll for content
    $scrollPanel = New-Object System.Windows.Forms.Panel
    $scrollPanel.Location = New-Object System.Drawing.Point(20,20)
    $scrollPanel.Size = New-Object System.Drawing.Size(495,370)
    $scrollPanel.BackColor = $t.CardColor
    $scrollPanel.AutoScroll = $true
    $clForm.Controls.Add($scrollPanel)
    
    # Label for text display (not editable)
    $lblContent = New-Object System.Windows.Forms.Label
    $lblContent.Text = $displayText
    $lblContent.Location = New-Object System.Drawing.Point(10,10)
    $lblContent.MaximumSize = New-Object System.Drawing.Size(460,0)
    $lblContent.AutoSize = $true
    $lblContent.ForeColor = $t.TextColor
    $lblContent.Font = New-Object System.Drawing.Font("Segoe UI",11)
    $scrollPanel.Controls.Add($lblContent)
    
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "סגור"
    $btnClose.Location = New-Object System.Drawing.Point(220,410)
    $btnClose.Size = New-Object System.Drawing.Size(110,40)
    $btnClose.BackColor = $t.BtnBg
    $btnClose.ForeColor = $t.TextColor
    $btnClose.FlatStyle = "Flat"
    $btnClose.Font = New-Object System.Drawing.Font("Segoe UI",11)
    $btnClose.Add_Click({ $clForm.Close() })
    $clForm.Controls.Add($btnClose)
    
    [void]$clForm.ShowDialog()
}

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "אוצריא - עדכון תוכנה"
$form.Size = New-Object System.Drawing.Size(720,860)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false
$form.RightToLeft = "Yes"
$form.RightToLeftLayout = $true
$form.FormBorderStyle = "FixedDialog"
$form.Font = New-Object System.Drawing.Font("Segoe UI",10)
$form.AutoScroll = $true
$form.AutoScrollMinSize = New-Object System.Drawing.Size(700,820)

# Load icon from current executable
try {
    $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath)
} catch { }

# Header
$header = New-Object System.Windows.Forms.Panel
$header.Dock = "Top"
$header.Height = 80

$title = New-Object System.Windows.Forms.Label
$title.Text = "אוצריא"
$title.Dock = "Top"
$title.Height = 48
$title.Font = New-Object System.Drawing.Font("Segoe UI",22,[System.Drawing.FontStyle]::Bold)
$title.TextAlign = "MiddleCenter"
$header.Controls.Add($title)

$sub = New-Object System.Windows.Forms.Label
$sub.Text = "עדכון ושדרוג התוכנה"
$sub.Dock = "Bottom"
$sub.Height = 28
$sub.Font = New-Object System.Drawing.Font("Segoe UI",11)
$sub.TextAlign = "MiddleCenter"
$header.Controls.Add($sub)

# Refresh Button (top-left corner, but in RTL it appears on left)
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "⟳"
$btnRefresh.Location = New-Object System.Drawing.Point(637,10)
$btnRefresh.Size = New-Object System.Drawing.Size(42,42)
$btnRefresh.FlatStyle = "Flat"
$btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI",16)
$btnRefresh.Cursor = "Hand"
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($btnRefresh, "רענן נתונים")
$form.Controls.Add($btnRefresh)

# Main Card - includes progress bar
$mainCard = New-Object System.Windows.Forms.Panel
$mainCard.Location = New-Object System.Drawing.Point(25,95)
$mainCard.Size = New-Object System.Drawing.Size(655,400)

# Title - centered above everything
$lblSec = New-Object System.Windows.Forms.Label
$lblSec.Text = "הורדת והתקנת אוצריא"
$lblSec.Location = New-Object System.Drawing.Point(20,12)
$lblSec.Size = New-Object System.Drawing.Size(610,28)
$lblSec.Font = New-Object System.Drawing.Font("Segoe UI Semibold",13)
$lblSec.TextAlign = "MiddleCenter"
$mainCard.Controls.Add($lblSec)

$div1 = New-Object System.Windows.Forms.Label
$div1.Location = New-Object System.Drawing.Point(20,45)
$div1.Size = New-Object System.Drawing.Size(610,1)
$mainCard.Controls.Add($div1)

# Current version row
$lblCurT = New-Object System.Windows.Forms.Label
$lblCurT.Text = "הגירסה בתיקייה:"
$lblCurT.Location = New-Object System.Drawing.Point(450,52)
$lblCurT.Size = New-Object System.Drawing.Size(180,24)
$lblCurT.TextAlign = "MiddleRight"
$lblCurT.Font = New-Object System.Drawing.Font("Segoe UI",10)
$mainCard.Controls.Add($lblCurT)

$lblCurV = New-Object System.Windows.Forms.Label
$lblCurV.Text = (Get-CurVer)
$lblCurV.Location = New-Object System.Drawing.Point(220,52)
$lblCurV.Size = New-Object System.Drawing.Size(220,24)
$lblCurV.Font = New-Object System.Drawing.Font("Segoe UI Semibold",10)
$lblCurV.TextAlign = "MiddleLeft"
$mainCard.Controls.Add($lblCurV)

# Install Type Selection - right side, below current version
$lblInstallType = New-Object System.Windows.Forms.Label
$lblInstallType.Text = "סוג:"
$lblInstallType.Location = New-Object System.Drawing.Point(560,82)
$lblInstallType.Size = New-Object System.Drawing.Size(70,24)
$lblInstallType.TextAlign = "MiddleRight"
$lblInstallType.Font = New-Object System.Drawing.Font("Segoe UI",10)
$mainCard.Controls.Add($lblInstallType)

$cmbInstallType = New-Object System.Windows.Forms.ComboBox
$cmbInstallType.Location = New-Object System.Drawing.Point(470,80)
$cmbInstallType.Size = New-Object System.Drawing.Size(80,26)
$cmbInstallType.Font = New-Object System.Drawing.Font("Segoe UI",9)
$cmbInstallType.DropDownStyle = "DropDownList"
$cmbInstallType.Items.AddRange(@("EXE", "MSIX"))
$cmbInstallType.SelectedIndex = 0
$mainCard.Controls.Add($cmbInstallType)

# Stable version row
$radioStable = New-Object System.Windows.Forms.RadioButton
$radioStable.Text = "גרסה יציבה"
$radioStable.Location = New-Object System.Drawing.Point(480,112)
$radioStable.Size = New-Object System.Drawing.Size(150,26)
$radioStable.Checked = $true
$radioStable.Font = New-Object System.Drawing.Font("Segoe UI",10)
$mainCard.Controls.Add($radioStable)

$lblStableV = New-Object System.Windows.Forms.Label
$lblStableV.Text = "טוען..."
$lblStableV.Location = New-Object System.Drawing.Point(280,112)
$lblStableV.Size = New-Object System.Drawing.Size(180,26)
$lblStableV.Font = New-Object System.Drawing.Font("Segoe UI Semibold",11)
$lblStableV.TextAlign = "MiddleCenter"
$mainCard.Controls.Add($lblStableV)

# Beta version row
$radioPre = New-Object System.Windows.Forms.RadioButton
$radioPre.Text = "גרסת בטא"
$radioPre.Location = New-Object System.Drawing.Point(480,144)
$radioPre.Size = New-Object System.Drawing.Size(150,26)
$radioPre.Font = New-Object System.Drawing.Font("Segoe UI",10)
$mainCard.Controls.Add($radioPre)

$lblPreV = New-Object System.Windows.Forms.Label
$lblPreV.Text = "טוען..."
$lblPreV.Location = New-Object System.Drawing.Point(280,144)
$lblPreV.Size = New-Object System.Drawing.Size(180,26)
$lblPreV.Font = New-Object System.Drawing.Font("Segoe UI Semibold",11)
$lblPreV.TextAlign = "MiddleCenter"
$mainCard.Controls.Add($lblPreV)

$btnPreChangelog = New-Object System.Windows.Forms.Button
$btnPreChangelog.Text = "מה חדש?"
$btnPreChangelog.Location = New-Object System.Drawing.Point(20,142)
$btnPreChangelog.Size = New-Object System.Drawing.Size(100,28)
$btnPreChangelog.FlatStyle = "Flat"
$btnPreChangelog.Font = New-Object System.Drawing.Font("Segoe UI",9)
$mainCard.Controls.Add($btnPreChangelog)

# Full version row (only for EXE)
$radioFull = New-Object System.Windows.Forms.RadioButton
$radioFull.Text = "גרסה מלאה"
$radioFull.Location = New-Object System.Drawing.Point(480,176)
$radioFull.Size = New-Object System.Drawing.Size(150,26)
$radioFull.Font = New-Object System.Drawing.Font("Segoe UI",10)
$mainCard.Controls.Add($radioFull)

$lblFullV = New-Object System.Windows.Forms.Label
$lblFullV.Text = "טוען..."
$lblFullV.Location = New-Object System.Drawing.Point(280,176)
$lblFullV.Size = New-Object System.Drawing.Size(180,26)
$lblFullV.Font = New-Object System.Drawing.Font("Segoe UI Semibold",11)
$lblFullV.TextAlign = "MiddleCenter"
$mainCard.Controls.Add($lblFullV)

# Downloaded version row (for existing files in folder)
$radioDownloaded = New-Object System.Windows.Forms.RadioButton
$radioDownloaded.Text = "זמין להתקנה"
$radioDownloaded.Location = New-Object System.Drawing.Point(480,208)
$radioDownloaded.Size = New-Object System.Drawing.Size(150,26)
$radioDownloaded.Font = New-Object System.Drawing.Font("Segoe UI",10)
$radioDownloaded.Visible = $false
$mainCard.Controls.Add($radioDownloaded)

$lblDownloadedV = New-Object System.Windows.Forms.Label
$lblDownloadedV.Text = ""
$lblDownloadedV.Location = New-Object System.Drawing.Point(280,208)
$lblDownloadedV.Size = New-Object System.Drawing.Size(180,26)
$lblDownloadedV.Font = New-Object System.Drawing.Font("Segoe UI Semibold",11)
$lblDownloadedV.TextAlign = "MiddleCenter"
$lblDownloadedV.Visible = $false
$mainCard.Controls.Add($lblDownloadedV)

$chkSilent = New-Object System.Windows.Forms.CheckBox
$chkSilent.Text = "התקנה שקטה"
$chkSilent.Location = New-Object System.Drawing.Point(240,244)
$chkSilent.Size = New-Object System.Drawing.Size(180,24)
$chkSilent.Font = New-Object System.Drawing.Font("Segoe UI",10)
$chkSilent.TextAlign = "MiddleCenter"
$chkSilent.Visible = $false
$chkSilent.Add_CheckedChanged({ $script:SilentInstall = $chkSilent.Checked })
$mainCard.Controls.Add($chkSilent)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "ביטול"
$btnCancel.Location = New-Object System.Drawing.Point(20,280)
$btnCancel.Size = New-Object System.Drawing.Size(100,36)
$btnCancel.FlatStyle = "Flat"
$btnCancel.Visible = $false
$mainCard.Controls.Add($btnCancel)

$btnDL = New-Object System.Windows.Forms.Button
$btnDL.Text = "הורדה"
$btnDL.Location = New-Object System.Drawing.Point(252,280)
$btnDL.Size = New-Object System.Drawing.Size(150,36)
$btnDL.FlatStyle = "Flat"
$btnDL.Enabled = $false
$mainCard.Controls.Add($btnDL)

# Progress bar inside main card
$progBar = New-Object ModernProgressBar
$progBar.Location = New-Object System.Drawing.Point(20,330)
$progBar.Size = New-Object System.Drawing.Size(610,22)
$progBar.IsRTL = $true
$mainCard.Controls.Add($progBar)

$statusLbl = New-Object System.Windows.Forms.Label
$statusLbl.Text = "טוען..."
$statusLbl.Location = New-Object System.Drawing.Point(20,357)
$statusLbl.Size = New-Object System.Drawing.Size(610,24)
$statusLbl.TextAlign = "MiddleCenter"
$statusLbl.Font = New-Object System.Drawing.Font("Segoe UI",10)
$mainCard.Controls.Add($statusLbl)

# Button to select install file (offline mode)
$btnSelectFile = New-Object System.Windows.Forms.Button
$btnSelectFile.Text = "בחר קובץ להתקנה"
$btnSelectFile.Location = New-Object System.Drawing.Point(227,280)
$btnSelectFile.Size = New-Object System.Drawing.Size(200,36)
$btnSelectFile.FlatStyle = "Flat"
$btnSelectFile.Visible = $false
$mainCard.Controls.Add($btnSelectFile)

# Size info label (shows download size inline)
$lblSizeInfo = New-Object System.Windows.Forms.Label
$lblSizeInfo.Text = ""
$lblSizeInfo.Location = New-Object System.Drawing.Point(20,352)
$lblSizeInfo.Size = New-Object System.Drawing.Size(610,20)
$lblSizeInfo.TextAlign = "MiddleCenter"
$lblSizeInfo.Font = New-Object System.Drawing.Font("Segoe UI",9)
$lblSizeInfo.Visible = $false
$mainCard.Controls.Add($lblSizeInfo)

# Library Card
$libCard = New-Object System.Windows.Forms.Panel
$libCard.Location = New-Object System.Drawing.Point(25,505)
$libCard.Size = New-Object System.Drawing.Size(655,165)

$lblLib = New-Object System.Windows.Forms.Label
$lblLib.Text = "ספריית אוצריא"
$lblLib.Location = New-Object System.Drawing.Point(20,12)
$lblLib.Size = New-Object System.Drawing.Size(610,28)
$lblLib.Font = New-Object System.Drawing.Font("Segoe UI Semibold",13)
$lblLib.TextAlign = "MiddleCenter"
$libCard.Controls.Add($lblLib)

$lblLibV = New-Object System.Windows.Forms.Label
$lblLibV.Text = "גרסה: טוען..."
$lblLibV.Location = New-Object System.Drawing.Point(120,42)
$lblLibV.Size = New-Object System.Drawing.Size(400,22)
$lblLibV.TextAlign = "MiddleCenter"
$lblLibV.Font = New-Object System.Drawing.Font("Segoe UI",10)
$libCard.Controls.Add($lblLibV)

$btnLibChangelog = New-Object System.Windows.Forms.Button
$btnLibChangelog.Text = "מה חדש?"
$btnLibChangelog.Location = New-Object System.Drawing.Point(20,40)
$btnLibChangelog.Size = New-Object System.Drawing.Size(100,28)
$btnLibChangelog.FlatStyle = "Flat"
$btnLibChangelog.Font = New-Object System.Drawing.Font("Segoe UI",9)
$btnLibChangelog.Visible = $true
$libCard.Controls.Add($btnLibChangelog)

$btnLibCancel = New-Object System.Windows.Forms.Button
$btnLibCancel.Text = "ביטול"
$btnLibCancel.Location = New-Object System.Drawing.Point(20,72)
$btnLibCancel.Size = New-Object System.Drawing.Size(100,34)
$btnLibCancel.FlatStyle = "Flat"
$btnLibCancel.Visible = $false
$libCard.Controls.Add($btnLibCancel)

# Extract button - for extracting existing library file
$btnLibExtract = New-Object System.Windows.Forms.Button
$btnLibExtract.Text = "חלץ לאוצריא"
$btnLibExtract.Location = New-Object System.Drawing.Point(347,72)
$btnLibExtract.Size = New-Object System.Drawing.Size(140,34)
$btnLibExtract.FlatStyle = "Flat"
$btnLibExtract.Visible = $false
$libCard.Controls.Add($btnLibExtract)

$btnLibDL = New-Object System.Windows.Forms.Button
$btnLibDL.Text = "הורדת הספרייה"
$btnLibDL.Location = New-Object System.Drawing.Point(237,72)
$btnLibDL.Size = New-Object System.Drawing.Size(180,34)
$btnLibDL.FlatStyle = "Flat"
$libCard.Controls.Add($btnLibDL)

# Select library file button (offline mode)
$btnLibSelectFile = New-Object System.Windows.Forms.Button
$btnLibSelectFile.Text = "בחר קובץ ספרייה"
$btnLibSelectFile.Location = New-Object System.Drawing.Point(227,72)
$btnLibSelectFile.Size = New-Object System.Drawing.Size(200,34)
$btnLibSelectFile.FlatStyle = "Flat"
$btnLibSelectFile.Visible = $false
$libCard.Controls.Add($btnLibSelectFile)

# Custom modern progress bar for library
$libProgBar = New-Object ModernProgressBar
$libProgBar.Location = New-Object System.Drawing.Point(20,112)
$libProgBar.Size = New-Object System.Drawing.Size(610,22)
$libProgBar.IsRTL = $true
$libProgBar.ShowText = $true
$libProgBar.Visible = $true
$libCard.Controls.Add($libProgBar)

$libStatusLbl = New-Object System.Windows.Forms.Label
$libStatusLbl.Text = ""
$libStatusLbl.Location = New-Object System.Drawing.Point(20,138)
$libStatusLbl.Size = New-Object System.Drawing.Size(610,22)
$libStatusLbl.TextAlign = "MiddleCenter"
$libStatusLbl.Font = New-Object System.Drawing.Font("Segoe UI",10)
$libStatusLbl.Visible = $false
$libCard.Controls.Add($libStatusLbl)

# Bottom Buttons
$btnClearCache = New-Object System.Windows.Forms.Button
$btnClearCache.Text = "מחיקת מטמון אוצריא"
$btnClearCache.Size = New-Object System.Drawing.Size(250,40)
$btnClearCache.Location = New-Object System.Drawing.Point(235,685)
$btnClearCache.FlatStyle = "Flat"
$btnClearCache.Font = New-Object System.Drawing.Font("Segoe UI Semibold",11)

$btnTheme = New-Object System.Windows.Forms.Button
$btnTheme.Text = "◐"
$btnTheme.Size = New-Object System.Drawing.Size(42,42)
$btnTheme.Location = New-Object System.Drawing.Point(590,10)
$btnTheme.FlatStyle = "Flat"
$btnTheme.Font = New-Object System.Drawing.Font("Segoe UI",16)
$btnTheme.Cursor = "Hand"
$toolTip.SetToolTip($btnTheme, "מצב כהה")
$form.Controls.Add($btnTheme)

$statusDivider = New-Object System.Windows.Forms.Label
$statusDivider.Location = New-Object System.Drawing.Point(15,730)
$statusDivider.Size = New-Object System.Drawing.Size(675,1)

$statusLocation = New-Object System.Windows.Forms.Label
$statusLocation.Text = "מיקום התקנה: טוען..."
$statusLocation.Location = New-Object System.Drawing.Point(15,738)
$statusLocation.Size = New-Object System.Drawing.Size(675,20)
$statusLocation.TextAlign = "MiddleCenter"
$statusLocation.Font = New-Object System.Drawing.Font("Segoe UI",9)

$statusBar = New-Object System.Windows.Forms.Label
$statusBar.Text = "טוען נתונים..."
$statusBar.Location = New-Object System.Drawing.Point(15,758)
$statusBar.Size = New-Object System.Drawing.Size(675,20)
$statusBar.TextAlign = "MiddleCenter"
$statusBar.Font = New-Object System.Drawing.Font("Segoe UI",9)

$lblNetwork = New-Object System.Windows.Forms.Label
$lblNetwork.Text = "אין חיבור לרשת"
$lblNetwork.Location = New-Object System.Drawing.Point(500,785)
$lblNetwork.Size = New-Object System.Drawing.Size(180,25)
$lblNetwork.ForeColor = [System.Drawing.Color]::FromArgb(200,50,50)
$lblNetwork.Font = New-Object System.Drawing.Font("Segoe UI Semibold",10)
$lblNetwork.TextAlign = "MiddleLeft"
$lblNetwork.Visible = $false

$form.Controls.AddRange(@($header,$mainCard,$libCard,$btnClearCache,$statusDivider,$statusLocation,$statusBar,$lblNetwork))

# Apply Theme Function
function Apply-Theme {
    $t = Get-Theme
    $form.BackColor = $t.BgColor
    $header.BackColor = $t.HeaderColor
    $title.ForeColor = $t.TextColor
    $sub.ForeColor = $t.TextLight
    $mainCard.BackColor = $t.CardColor
    $lblCurT.ForeColor = $t.TextLight
    $lblCurV.ForeColor = $t.TextColor
    $div1.BackColor = $t.Divider
    $lblSec.ForeColor = $t.TextColor
    $radioStable.ForeColor = $t.TextColor
    $radioStable.BackColor = $t.CardColor
    $lblStableV.ForeColor = $t.TextColor
    $radioPre.ForeColor = $t.TextColor
    $radioPre.BackColor = $t.CardColor
    $lblPreV.ForeColor = $t.TextLight
    $radioFull.ForeColor = $t.TextColor
    $radioFull.BackColor = $t.CardColor
    $lblFullV.ForeColor = $t.TextLight
    $radioDownloaded.ForeColor = $t.TextColor
    $radioDownloaded.BackColor = $t.CardColor
    $lblDownloadedV.ForeColor = $t.TextLight
    $lblInstallType.ForeColor = $t.TextColor
    $cmbInstallType.BackColor = $t.BtnBg
    $cmbInstallType.ForeColor = $t.TextColor
    $chkSilent.ForeColor = $t.TextColor
    $chkSilent.BackColor = $t.CardColor
    $progBar.BgColor = $t.ProgBg
    $progBar.BarColor = $t.ProgBar
    $progBar.TextColor = $t.TextColor
    $statusLbl.ForeColor = $t.TextLight
    $lblSizeInfo.ForeColor = $t.TextLight
    $libCard.BackColor = $t.CardColor
    $lblLib.ForeColor = $t.TextColor
    $lblLibV.ForeColor = $t.TextLight
    $libProgBar.BgColor = $t.ProgBg
    $libProgBar.BarColor = $t.ProgBar
    $libProgBar.TextColor = $t.TextColor
    $libStatusLbl.ForeColor = $t.TextLight
    $statusDivider.BackColor = $t.Divider
    $statusLocation.ForeColor = $t.TextLight
    $statusBar.ForeColor = $t.TextLight
    foreach ($btn in @($btnDL, $btnCancel, $btnPreChangelog, $btnLibDL, $btnLibCancel, $btnLibChangelog, $btnLibExtract, $btnLibSelectFile, $btnClearCache, $btnTheme, $btnRefresh, $btnSelectFile)) {
        $btn.BackColor = $t.BtnBg
        $btn.ForeColor = $t.TextColor
        $btn.FlatAppearance.BorderColor = $t.BtnBorder
    }
    $form.Refresh()
}

# Timer
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 200

$timer.Add_Tick({
    # Software Download
    if ($script:DownloadJob -and $script:IsDownloading) {
        $output = Receive-Job -Job $script:DownloadJob -ErrorAction SilentlyContinue
        foreach ($line in $output) {
            if ($line -match '^PROGRESS:(\d+):(\d+)$') {
                $bytes = [long]$matches[1]; $pct = [int]$matches[2]
                $progBar.Value = $pct
                $progBar.Refresh()
                $mbDone = [math]::Round($bytes / 1MB, 1); $mbTotal = [math]::Round($script:TotalSize / 1MB, 1)
                $statusLbl.Text = "הורד $mbDone מתוך $mbTotal מ`"ב"
                try { [TBProg]::SetVal($form.Handle, $pct, 100) } catch { }
            }
            elseif ($line -eq "DONE") {
                $script:IsDownloading = $false
                if ($script:TempFile -and (Test-Path $script:TempFile)) { Rename-Item $script:TempFile $script:FinalFile -Force -ErrorAction SilentlyContinue }
                $progBar.Value = 100; $statusLbl.Text = "ההורדה הושלמה!"; $btnDL.Text = "התקן"; $btnCancel.Visible = $false; $script:DownloadCompleted = $true
                $chkSilent.Visible = $true
                try { [TBProg]::SetState($form.Handle, 0) } catch { }
                Remove-Job -Job $script:DownloadJob -Force -ErrorAction SilentlyContinue; $script:DownloadJob = $null
            }
            elseif ($line -eq "STOPPED") {
                $script:IsDownloading = $false; $script:IsPaused = $true
                $downloaded = if ($script:TempFile -and (Test-Path $script:TempFile)) { (Get-Item $script:TempFile).Length } else { 0 }
                $btnDL.Text = "המשך"; $statusLbl.Text = "הופסק - הורדו $([math]::Round($downloaded/1MB,1)) מ`"ב"
                $chkSilent.Visible = $false
                try { [TBProg]::SetState($form.Handle, 8) } catch { }
                Remove-Job -Job $script:DownloadJob -Force -ErrorAction SilentlyContinue; $script:DownloadJob = $null
            }
            elseif ($line -match '^ERROR:(.*)$') {
                $script:IsDownloading = $false; $statusLbl.Text = "שגיאה בהורדה"; $btnDL.Text = "הורדה"; $btnCancel.Visible = $false
                try { [TBProg]::SetState($form.Handle, 0) } catch { }
                Remove-Job -Job $script:DownloadJob -Force -ErrorAction SilentlyContinue; $script:DownloadJob = $null
            }
        }
    }
    
    # Library Download
    if ($script:LibDownloadJob -and $script:LibIsDownloading) {
        $output = Receive-Job -Job $script:LibDownloadJob -ErrorAction SilentlyContinue
        foreach ($line in $output) {
            if ($line -match '^PROGRESS:(\d+):(\d+)$') {
                $bytes = [long]$matches[1]; $pct = [int]$matches[2]
                $libProgBar.Value = $pct
                $libProgBar.Refresh()
                $mbDone = [math]::Round($bytes / 1MB, 1); $mbTotal = [math]::Round($script:LibTotalSize / 1MB, 1)
                $libStatusLbl.Text = "הורד $mbDone מתוך $mbTotal מ`"ב"
                try { [TBProg]::SetVal($form.Handle, $pct, 100) } catch { }
            }
            elseif ($line -eq "DONE") {
                $script:LibIsDownloading = $false
                if ($script:LibTempFile -and (Test-Path $script:LibTempFile)) { Rename-Item $script:LibTempFile $script:LibFinalFile -Force -ErrorAction SilentlyContinue }
                $libProgBar.Value = 100; $libStatusLbl.Text = "ההורדה הושלמה!"; $btnLibDL.Text = "חלץ לאוצריא"; $btnLibCancel.Visible = $false; $script:LibDownloadCompleted = $true
                $script:LibFileToExtract = $script:LibFinalFile
                try { [TBProg]::SetState($form.Handle, 0) } catch { }
                Remove-Job -Job $script:LibDownloadJob -Force -ErrorAction SilentlyContinue; $script:LibDownloadJob = $null
            }
            elseif ($line -eq "STOPPED") {
                $script:LibIsDownloading = $false; $script:LibIsPaused = $true
                $downloaded = if ($script:LibTempFile -and (Test-Path $script:LibTempFile)) { (Get-Item $script:LibTempFile).Length } else { 0 }
                $btnLibDL.Text = "המשך"; $libStatusLbl.Text = "הופסק - הורדו $([math]::Round($downloaded/1MB,1)) מ`"ב"
                try { [TBProg]::SetState($form.Handle, 8) } catch { }
                Remove-Job -Job $script:LibDownloadJob -Force -ErrorAction SilentlyContinue; $script:LibDownloadJob = $null
            }
            elseif ($line -match '^ERROR:(.*)$') {
                $script:LibIsDownloading = $false; $libStatusLbl.Text = "שגיאה בהורדה"; $btnLibDL.Text = "הורדת הספרייה"; $btnLibCancel.Visible = $false
                try { [TBProg]::SetState($form.Handle, 0) } catch { }
                Remove-Job -Job $script:LibDownloadJob -Force -ErrorAction SilentlyContinue; $script:LibDownloadJob = $null
            }
        }
    }
})

function Update-Btn {
    if ($script:IsDownloading -or $script:IsPaused) { return }
    $s = $null
    
    if ($radioDownloaded.Checked) {
        # Downloaded version from folder
        $s = $script:DownloadedRelease
    } elseif ($script:InstallType -eq "MSIX") {
        # MSIX mode - only stable and beta, no full version
        if ($radioStable.Checked) { $s = $script:StableMsixRelease }
        elseif ($radioPre.Checked) { $s = $script:PreMsixRelease }
    } else {
        # EXE mode
        if ($radioStable.Checked) { $s = $script:StableRelease }
        elseif ($radioPre.Checked) { $s = $script:PreRelease }
        elseif ($radioFull.Checked) { $s = $script:FullRelease }
    }
    
    if ($s -and (Test-Path $s.File)) { 
        $btnDL.Enabled = $true
        $btnDL.Text = "התקן"
        $script:DownloadCompleted = $true
        $script:FinalFile = $s.File
        $chkSilent.Visible = $true
        $statusLbl.Text = "קובץ קיים - מוכן להתקנה"
    }
    elseif ($s -and $s.Url) { 
        $btnDL.Enabled = $true
        $btnDL.Text = "הורדה"
        $script:DownloadCompleted = $false
        $chkSilent.Visible = $false
        # Show download size inline with status
        $size = Get-Size $s.Url
        if ($size -gt 0) {
            $sizeMB = [math]::Round($size / 1MB, 1)
            $statusLbl.Text = "מוכן להורדה   |   גודל: $sizeMB מ`"ב"
        } else {
            $statusLbl.Text = "מוכן להורדה"
        }
    }
    else {
        $btnDL.Enabled = $false
        $btnDL.Text = "הורדה"
        $script:DownloadCompleted = $false
        $chkSilent.Visible = $false
        $statusLbl.Text = "בחר גרסה"
    }
}

function Check-DownloadedFiles {
    # Check for existing files in folder that don't match server versions
    $script:DownloadedRelease = $null
    $radioDownloaded.Visible = $false
    $lblDownloadedV.Visible = $false
    
    $currentPath = (Get-Location).Path
    [System.IO.Directory]::GetFiles($currentPath) | Out-Null
    $dirInfo = [System.IO.DirectoryInfo]::new($currentPath)
    
    if ($script:InstallType -eq "MSIX") {
        $allMsixFiles = @($dirInfo.GetFiles("*.msix", [System.IO.SearchOption]::TopDirectoryOnly) | Where-Object { $_.Name.StartsWith("otzaria-") })
        $stableVer = if ($script:StableMsixRelease) { $script:StableMsixRelease.FullVersion } else { "" }
        $preVer = if ($script:PreMsixRelease) { $script:PreMsixRelease.FullVersion } else { "" }
        
        # Find newest file that doesn't match server versions
        foreach ($msixFile in ($allMsixFiles | Sort-Object LastWriteTime -Descending)) {
            if ($msixFile.Name -match 'otzaria-([\d\.]+)\.msix$') {
                $fileVer = $matches[1]
                if ($fileVer -ne $stableVer -and $fileVer -ne $preVer) {
                    $script:DownloadedRelease = @{ FullVersion = $fileVer; File = $msixFile.Name; Url = $null }
                    $radioDownloaded.Visible = $true
                    $lblDownloadedV.Text = $fileVer
                    $lblDownloadedV.Visible = $true
                    break
                }
            }
        }
    } else {
        # EXE mode - check for regular and full EXE
        $allExeFiles = @($dirInfo.GetFiles("*.exe", [System.IO.SearchOption]::TopDirectoryOnly) | Where-Object { $_.Name.StartsWith("otzaria-") -and $_.Name.Contains("-windows") })
        
        $stableVer = if ($script:StableRelease) { $script:StableRelease.FullVersion } else { "" }
        $preVer = if ($script:PreRelease) { $script:PreRelease.FullVersion } else { "" }
        $fullVer = if ($script:FullRelease) { $script:FullRelease.FullVersion } else { "" }
        
        # Find newest file that doesn't match server versions
        foreach ($exeFile in ($allExeFiles | Sort-Object LastWriteTime -Descending)) {
            $fileVer = $null
            if ($exeFile.Name -match 'otzaria-([\d\.]+)-windows(-full)?\.exe$') {
                $fileVer = $matches[1]
            }
            if ($fileVer -and $fileVer -ne $stableVer -and $fileVer -ne $preVer -and $fileVer -ne $fullVer) {
                $script:DownloadedRelease = @{ FullVersion = $fileVer; File = $exeFile.Name; Url = $null }
                $radioDownloaded.Visible = $true
                $lblDownloadedV.Text = $fileVer
                $lblDownloadedV.Visible = $true
                break
            }
        }
    }
    
    # If downloaded radio was selected but no longer visible, switch to stable
    if ($radioDownloaded.Checked -and -not $radioDownloaded.Visible) {
        $radioStable.Checked = $true
    }
}

# Install Type ComboBox Change Handler
$cmbInstallType.Add_SelectedIndexChanged({
    $script:InstallType = $cmbInstallType.SelectedItem.ToString()
    
    if ($script:InstallType -eq "MSIX") {
        # Show MSIX versions
        $radioFull.Visible = $false
        $lblFullV.Visible = $false
        if ($script:StableMsixRelease) { $lblStableV.Text = $script:StableMsixRelease.FullVersion }
        else { $lblStableV.Text = "לא זמין" }
        if ($script:PreMsixRelease) { $lblPreV.Text = $script:PreMsixRelease.FullVersion }
        else { $lblPreV.Text = "לא זמין" }
        if ($radioFull.Checked -or $radioDownloaded.Checked) { $radioStable.Checked = $true }
    } else {
        # Show EXE versions
        $radioFull.Visible = $true
        $lblFullV.Visible = $true
        if ($script:StableRelease) { $lblStableV.Text = $script:StableRelease.FullVersion }
        else { $lblStableV.Text = "לא זמין" }
        if ($script:PreRelease) { $lblPreV.Text = $script:PreRelease.FullVersion }
        else { $lblPreV.Text = "לא זמין" }
        if ($script:FullRelease) { $lblFullV.Text = $script:FullRelease.FullVersion }
        else { $lblFullV.Text = "לא זמין" }
    }
    
    Check-DownloadedFiles
    Update-Btn
})

# Form Events
$form.Add_Shown({
    Apply-Theme
    [System.Windows.Forms.Application]::DoEvents()
    $timer.Start()
    
    $script:InstallPath = Find-OtzariaInstallPath
    if ($script:InstallPath) {
        $script:InstalledVersion = Get-InstalledVersion $script:InstallPath
        $script:InstalledLibVersion = Get-InstalledLibVersion $script:InstallPath
        # Use LRM (Left-to-Right Mark) to display path correctly
        $lrm = [char]0x200E
        $statusLocation.Text = "מיקום התקנה: $lrm$($script:InstallPath)"
    } else { 
        $statusLocation.Text = "מיקום התקנה: לא נמצא"
    }
    
    # Check for MSIX installation
    $msixInfo = Get-MsixInstallation
    if ($msixInfo.Installed) {
        $script:MsixInstalled = $true
        $script:MsixVersion = $msixInfo.Version
        $script:InstallType = "MSIX"
        $cmbInstallType.SelectedIndex = 1  # Select MSIX
        $script:InstalledVersion = $msixInfo.Version
        $lrm = [char]0x200E
        $statusLocation.Text = "מיקום התקנה: אפליקציות (MSIX)"
    }
    
    $statusBar.Text = "גרסה מותקנת: $($script:InstalledVersion)   |   גרסת ספרייה מותקנת: $($script:InstalledLibVersion)"
    
    # Load data from server - sequential but fast
    $statusLbl.Text = "טוען נתונים מהשרת..."
    [System.Windows.Forms.Application]::DoEvents()
    
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $wc = New-Object System.Net.WebClient
    $wc.Encoding = [System.Text.Encoding]::UTF8
    $wc.Headers.Add("User-Agent", "PS")
    
    # Get main releases
    try {
        $json = $wc.DownloadString("https://api.github.com/repos/Otzaria/otzaria/releases") | ConvertFrom-Json
        $lblNetwork.Visible = $false
        
        # Process stable release (EXE)
        $stableRel = $json | Where-Object { $_.prerelease -eq $false } | Select-Object -First 1
        if ($stableRel) {
            # Parse tag version once
            $stableVer = $null
            if ($stableRel.tag_name -match '^([\d\.]+)\+(\d+)$') {
                $stableVer = "$($matches[1]).$($matches[2])"
            }
            
            $asset = $stableRel.assets | Where-Object { $_.name -like "otzaria-*-windows.exe" -and $_.name -notlike "*-full*" } | Select-Object -First 1
            if ($asset -and $stableVer) {
                $script:StableRelease = @{
                    FullVersion = $stableVer
                    File = "otzaria-$stableVer-windows.exe"
                    Url = $asset.browser_download_url
                }
                $script:StableChangelog = $stableRel.body
            }
            # Check for MSIX in stable
            $msixAsset = $stableRel.assets | Where-Object { $_.name -like "*.msix" } | Select-Object -First 1
            if ($msixAsset -and $stableVer) {
                $script:StableMsixRelease = @{
                    FullVersion = $stableVer
                    File = "otzaria-$stableVer.msix"
                    Url = $msixAsset.browser_download_url
                }
            }
        }
        
        # Process pre-release (EXE)
        $preRel = $json | Where-Object { $_.prerelease -eq $true } | Select-Object -First 1
        if ($preRel) {
            # Parse tag version once
            $preVer = $null
            if ($preRel.tag_name -match '^([\d\.]+)\+(\d+)$') {
                $preVer = "$($matches[1]).$($matches[2])"
            }
            
            $asset = $preRel.assets | Where-Object { $_.name -like "otzaria-*-windows.exe" -and $_.name -notlike "*-full*" } | Select-Object -First 1
            if ($asset -and $preVer) {
                $script:PreRelease = @{
                    FullVersion = $preVer
                    File = "otzaria-$preVer-windows.exe"
                    Url = $asset.browser_download_url
                }
                $script:PreChangelog = $preRel.body
            }
            # Check for MSIX in pre-release
            $msixAsset = $preRel.assets | Where-Object { $_.name -like "*.msix" } | Select-Object -First 1
            if ($msixAsset -and $preVer) {
                $script:PreMsixRelease = @{
                    FullVersion = $preVer
                    File = "otzaria-$preVer.msix"
                    Url = $msixAsset.browser_download_url
                }
            }
        }
        
        # Search for full EXE in ALL releases (it might not be in stable or pre)
        foreach ($rel in $json) {
            $fullAsset = $rel.assets | Where-Object { $_.name -like "otzaria-*-windows-full.exe" } | Select-Object -First 1
            if ($fullAsset) {
                $fullVer = $null
                $fullFileName = $fullAsset.name
                # Get version from tag (includes build number)
                if ($rel.tag_name -match '^([\d\.]+)\+(\d+)$') {
                    $fullVer = "$($matches[1]).$($matches[2])"
                    # Create filename with build number
                    $fullFileName = "otzaria-$fullVer-windows-full.exe"
                } elseif ($fullAsset.name -match 'otzaria-([\d\.]+)-windows-full\.exe') {
                    # Fallback to filename if tag doesn't match
                    $fullVer = $matches[1]
                }
                if ($fullVer) {
                    $script:FullRelease = @{
                        FullVersion = $fullVer
                        File = $fullFileName
                        Url = $fullAsset.browser_download_url
                    }
                    break  # Found it, stop searching
                }
            }
        }
    } catch {
        $lblNetwork.Visible = $true
    }
    
    # Fallback: Get full release from Sivan22/otzaria-full if not found in main repo
    if (-not $script:FullRelease) {
        try {
            $fullRepo = $wc.DownloadString("https://api.github.com/repos/Sivan22/otzaria-full/releases/latest") | ConvertFrom-Json
            if ($fullRepo) {
                $fullAsset = $fullRepo.assets | Where-Object { $_.name -like "otzaria-*-windows-full.exe" } | Select-Object -First 1
                if ($fullAsset) {
                    # Try different tag formats
                    $fullVer = $null
                    if ($fullRepo.tag_name -match '^([\d\.]+)\+(\d+)$') {
                        $fullVer = "$($matches[1]).$($matches[2])"
                    } elseif ($fullRepo.tag_name -match '^v?([\d\.]+)$') {
                        $fullVer = $matches[1]
                    } elseif ($fullAsset.name -match 'otzaria-([\d\.]+)-windows-full\.exe') {
                        $fullVer = $matches[1]
                    }
                    if ($fullVer) {
                        $script:FullRelease = @{
                            FullVersion = $fullVer
                            File = $fullAsset.name
                            Url = $fullAsset.browser_download_url
                        }
                    }
                }
            }
        } catch { }
    }
    
    # Get library version (separate try/catch)
    try {
        $script:LibVer = $wc.DownloadString("https://github.com/Otzaria/otzaria-library/raw/refs/heads/main/MoreBooks/%D7%A1%D7%A4%D7%A8%D7%99%D7%9D/%D7%90%D7%95%D7%A6%D7%A8%D7%99%D7%90/%D7%90%D7%95%D7%93%D7%95%D7%AA%20%D7%94%D7%AA%D7%95%D7%9B%D7%A0%D7%94/%D7%92%D7%99%D7%A8%D7%A1%D7%AA%20%D7%A1%D7%A4%D7%A8%D7%99%D7%94.txt").Trim()
    } catch {
        $script:LibVer = "?"
    }
    
    # Display versions based on install type
    if ($script:InstallType -eq "MSIX") {
        $radioFull.Visible = $false
        $lblFullV.Visible = $false
        if ($script:StableMsixRelease) { $lblStableV.Text = $script:StableMsixRelease.FullVersion }
        else { $lblStableV.Text = "לא זמין" }
        if ($script:PreMsixRelease) { $lblPreV.Text = $script:PreMsixRelease.FullVersion }
        else { $lblPreV.Text = "לא זמין" }
    } else {
        if ($script:StableRelease) { $lblStableV.Text = $script:StableRelease.FullVersion }
        else { $lblStableV.Text = "לא זמין" }
        if ($script:PreRelease) { $lblPreV.Text = $script:PreRelease.FullVersion }
        else { $lblPreV.Text = "לא זמין" }
        if ($script:FullRelease) { $lblFullV.Text = $script:FullRelease.FullVersion }
        else { $lblFullV.Text = "לא זמין" }
    }
    
    # Process library version
    if (-not $script:LibVer) { $script:LibVer = "?" }
    $script:LibFinalFile = "otzaria_latest_$($script:LibVer).zip"
    $script:LibTempFile = "otzaria_latest_$($script:LibVer).part"
    
    # Check for existing library files first
    $existingLibZip = Get-ChildItem "." -Filter "otzaria_latest*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($existingLibZip) {
        # Library file exists - check if newer version available
        $zipVer = Get-ZipLibVersion $existingLibZip.Name
        $script:LibFileToExtract = $existingLibZip.Name
        
        # Compare versions to see if newer is available on server
        $newerAvailable = $false
        if ($zipVer -and $script:LibVer -and $script:LibVer -ne "?") {
            try {
                $localParts = $zipVer -split '\.'
                $serverParts = $script:LibVer -split '\.'
                for ($i = 0; $i -lt [Math]::Max($localParts.Count, $serverParts.Count); $i++) {
                    $local = if ($i -lt $localParts.Count) { [int]$localParts[$i] } else { 0 }
                    $server = if ($i -lt $serverParts.Count) { [int]$serverParts[$i] } else { 0 }
                    if ($server -gt $local) { $newerAvailable = $true; break }
                    if ($local -gt $server) { break }
                }
            } catch { }
        }
        
        if ($newerAvailable) {
            # Newer version on server - show both buttons side by side centered
            $libSize = Get-Size $script:LibUrl
            if ($libSize -gt 0) {
                $libSizeMB = [math]::Round($libSize / 1MB, 0)
                $lblLibV.Text = "גרסה בשרת: $($script:LibVer)   |   גודל: $libSizeMB מ`"ב"
            } else {
                $lblLibV.Text = "גרסה בשרת: $($script:LibVer)"
            }
            $btnLibDL.Text = "הורדה"
            $btnLibDL.Size = New-Object System.Drawing.Size(120,34)
            $btnLibDL.Location = New-Object System.Drawing.Point(192,72)
            $btnLibDL.Visible = $true
            $btnLibExtract.Size = New-Object System.Drawing.Size(140,34)
            $btnLibExtract.Location = New-Object System.Drawing.Point(322,72)
            $btnLibExtract.Visible = $true
            $btnLibSelectFile.Visible = $false
            $script:LibDownloadCompleted = $false
            $libStatusLbl.Visible = $true
            $libStatusLbl.Text = "קובץ ספרייה גרסה $zipVer קיים - גרסה חדשה יותר זמינה"
        } else {
            # Same or newer version locally - just show extract centered
            $lblLibV.Text = "גרסה בשרת: $($script:LibVer)"
            $script:LibDownloadCompleted = $true
            $btnLibDL.Text = "חלץ לאוצריא"
            $btnLibDL.Size = New-Object System.Drawing.Size(180,34)
            $btnLibDL.Location = New-Object System.Drawing.Point(237,72)
            $btnLibDL.Visible = $true
            $btnLibExtract.Visible = $false
            $btnLibSelectFile.Visible = $false
            $libStatusLbl.Visible = $true
            if ($zipVer) {
                $libStatusLbl.Text = "קובץ ספרייה גרסה $zipVer קיים בתיקייה - לחץ לחילוץ"
            } else {
                $libStatusLbl.Text = "קובץ ספרייה קיים בתיקייה - לחץ לחילוץ"
            }
        }
    } else {
        # No library file - show server version with download size or select file button
        if ($script:LibVer -and $script:LibVer -ne "?") {
            # Online - show download option centered
            $lblLibV.Text = "גרסה בשרת: $($script:LibVer)"
            $libSize = Get-Size $script:LibUrl
            if ($libSize -gt 0) {
                $libSizeMB = [math]::Round($libSize / 1MB, 0)
                $lblLibV.Text = "גרסה בשרת: $($script:LibVer)   |   גודל: $libSizeMB מ`"ב"
            }
            $btnLibDL.Text = "הורדת הספרייה"
            $btnLibDL.Size = New-Object System.Drawing.Size(180,34)
            $btnLibDL.Location = New-Object System.Drawing.Point(237,72)
            $btnLibDL.Visible = $true
            $btnLibSelectFile.Visible = $false
        } else {
            # Offline and no file - show select file button centered
            $lblLibV.Text = "גרסה בשרת: ?"
            $btnLibDL.Visible = $false
            $btnLibSelectFile.Visible = $true
            $libStatusLbl.Visible = $true
            $libStatusLbl.Text = "אין חיבור לרשת - בחר קובץ ספרייה"
        }
        $btnLibExtract.Visible = $false
    }
    
    # Check for existing EXE files in folder (offline mode support)
    $existingExe = Get-ChildItem "." -Filter "otzaria-*-windows.exe" -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*-full.exe" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $existingFullExe = Get-ChildItem "." -Filter "otzaria-*-windows-full.exe" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $existingMsix = Get-ChildItem "." -Filter "*.msix" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    
    # Offline mode - no network releases available
    if (-not $script:StableRelease -and -not $script:PreRelease -and -not $script:FullRelease -and -not $script:StableMsixRelease -and -not $script:PreMsixRelease) {
        $lblNetwork.Visible = $true
        
        # Detect ALL existing files and create releases for them
        if ($existingMsix -and $existingMsix.Name -match 'otzaria-([\d\.]+)\.msix$') {
            $script:StableMsixRelease = @{ FullVersion = $matches[1]; File = $existingMsix.Name; Url = $null }
        }
        if ($existingExe -and $existingExe.Name -match 'otzaria-([\d\.]+)-windows\.exe') {
            $script:StableRelease = @{ FullVersion = $matches[1]; File = $existingExe.Name; Url = $null }
        }
        if ($existingFullExe -and $existingFullExe.Name -match 'otzaria-([\d\.]+)-windows-full\.exe') {
            $script:FullRelease = @{ FullVersion = $matches[1]; File = $existingFullExe.Name; Url = $null }
        }
        
        # Update UI based on what was found
        $hasExeFiles = $script:StableRelease -or $script:FullRelease
        $hasMsixFiles = $script:StableMsixRelease
        
        if ($hasExeFiles -or $hasMsixFiles) {
            # Set install type based on current selection or available files
            if ($script:InstallType -eq "MSIX" -and $hasMsixFiles) {
                $cmbInstallType.SelectedIndex = 1
                $lblStableV.Text = if ($script:StableMsixRelease) { $script:StableMsixRelease.FullVersion } else { "לא זמין" }
                $lblPreV.Text = "לא זמין"
                $radioFull.Visible = $false
                $lblFullV.Visible = $false
            } elseif ($hasExeFiles) {
                $script:InstallType = "EXE"
                $cmbInstallType.SelectedIndex = 0
                $lblStableV.Text = if ($script:StableRelease) { $script:StableRelease.FullVersion } else { "לא זמין" }
                $lblPreV.Text = "לא זמין"
                $lblFullV.Text = if ($script:FullRelease) { $script:FullRelease.FullVersion } else { "לא זמין" }
                $radioFull.Visible = $true
                $lblFullV.Visible = $true
            } elseif ($hasMsixFiles) {
                $script:InstallType = "MSIX"
                $cmbInstallType.SelectedIndex = 1
                $lblStableV.Text = if ($script:StableMsixRelease) { $script:StableMsixRelease.FullVersion } else { "לא זמין" }
                $lblPreV.Text = "לא זמין"
                $radioFull.Visible = $false
                $lblFullV.Visible = $false
            }
            
            $radioStable.Checked = $true
            $btnDL.Visible = $true
            $btnSelectFile.Visible = $false
            $statusLbl.Text = "קובץ התקנה נמצא - מוכן להתקנה"
            Update-Btn
        } else {
            # No files found - show select file button
            $statusLbl.Text = "אין חיבור לרשת - בחר קובץ להתקנה"
            $btnDL.Visible = $false
            $btnSelectFile.Visible = $true
        }
    } else { 
        $statusLbl.Text = "מוכן להורדה"
        
        # Check for partial software download to resume
        if ($script:StableRelease) {
            $partFile = "$($script:StableRelease.File).part"
            if (Test-Path $partFile) {
                $script:TempFile = $partFile
                $script:FinalFile = $script:StableRelease.File
                $script:SelectedRelease = $script:StableRelease
                $script:IsPaused = $true
                $script:TotalSize = Get-Size $script:StableRelease.Url
                $downloaded = (Get-Item $partFile).Length
                $pct = if ($script:TotalSize -gt 0) { [math]::Round(($downloaded / $script:TotalSize) * 100) } else { 0 }
                $progBar.Value = $pct
                $statusLbl.Text = "נמצאה הורדה קודמת - $([math]::Round($downloaded/1MB,1)) מ`"ב"
                $btnDL.Text = "המשך"
                $btnDL.Enabled = $true
                $radioStable.Checked = $true
            }
        }
        if (-not $script:IsPaused -and $script:PreRelease) {
            $partFile = "$($script:PreRelease.File).part"
            if (Test-Path $partFile) {
                $script:TempFile = $partFile
                $script:FinalFile = $script:PreRelease.File
                $script:SelectedRelease = $script:PreRelease
                $script:IsPaused = $true
                $script:TotalSize = Get-Size $script:PreRelease.Url
                $downloaded = (Get-Item $partFile).Length
                $pct = if ($script:TotalSize -gt 0) { [math]::Round(($downloaded / $script:TotalSize) * 100) } else { 0 }
                $progBar.Value = $pct
                $statusLbl.Text = "נמצאה הורדה קודמת - $([math]::Round($downloaded/1MB,1)) מ`"ב"
                $btnDL.Text = "המשך"
                $btnDL.Enabled = $true
                $radioPre.Checked = $true
            }
        }
        
        if (-not $script:IsPaused) { Update-Btn }
    }
    
    # Check for partial library download to resume
    if (-not $script:LibDownloadCompleted) {
        if (Test-Path $script:LibTempFile) {
            $script:LibIsPaused = $true
            $script:LibTotalSize = Get-Size $script:LibUrl
            $downloaded = (Get-Item $script:LibTempFile).Length
            $pct = if ($script:LibTotalSize -gt 0) { [math]::Round(($downloaded / $script:LibTotalSize) * 100) } else { 0 }
            $libProgBar.Value = $pct
            $libStatusLbl.Visible = $true
            $libStatusLbl.Text = "נמצאה הורדה קודמת - $([math]::Round($downloaded/1MB,1)) מ`"ב"
            $btnLibDL.Text = "המשך"
        }
        elseif (Test-Path "otzaria_latest.part") {
            $script:LibTempFile = "otzaria_latest.part"
            $script:LibFinalFile = "otzaria_latest.zip"
            $script:LibIsPaused = $true
            $script:LibTotalSize = Get-Size $script:LibUrl
            $downloaded = (Get-Item $script:LibTempFile).Length
            $pct = if ($script:LibTotalSize -gt 0) { [math]::Round(($downloaded / $script:LibTotalSize) * 100) } else { 0 }
            $libProgBar.Value = $pct
            $libStatusLbl.Visible = $true
            $libStatusLbl.Text = "נמצאה הורדה קודמת - $([math]::Round($downloaded/1MB,1)) מ`"ב"
            $btnLibDL.Text = "המשך"
        }
    }
    
    if ($script:InstalledLibVersion -ne "לא נמצא" -and -not $script:LibDownloadCompleted -and -not $script:LibIsPaused) {
        # Compare installed version with server version
        try {
            $installedNum = [int]($script:InstalledLibVersion -replace '\D','')
            $serverNum = [int]($script:LibVer -replace '\D','')
            if ($installedNum -ge $serverNum) {
                $libStatusLbl.Visible = $true
                $libStatusLbl.Text = "הספרייה מעודכנת"
            }
        } catch { }
    }
    
    # Check for downloaded files that don't match server versions
    Check-DownloadedFiles
})

$form.Add_FormClosing({
    $stopFile = "$env:TEMP\otzaria_stop.flag"
    $libStopFile = "$env:TEMP\otzaria_lateststop.flag"
    if ($script:DownloadJob) {
        "stop" | Out-File $stopFile -Force
        Start-Sleep -Milliseconds 500
        Stop-Job -Job $script:DownloadJob -ErrorAction SilentlyContinue
        Remove-Job -Job $script:DownloadJob -Force -ErrorAction SilentlyContinue
    }
    if ($script:LibDownloadJob) {
        "stop" | Out-File $libStopFile -Force
        Start-Sleep -Milliseconds 500
        Stop-Job -Job $script:LibDownloadJob -ErrorAction SilentlyContinue
        Remove-Job -Job $script:LibDownloadJob -Force -ErrorAction SilentlyContinue
    }
    $timer.Stop()
})

$radioStable.Add_CheckedChanged({ Update-Btn })
$radioPre.Add_CheckedChanged({ Update-Btn })
$radioFull.Add_CheckedChanged({ Update-Btn })
$radioDownloaded.Add_CheckedChanged({ Update-Btn })

# Select File Button (offline mode)
$btnSelectFile.Add_Click({
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Title = "בחר קובץ התקנה"
    $openDialog.Filter = "קבצי התקנה (*.exe;*.msix)|*.exe;*.msix|כל הקבצים (*.*)|*.*"
    $openDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($openDialog.ShowDialog() -eq "OK") {
        $selectedFile = $openDialog.FileName
        $fileName = [System.IO.Path]::GetFileName($selectedFile)
        
        # Copy to current directory if not already there
        $targetPath = Join-Path (Get-Location) $fileName
        if ($selectedFile -ne $targetPath) {
            Copy-Item $selectedFile $targetPath -Force
        }
        
        # Detect version from filename
        $version = "לא ידוע"
        if ($fileName -match 'otzaria-([\d\.]+)') {
            $version = $matches[1]
        }
        
        # Detect type and set appropriate release
        if ($fileName -like "*.msix") {
            $script:InstallType = "MSIX"
            $cmbInstallType.SelectedIndex = 1
            $script:StableMsixRelease = @{ FullVersion = $version; File = $fileName; Url = $null }
            $lblStableV.Text = $version
            $radioFull.Visible = $false
            $lblFullV.Visible = $false
        } elseif ($fileName -like "*-full.exe") {
            $script:InstallType = "EXE"
            $cmbInstallType.SelectedIndex = 0
            $script:FullRelease = @{ FullVersion = $version; File = $fileName; Url = $null }
            $lblFullV.Text = $version
            $radioFull.Checked = $true
        } else {
            $script:InstallType = "EXE"
            $cmbInstallType.SelectedIndex = 0
            $script:StableRelease = @{ FullVersion = $version; File = $fileName; Url = $null }
            $lblStableV.Text = $version
            $radioStable.Checked = $true
        }
        
        $script:DownloadCompleted = $true
        $script:FinalFile = $fileName
        $btnDL.Text = "התקן"
        $btnDL.Enabled = $true
        $btnDL.Visible = $true
        $btnSelectFile.Visible = $false
        if ($script:InstallType -eq "EXE") { $chkSilent.Visible = $true }
        $statusLbl.Text = "קובץ נבחר: $fileName"
    }
})

# Beta Changelog Button
$btnPreChangelog.Add_Click({
    $cl = Get-BetaChangelog
    Show-Changelog "מה חדש בגרסת הבטא?" $cl
})

# Library Changelog Button
$btnLibChangelog.Add_Click({
    $cl = Get-LibChangelog
    Show-Changelog "עדכוני ספרייה" $cl
})

# Theme Button
$btnTheme.Add_Click({
    $script:DarkMode = -not $script:DarkMode
    if ($script:DarkMode) {
        $btnTheme.Text = "◐"
        $toolTip.SetToolTip($btnTheme, "מצב בהיר")
    } else {
        $btnTheme.Text = "◐"
        $toolTip.SetToolTip($btnTheme, "מצב כהה")
    }
    Apply-Theme
})

# Refresh Button
$btnRefresh.Add_Click({
    # Reset states
    $script:IsDownloading = $false
    $script:IsPaused = $false
    $script:DownloadCompleted = $false
    $script:LibIsDownloading = $false
    $script:LibIsPaused = $false
    $script:LibDownloadCompleted = $false
    
    # Reset releases
    $script:StableRelease = $null
    $script:PreRelease = $null
    $script:FullRelease = $null
    $script:StableMsixRelease = $null
    $script:PreMsixRelease = $null
    
    # Reset UI
    $lblStableV.Text = "טוען..."
    $lblPreV.Text = "טוען..."
    $lblFullV.Text = "טוען..."
    $lblLibV.Text = "גרסה בשרת: טוען..."
    $lblCurV.Text = (Get-CurVer)
    $statusLbl.Text = "בודק חיבור לרשת..."
    $libStatusLbl.Text = ""
    $libStatusLbl.Visible = $false
    $progBar.Value = 0
    $libProgBar.Value = 0
    $btnDL.Text = "הורדה"
    $btnDL.Enabled = $false
    $btnDL.Visible = $true
    $btnSelectFile.Visible = $false
    $btnLibDL.Text = "הורדת הספרייה"
    $btnLibDL.Size = New-Object System.Drawing.Size(180,34)
    $btnLibDL.Location = New-Object System.Drawing.Point(237,72)
    $btnLibSelectFile.Visible = $false
    $chkSilent.Visible = $false
    $lblNetwork.Visible = $false
    [System.Windows.Forms.Application]::DoEvents()
    
    # Reload install info
    $script:InstallPath = Find-OtzariaInstallPath
    if ($script:InstallPath) {
        $script:InstalledVersion = Get-InstalledVersion $script:InstallPath
        $script:InstalledLibVersion = Get-InstalledLibVersion $script:InstallPath
        $lrm = [char]0x200E
        $statusLocation.Text = "מיקום התקנה: $lrm$($script:InstallPath)"
    } else { 
        $script:InstalledVersion = "לא נמצא"
        $script:InstalledLibVersion = "לא נמצא"
        $statusLocation.Text = "מיקום התקנה: לא נמצא"
    }
    $statusBar.Text = "גרסה מותקנת: $($script:InstalledVersion)   |   גרסת ספרייה מותקנת: $($script:InstalledLibVersion)"
    
    # Try to load data from server
    $statusLbl.Text = "מרענן נתונים..."
    [System.Windows.Forms.Application]::DoEvents()
    
    # Reload from server
    $stableData = Get-Release $false
    $script:StableRelease = $stableData.Release
    $script:StableChangelog = $stableData.Changelog
    
    $preData = Get-Release $true
    $script:PreRelease = $preData.Release
    $script:PreChangelog = $preData.Changelog
    
    $script:FullRelease = Get-FullRelease
    
    # Load MSIX releases
    $script:StableMsixRelease = Get-MsixRelease $false
    $script:PreMsixRelease = Get-MsixRelease $true
    
    # Check if we got any data from server
    $hasNetwork = $script:StableRelease -or $script:PreRelease -or $script:FullRelease -or $script:StableMsixRelease -or $script:PreMsixRelease
    
    if ($hasNetwork) {
        $lblNetwork.Visible = $false
        
        # Update labels based on install type
        if ($script:InstallType -eq "MSIX") {
            if ($script:StableMsixRelease) { $lblStableV.Text = $script:StableMsixRelease.FullVersion }
            else { $lblStableV.Text = "לא זמין" }
            if ($script:PreMsixRelease) { $lblPreV.Text = $script:PreMsixRelease.FullVersion }
            else { $lblPreV.Text = "לא זמין" }
        } else {
            if ($script:StableRelease) { $lblStableV.Text = $script:StableRelease.FullVersion }
            else { $lblStableV.Text = "לא זמין" }
            if ($script:PreRelease) { $lblPreV.Text = $script:PreRelease.FullVersion }
            else { $lblPreV.Text = "לא זמין" }
            if ($script:FullRelease) { $lblFullV.Text = $script:FullRelease.FullVersion }
            else { $lblFullV.Text = "לא זמין" }
        }
        
        $script:LibVer = Get-LibVer
        $script:LibFinalFile = "otzaria_latest_$($script:LibVer).zip"
        $script:LibTempFile = "otzaria_latest_$($script:LibVer).part"
    
        # Check for existing library files first
        $existingLibZip = Get-ChildItem "." -Filter "otzaria_latest*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($existingLibZip) {
            $zipVer = Get-ZipLibVersion $existingLibZip.Name
            # Library file exists - still show server version
            $lblLibV.Text = "גרסה בשרת: $($script:LibVer)"
            $script:LibDownloadCompleted = $true
            $script:LibFileToExtract = $existingLibZip.Name
            $btnLibDL.Text = "חלץ לאוצריא"
            $btnLibDL.Size = New-Object System.Drawing.Size(180,34)
            $btnLibDL.Location = New-Object System.Drawing.Point(237,72)
            $btnLibExtract.Visible = $false
            $libStatusLbl.Visible = $true
            if ($zipVer) {
                $libStatusLbl.Text = "קובץ ספרייה גרסה $zipVer קיים בתיקייה - לחץ לחילוץ"
            } else {
                $libStatusLbl.Text = "קובץ ספרייה קיים בתיקייה - לחץ לחילוץ"
            }
        } elseif (Test-Path "otzaria_latest.zip") {
            $lblLibV.Text = "גרסה בשרת: $($script:LibVer)"
            $script:LibDownloadCompleted = $true
            $script:LibFileToExtract = "otzaria_latest.zip"
            $btnLibDL.Text = "חלץ לאוצריא"
            $btnLibDL.Size = New-Object System.Drawing.Size(180,34)
            $btnLibDL.Location = New-Object System.Drawing.Point(237,72)
            $btnLibExtract.Visible = $false
            $libStatusLbl.Visible = $true
            $zipVer = Get-ZipLibVersion "otzaria_latest.zip"
            if ($zipVer) {
                $libStatusLbl.Text = "קובץ ספרייה גרסה $zipVer קיים בתיקייה - לחץ לחילוץ"
            } else {
                $libStatusLbl.Text = "קובץ ספרייה קיים בתיקייה - לחץ לחילוץ"
            }
        } else {
            # Show server version with download size
            $lblLibV.Text = "גרסה בשרת: $($script:LibVer)"
            $libSize = Get-Size $script:LibUrl
            if ($libSize -gt 0) {
                $libSizeMB = [math]::Round($libSize / 1MB, 0)
                $lblLibV.Text = "גרסה בשרת: $($script:LibVer)   |   גודל: $libSizeMB מ`"ב"
            }
            $btnLibExtract.Visible = $false
        }
        
        Check-DownloadedFiles
        $statusLbl.Text = "מוכן להורדה"
        Update-Btn
    } else {
        # No network - check for local files
        $lblNetwork.Visible = $true
        $lblStableV.Text = "לא זמין"
        $lblPreV.Text = "לא זמין"
        $lblFullV.Text = "לא זמין"
        $lblLibV.Text = "גרסה בשרת: ?"
        
        # Check for existing installation files in folder
        $existingExe = Get-ChildItem "." -Filter "otzaria-*-windows.exe" -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*-full.exe" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $existingFullExe = Get-ChildItem "." -Filter "otzaria-*-windows-full.exe" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $existingMsix = Get-ChildItem "." -Filter "*.msix" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        
        # Detect ALL existing files and create releases for them
        if ($existingMsix -and $existingMsix.Name -match 'otzaria-([\d\.]+)\.msix$') {
            $script:StableMsixRelease = @{ FullVersion = $matches[1]; File = $existingMsix.Name; Url = $null }
        }
        if ($existingExe -and $existingExe.Name -match 'otzaria-([\d\.]+)-windows\.exe') {
            $script:StableRelease = @{ FullVersion = $matches[1]; File = $existingExe.Name; Url = $null }
        }
        if ($existingFullExe -and $existingFullExe.Name -match 'otzaria-([\d\.]+)-windows-full\.exe') {
            $script:FullRelease = @{ FullVersion = $matches[1]; File = $existingFullExe.Name; Url = $null }
        }
        
        $hasExeFiles = $script:StableRelease -or $script:FullRelease
        $hasMsixFiles = $script:StableMsixRelease
        
        if ($hasExeFiles -or $hasMsixFiles) {
            # Update UI based on what was found
            if ($script:InstallType -eq "MSIX" -and $hasMsixFiles) {
                $lblStableV.Text = if ($script:StableMsixRelease) { $script:StableMsixRelease.FullVersion } else { "לא זמין" }
                $lblPreV.Text = "לא זמין"
            } elseif ($hasExeFiles) {
                $script:InstallType = "EXE"
                $cmbInstallType.SelectedIndex = 0
                $lblStableV.Text = if ($script:StableRelease) { $script:StableRelease.FullVersion } else { "לא זמין" }
                $lblPreV.Text = "לא זמין"
                $lblFullV.Text = if ($script:FullRelease) { $script:FullRelease.FullVersion } else { "לא זמין" }
                $radioFull.Visible = $true
                $lblFullV.Visible = $true
            } elseif ($hasMsixFiles) {
                $script:InstallType = "MSIX"
                $cmbInstallType.SelectedIndex = 1
                $lblStableV.Text = if ($script:StableMsixRelease) { $script:StableMsixRelease.FullVersion } else { "לא זמין" }
                $lblPreV.Text = "לא זמין"
            }
            
            $radioStable.Checked = $true
            $btnDL.Visible = $true
            $btnSelectFile.Visible = $false
            $statusLbl.Text = "קובץ התקנה נמצא - מוכן להתקנה"
            Update-Btn
        } else {
            $statusLbl.Text = "אין חיבור לרשת - בחר קובץ להתקנה"
            $btnDL.Visible = $false
            $btnSelectFile.Visible = $true
        }
        
        # Check for existing library files in offline mode
        $existingLibZip = Get-ChildItem "." -Filter "otzaria_latest*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($existingLibZip) {
            $zipVer = Get-ZipLibVersion $existingLibZip.Name
            $script:LibDownloadCompleted = $true
            $script:LibFileToExtract = $existingLibZip.Name
            $btnLibDL.Text = "חלץ לאוצריא"
            $btnLibDL.Visible = $true
            $btnLibSelectFile.Visible = $false
            $libStatusLbl.Visible = $true
            if ($zipVer) {
                $lblLibV.Text = "גרסה בקובץ: $zipVer"
                $libStatusLbl.Text = "קובץ ספרייה גרסה $zipVer קיים בתיקייה - לחץ לחילוץ"
            } else {
                $libStatusLbl.Text = "קובץ ספרייה קיים בתיקייה - לחץ לחילוץ"
            }
        } else {
            $btnLibDL.Visible = $false
            $btnLibSelectFile.Visible = $true
            $libStatusLbl.Visible = $true
            $libStatusLbl.Text = "אין חיבור לרשת - בחר קובץ ספרייה"
        }
    }
})

# Download Button
$btnDL.Add_Click({
    # Install
    if ($script:DownloadCompleted) {
        if ($script:FinalFile -and (Test-Path $script:FinalFile)) {
            $filePath = (Resolve-Path ".\$($script:FinalFile)").Path
            $statusLbl.Text = "מתקין..."
            [System.Windows.Forms.Application]::DoEvents()
            
            if ($script:FinalFile -like "*.msix") {
                if ($script:SilentInstall) {
                    # Silent MSIX Installation - no window at all
                    $statusLbl.Text = "מתקין..."
                    $btnDL.Enabled = $false
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    try {
                        # Use PowerShell hidden window to run Add-AppxPackage
                        $psi = New-Object System.Diagnostics.ProcessStartInfo
                        $psi.FileName = "powershell.exe"
                        $psi.Arguments = "-WindowStyle Hidden -Command `"Add-AppxPackage -Path '$filePath'`""
                        $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
                        $psi.CreateNoWindow = $true
                        $psi.UseShellExecute = $false
                        
                        $proc = [System.Diagnostics.Process]::Start($psi)
                        
                        # Wait for completion with UI updates
                        while (-not $proc.HasExited) {
                            [System.Windows.Forms.Application]::DoEvents()
                            Start-Sleep -Milliseconds 100
                        }
                        
                        $btnDL.Enabled = $true
                        if ($proc.ExitCode -eq 0) {
                            $statusLbl.Text = "אוצריא הותקן בהצלחה!"
                            Show-RTLMessageBox "אוצריא הותקן בהצלחה!" "התקנה הושלמה" "OK" "Information"
                        } else {
                            $statusLbl.Text = "שגיאה בהתקנה"
                        }
                    } catch {
                        $btnDL.Enabled = $true
                        $statusLbl.Text = "שגיאה: $($_.Exception.Message)"
                    }
                } else {
                    # Regular MSIX Installation - open App Installer UI
                    Start-Process $filePath
                    $statusLbl.Text = "חלון ההתקנה נפתח"
                }
            } elseif ($script:SilentInstall) {
                # Silent EXE install with existing install path if available
                $installArgs = @("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART")
                if ($script:InstallPath) {
                    $installArgs += "/DIR=`"$($script:InstallPath)`""
                }
                Start-Process -FilePath $filePath -ArgumentList $installArgs -Wait
                $statusLbl.Text = "אוצריא הותקן בהצלחה!"
            } else {
                # Regular EXE install
                Start-Process -FilePath $filePath
                $statusLbl.Text = "תוכנת ההתקנה הופעלה"
            }
        }
        return
    }
    
    # Stop
    if ($script:IsDownloading) {
        "stop" | Out-File "$env:TEMP\otzaria_stop.flag" -Force
        return
    }
    
    # Resume from paused state
    if ($script:IsPaused -and $script:TempFile -and (Test-Path $script:TempFile)) {
        $startByte = (Get-Item $script:TempFile).Length
        
        $script:IsPaused = $false
        $script:IsDownloading = $true
        $script:DownloadCompleted = $false
        $btnDL.Text = "עצור"
        $btnCancel.Visible = $true
        $statusLbl.Text = "ממשיך הורדה..."
        
        $url = $script:SelectedRelease.Url
        $tempFile = $script:TempFile
        $totalSize = $script:TotalSize
        $stopFile = "$env:TEMP\otzaria_stop.flag"
        if (Test-Path $stopFile) { Remove-Item $stopFile -Force -ErrorAction SilentlyContinue }
        
        $script:DownloadJob = Start-Job -ScriptBlock {
            param($url, $tempFile, $startByte, $totalSize, $stopFlagFile)
            try {
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                $request = [System.Net.HttpWebRequest]::Create($url)
                $request.Method = "GET"
                $request.Timeout = 60000
                $request.ReadWriteTimeout = 60000
                if ($startByte -gt 0) { $request.AddRange($startByte) }
                $response = $request.GetResponse()
                $stream = $response.GetResponseStream()
                $fileMode = if ($startByte -gt 0) { [System.IO.FileMode]::Append } else { [System.IO.FileMode]::Create }
                $fileStream = [System.IO.File]::Open($tempFile, $fileMode, [System.IO.FileAccess]::Write)
                $buffer = New-Object byte[] 65536
                $totalRead = $startByte
                $lastReport = 0
                while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $fileStream.Write($buffer, 0, $read)
                    $fileStream.Flush()
                    $totalRead += $read
                    if (Test-Path $stopFlagFile) {
                        Remove-Item $stopFlagFile -Force -ErrorAction SilentlyContinue
                        $fileStream.Close(); $stream.Close(); $response.Close()
                        Write-Output "STOPPED"
                        return
                    }
                    if (($totalRead - $lastReport) -gt 102400) {
                        $pct = if ($totalSize -gt 0) { [math]::Min([math]::Round(($totalRead / $totalSize) * 100), 100) } else { 0 }
                        Write-Output "PROGRESS:${totalRead}:${pct}"
                        $lastReport = $totalRead
                    }
                }
                $fileStream.Close(); $stream.Close(); $response.Close()
                Write-Output "DONE"
            } catch {
                Write-Output "ERROR:$($_.Exception.Message)"
            }
        } -ArgumentList $url, $tempFile, $startByte, $totalSize, $stopFile
        
        try { [TBProg]::SetState($form.Handle, 2) } catch { }
        return
    }
    
    # Select Version for new download
    $script:SelectedRelease = $null
    if ($script:InstallType -eq "MSIX") {
        if ($radioStable.Checked) { $script:SelectedRelease = $script:StableMsixRelease }
        elseif ($radioPre.Checked) { $script:SelectedRelease = $script:PreMsixRelease }
    } else {
        if ($radioStable.Checked) { $script:SelectedRelease = $script:StableRelease }
        elseif ($radioPre.Checked) { $script:SelectedRelease = $script:PreRelease }
        elseif ($radioFull.Checked) { $script:SelectedRelease = $script:FullRelease }
    }
    if (-not $script:SelectedRelease) { $statusLbl.Text = "לא נבחרה גרסה"; return }
    
    $script:FinalFile = $script:SelectedRelease.File
    
    # If file exists, mark as ready for install
    if (Test-Path $script:FinalFile) {
        $script:DownloadCompleted = $true
        $btnDL.Text = "התקן"
        if ($script:InstallType -eq "EXE") { $chkSilent.Visible = $true }
        $statusLbl.Text = "הגרסה קיימת - לחץ להתקנה"
        return
    }
    
    $script:TempFile = "$($script:FinalFile).part"
    $startByte = 0
    if (Test-Path $script:TempFile) {
        $startByte = (Get-Item $script:TempFile).Length
    }
    $script:TotalSize = Get-Size $script:SelectedRelease.Url
    
    # Check for old versions in folder and offer to delete
    $downloadingVersion = $script:SelectedRelease.FullVersion
    $downloadingParts = @($downloadingVersion -split '\.')
    
    # Get all matching files except the one being downloaded
    $currentPath = (Get-Location).Path
    
    # Force refresh directory - clear any cache
    [System.IO.Directory]::GetFiles($currentPath) | Out-Null
    
    # Get otzaria exe files
    $dirInfo = [System.IO.DirectoryInfo]::new($currentPath)
    $allExeFiles = $dirInfo.GetFiles("*.exe", [System.IO.SearchOption]::TopDirectoryOnly)
    
    # Use simple string check instead of regex
    $otzariaExeFiles = @($allExeFiles | Where-Object { $_.Name.StartsWith("otzaria-") -and $_.Name.Contains("-windows") -and $_.Name.EndsWith(".exe") })
    
    if ($script:InstallType -eq "MSIX") {
        $allMsixFiles = $dirInfo.GetFiles("*.msix", [System.IO.SearchOption]::TopDirectoryOnly)
        $allFiles = @($allMsixFiles | Where-Object { $_.Name.StartsWith("otzaria-") -and $_.Name -ne $script:FinalFile })
        $versionPattern = 'otzaria-([\d\.]+)\.msix$'
    } else {
        $allFiles = @($otzariaExeFiles | Where-Object { $_.Name -ne $script:FinalFile })
        $versionPattern = 'otzaria-([\d\.]+)-windows'
    }
    
    # Build list of old files using a simple loop and generic list
    $oldFilesList = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    
    foreach ($f in $allFiles) {
        if ($f.Name -match $versionPattern) {
            $fileVer = $matches[1]
            $fileParts = @($fileVer -split '\.')
            
            # Compare versions
            $isOlder = $false
            $maxLen = [Math]::Max($fileParts.Count, $downloadingParts.Count)
            for ($i = 0; $i -lt $maxLen; $i++) {
                $p1 = 0
                $p2 = 0
                if ($i -lt $fileParts.Count) { [int]::TryParse($fileParts[$i], [ref]$p1) | Out-Null }
                if ($i -lt $downloadingParts.Count) { [int]::TryParse($downloadingParts[$i], [ref]$p2) | Out-Null }
                
                if ($p1 -lt $p2) { $isOlder = $true; break }
                if ($p1 -gt $p2) { break }
            }
            
            if ($isOlder) {
                $oldFilesList.Add($f)
            }
        }
    }
    
    if ($oldFilesList.Count -gt 0) {
        $sb = [System.Text.StringBuilder]::new()
        [void]$sb.AppendLine("נמצאו $($oldFilesList.Count) גרסאות ישנות בתיקייה:")
        [void]$sb.AppendLine("")
        foreach ($f in $oldFilesList) {
            [void]$sb.AppendLine("• $($f.Name)")
        }
        [void]$sb.AppendLine("")
        [void]$sb.Append("האם למחוק אותן?")
        
        $r = Show-RTLMessageBox $sb.ToString() "מחיקת גרסאות ישנות" "YesNo" "Question"
        if ($r -eq "Yes") {
            foreach ($f in $oldFilesList) {
                Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue
            }
            $statusLbl.Text = "גרסאות ישנות נמחקו"
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 500
        }
    }
    
    $script:IsPaused = $false
    $script:IsDownloading = $true
    $script:DownloadCompleted = $false
    $btnDL.Text = "עצור"
    $btnCancel.Visible = $true
    $statusLbl.Text = "מתחיל הורדה..."
    
    $url = $script:SelectedRelease.Url
    $tempFile = $script:TempFile
    $totalSize = $script:TotalSize
    $stopFile = "$env:TEMP\otzaria_stop.flag"
    if (Test-Path $stopFile) { Remove-Item $stopFile -Force -ErrorAction SilentlyContinue }
    
    $script:DownloadJob = Start-Job -ScriptBlock {
        param($url, $tempFile, $startByte, $totalSize, $stopFlagFile)
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            $request = [System.Net.HttpWebRequest]::Create($url)
            $request.Method = "GET"
            $request.Timeout = 60000
            $request.ReadWriteTimeout = 60000
            if ($startByte -gt 0) { $request.AddRange($startByte) }
            $response = $request.GetResponse()
            $stream = $response.GetResponseStream()
            $fileMode = if ($startByte -gt 0) { [System.IO.FileMode]::Append } else { [System.IO.FileMode]::Create }
            $fileStream = [System.IO.File]::Open($tempFile, $fileMode, [System.IO.FileAccess]::Write)
            $buffer = New-Object byte[] 65536
            $totalRead = $startByte
            $lastReport = 0
            while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $fileStream.Write($buffer, 0, $read)
                $fileStream.Flush()
                $totalRead += $read
                if (Test-Path $stopFlagFile) {
                    Remove-Item $stopFlagFile -Force -ErrorAction SilentlyContinue
                    $fileStream.Close(); $stream.Close(); $response.Close()
                    Write-Output "STOPPED"
                    return
                }
                if (($totalRead - $lastReport) -gt 102400) {
                    $pct = if ($totalSize -gt 0) { [math]::Min([math]::Round(($totalRead / $totalSize) * 100), 100) } else { 0 }
                    Write-Output "PROGRESS:${totalRead}:${pct}"
                    $lastReport = $totalRead
                }
            }
            $fileStream.Close(); $stream.Close(); $response.Close()
            Write-Output "DONE"
        } catch {
            Write-Output "ERROR:$($_.Exception.Message)"
        }
    } -ArgumentList $url, $tempFile, $startByte, $totalSize, $stopFile
    
    try { [TBProg]::SetState($form.Handle, 2) } catch { }
})

# Cancel Button
$btnCancel.Add_Click({
    $r = Show-RTLMessageBox "האם לבטל את ההורדה?`nכל מה שהורד עד כה יימחק." "אישור ביטול" "YesNo" "Question"
    if ($r -eq "Yes") {
        "stop" | Out-File "$env:TEMP\otzaria_stop.flag" -Force
        Start-Sleep -Milliseconds 500
        if ($script:DownloadJob) {
            Stop-Job -Job $script:DownloadJob -ErrorAction SilentlyContinue
            Remove-Job -Job $script:DownloadJob -Force -ErrorAction SilentlyContinue
            $script:DownloadJob = $null
        }
        $script:IsDownloading = $false
        $script:IsPaused = $false
        if ($script:TempFile -and (Test-Path $script:TempFile)) {
            Remove-Item $script:TempFile -Force -ErrorAction SilentlyContinue
        }
        $progBar.Value = 0
        $statusLbl.Text = "ההורדה בוטלה"
        $btnDL.Text = "הורדה"
        $btnCancel.Visible = $false
        try { [TBProg]::SetState($form.Handle, 0) } catch { }
    }
})

# Library Download Button
$btnLibDL.Add_Click({
    # Extract
    if ($script:LibDownloadCompleted) {
        # Determine which file to extract
        $fileToExtract = $null
        if ($script:LibFileToExtract -and (Test-Path $script:LibFileToExtract)) {
            $fileToExtract = $script:LibFileToExtract
        } elseif (Test-Path $script:LibFinalFile) {
            $fileToExtract = $script:LibFinalFile
        } elseif (Test-Path "otzaria_latest.zip") {
            $fileToExtract = "otzaria_latest.zip"
        }
        
        if (-not $fileToExtract) {
            $libStatusLbl.Visible = $true
            $libStatusLbl.Text = "לא נמצא קובץ ספרייה"
            return
        }
        
        $libStatusLbl.Visible = $true
        
        # Use the install path found via registry
        $p = $script:InstallPath
        if (-not $p -or -not (Test-Path $p)) {
            $libStatusLbl.Text = "לא נמצאה תיקיית אוצריא"
            return
        }
        
        # Warning before deletion
        $r = Show-RTLMessageBox "שים לב: כעת יימחקו כל הספרים שבספרייה.`n`nהאם להמשיך?" "אישור מחיקה" "YesNo" "Warning"
        if ($r -ne "Yes") {
            $libStatusLbl.Text = "החילוץ בוטל"
            return
        }
        
        $btnLibDL.Visible = $false
        $libStatusLbl.Text = "מוחק קבצים ישנים..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Delete old library files and folders
        try {
            $libFolder = Join-Path $p "אוצריא"
            $linksFolder = Join-Path $p "links"
            $manifestFile = Join-Path $p "files_manifest.json"
            $metadataFile = Join-Path $p "metadata.json"
            
            if (Test-Path $libFolder) { Remove-Item $libFolder -Recurse -Force -ErrorAction SilentlyContinue }
            if (Test-Path $linksFolder) { Remove-Item $linksFolder -Recurse -Force -ErrorAction SilentlyContinue }
            if (Test-Path $manifestFile) { Remove-Item $manifestFile -Force -ErrorAction SilentlyContinue }
            if (Test-Path $metadataFile) { Remove-Item $metadataFile -Force -ErrorAction SilentlyContinue }
        } catch {
            $libStatusLbl.Text = "שגיאה במחיקת קבצים ישנים"
            $btnLibDL.Visible = $true
            return
        }
        
        $libStatusLbl.Text = "מחלץ קבצים..."
        $libProgBar.Value = 0
        $libProgBar.Visible = $true
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            $fullZipPath = (Resolve-Path $fileToExtract).Path
            
            # Use System.IO.Compression with proper encoding for Hebrew
            Add-Type -AssemblyName System.IO.Compression
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            
            # Open with IBM862 (Hebrew DOS) or UTF-8 encoding
            $encoding = [System.Text.Encoding]::GetEncoding(862)
            $zipArchive = [System.IO.Compression.ZipFile]::Open($fullZipPath, 'Read', $encoding)
            
            $entries = $zipArchive.Entries
            $totalEntries = $entries.Count
            $currentEntry = 0
            
            foreach ($entry in $entries) {
                $currentEntry++
                $pct = [math]::Round(($currentEntry / $totalEntries) * 100)
                $libProgBar.Value = $pct
                $libStatusLbl.Text = "מחלץ: $currentEntry / $totalEntries"
                [System.Windows.Forms.Application]::DoEvents()
                
                $destPath = Join-Path $p $entry.FullName
                
                if ($entry.FullName.EndsWith('/')) {
                    # Directory
                    if (-not (Test-Path $destPath)) {
                        New-Item -ItemType Directory -Path $destPath -Force | Out-Null
                    }
                } else {
                    # File
                    $destDir = Split-Path $destPath -Parent
                    if (-not (Test-Path $destDir)) {
                        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                    }
                    $entryStream = $entry.Open()
                    $fileStream = [System.IO.File]::Create($destPath)
                    $entryStream.CopyTo($fileStream)
                    $fileStream.Close()
                    $entryStream.Close()
                }
            }
            
            $zipArchive.Dispose()
            
            $libProgBar.Value = 100
            $libStatusLbl.Text = "הספרייה חולצה בהצלחה!"
            $btnLibDL.Text = "הורדת הספרייה"
            $script:LibDownloadCompleted = $false
            $script:LibFileToExtract = $null
            
            # Ask user before deleting the archive
            if (Test-Path $fileToExtract) {
                $delResult = Show-RTLMessageBox "האם ברצונך למחוק את ארכיון הספרים?" "מחיקת ארכיון" "YesNo" "Question"
                if ($delResult -eq "Yes") {
                    Remove-Item $fileToExtract -Force -ErrorAction SilentlyContinue
                    $libStatusLbl.Text = "הספרייה חולצה והארכיון נמחק!"
                }
            }
        } catch {
            $libStatusLbl.Text = "שגיאה בחילוץ: $_"
        }
        $libProgBar.Value = 0
        $btnLibDL.Visible = $true
        return
    }
    
    # Stop
    if ($script:LibIsDownloading) {
        "stop" | Out-File "$env:TEMP\otzaria_lateststop.flag" -Force
        return
    }
    
    # Resume from paused state
    if ($script:LibIsPaused -and $script:LibTempFile -and (Test-Path $script:LibTempFile)) {
        $startByte = (Get-Item $script:LibTempFile).Length
        
        $script:LibIsPaused = $false
        $script:LibIsDownloading = $true
        $script:LibDownloadCompleted = $false
        $btnLibDL.Text = "עצור"
        $btnLibCancel.Visible = $true
        $libStatusLbl.Visible = $true
        $libStatusLbl.Text = "ממשיך הורדה..."
        
        $url = $script:LibUrl
        $tempFile = $script:LibTempFile
        $totalSize = $script:LibTotalSize
        $stopFile = "$env:TEMP\otzaria_lateststop.flag"
        if (Test-Path $stopFile) { Remove-Item $stopFile -Force -ErrorAction SilentlyContinue }
        
        $script:LibDownloadJob = Start-Job -ScriptBlock {
            param($url, $tempFile, $startByte, $totalSize, $stopFlagFile)
            try {
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                $request = [System.Net.HttpWebRequest]::Create($url)
                $request.Method = "GET"
                $request.Timeout = 60000
                $request.ReadWriteTimeout = 60000
                if ($startByte -gt 0) { $request.AddRange($startByte) }
                $response = $request.GetResponse()
                $stream = $response.GetResponseStream()
                $fileMode = if ($startByte -gt 0) { [System.IO.FileMode]::Append } else { [System.IO.FileMode]::Create }
                $fileStream = [System.IO.File]::Open($tempFile, $fileMode, [System.IO.FileAccess]::Write)
                $buffer = New-Object byte[] 65536
                $totalRead = $startByte
                $lastReport = 0
                while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $fileStream.Write($buffer, 0, $read)
                    $fileStream.Flush()
                    $totalRead += $read
                    if (Test-Path $stopFlagFile) {
                        Remove-Item $stopFlagFile -Force -ErrorAction SilentlyContinue
                        $fileStream.Close(); $stream.Close(); $response.Close()
                        Write-Output "STOPPED"
                        return
                    }
                    if (($totalRead - $lastReport) -gt 102400) {
                        $pct = if ($totalSize -gt 0) { [math]::Min([math]::Round(($totalRead / $totalSize) * 100), 100) } else { 0 }
                        Write-Output "PROGRESS:${totalRead}:${pct}"
                        $lastReport = $totalRead
                    }
                }
                $fileStream.Close(); $stream.Close(); $response.Close()
                Write-Output "DONE"
            } catch {
                Write-Output "ERROR:$($_.Exception.Message)"
            }
        } -ArgumentList $url, $tempFile, $startByte, $totalSize, $stopFile
        
        try { [TBProg]::SetState($form.Handle, 2) } catch { }
        return
    }
    
    # Check if file already exists (versioned or latest)
    $existingFile = $null
    if (Test-Path $script:LibFinalFile) {
        $existingFile = $script:LibFinalFile
    } elseif (Test-Path "otzaria_latest.zip") {
        $existingFile = "otzaria_latest.zip"
    }
    
    if ($existingFile) {
        $script:LibDownloadCompleted = $true
        $script:LibFileToExtract = $existingFile
        $btnLibDL.Text = "חלץ לאוצריא"
        $libStatusLbl.Visible = $true
        $libStatusLbl.Text = "קובץ ספרייה קיים - לחץ לחילוץ"
        return
    }
    
    # New download
    $startByte = 0
    if (Test-Path $script:LibTempFile) {
        $startByte = (Get-Item $script:LibTempFile).Length
    }
    $script:LibTotalSize = Get-Size $script:LibUrl
    
    $script:LibIsPaused = $false
    $script:LibIsDownloading = $true
    $script:LibDownloadCompleted = $false
    $btnLibDL.Text = "עצור"
    $btnLibCancel.Visible = $true
    $libStatusLbl.Visible = $true
    $libStatusLbl.Text = "מתחיל הורדה..."
    
    $url = $script:LibUrl
    $tempFile = $script:LibTempFile
    $totalSize = $script:LibTotalSize
    $stopFile = "$env:TEMP\otzaria_lateststop.flag"
    if (Test-Path $stopFile) { Remove-Item $stopFile -Force -ErrorAction SilentlyContinue }
    
    $script:LibDownloadJob = Start-Job -ScriptBlock {
        param($url, $tempFile, $startByte, $totalSize, $stopFlagFile)
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            $request = [System.Net.HttpWebRequest]::Create($url)
            $request.Method = "GET"
            $request.Timeout = 60000
            $request.ReadWriteTimeout = 60000
            if ($startByte -gt 0) { $request.AddRange($startByte) }
            $response = $request.GetResponse()
            $stream = $response.GetResponseStream()
            $fileMode = if ($startByte -gt 0) { [System.IO.FileMode]::Append } else { [System.IO.FileMode]::Create }
            $fileStream = [System.IO.File]::Open($tempFile, $fileMode, [System.IO.FileAccess]::Write)
            $buffer = New-Object byte[] 65536
            $totalRead = $startByte
            $lastReport = 0
            while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $fileStream.Write($buffer, 0, $read)
                $fileStream.Flush()
                $totalRead += $read
                if (Test-Path $stopFlagFile) {
                    Remove-Item $stopFlagFile -Force -ErrorAction SilentlyContinue
                    $fileStream.Close(); $stream.Close(); $response.Close()
                    Write-Output "STOPPED"
                    return
                }
                if (($totalRead - $lastReport) -gt 102400) {
                    $pct = if ($totalSize -gt 0) { [math]::Min([math]::Round(($totalRead / $totalSize) * 100), 100) } else { 0 }
                    Write-Output "PROGRESS:${totalRead}:${pct}"
                    $lastReport = $totalRead
                }
            }
            $fileStream.Close(); $stream.Close(); $response.Close()
            Write-Output "DONE"
        } catch {
            Write-Output "ERROR:$($_.Exception.Message)"
        }
    } -ArgumentList $url, $tempFile, $startByte, $totalSize, $stopFile
    
    try { [TBProg]::SetState($form.Handle, 2) } catch { }
})

# Library Cancel Button
$btnLibCancel.Add_Click({
    $r = Show-RTLMessageBox "האם לבטל את ההורדה?`nכל מה שהורד עד כה יימחק." "אישור ביטול" "YesNo" "Question"
    if ($r -eq "Yes") {
        "stop" | Out-File "$env:TEMP\otzaria_lateststop.flag" -Force
        Start-Sleep -Milliseconds 500
        if ($script:LibDownloadJob) {
            Stop-Job -Job $script:LibDownloadJob -ErrorAction SilentlyContinue
            Remove-Job -Job $script:LibDownloadJob -Force -ErrorAction SilentlyContinue
            $script:LibDownloadJob = $null
        }
        $script:LibIsDownloading = $false
        $script:LibIsPaused = $false
        if ($script:LibTempFile -and (Test-Path $script:LibTempFile)) {
            Remove-Item $script:LibTempFile -Force -ErrorAction SilentlyContinue
        }
        $libProgBar.Value = 0
        $libStatusLbl.Text = "ההורדה בוטלה"
        $btnLibDL.Text = "הורדת הספרייה"
        $btnLibCancel.Visible = $false
        try { [TBProg]::SetState($form.Handle, 0) } catch { }
    }
})

# Library Select File Button (offline mode)
$btnLibSelectFile.Add_Click({
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Title = "בחר קובץ ספרייה"
    $openDialog.Filter = "קבצי ZIP (*.zip)|*.zip|כל הקבצים (*.*)|*.*"
    $openDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($openDialog.ShowDialog() -eq "OK") {
        $selectedFile = $openDialog.FileName
        $fileName = [System.IO.Path]::GetFileName($selectedFile)
        
        # Copy to current directory if not already there
        $targetPath = Join-Path (Get-Location) $fileName
        if ($selectedFile -ne $targetPath) {
            $libStatusLbl.Visible = $true
            $libStatusLbl.Text = "מעתיק קובץ..."
            [System.Windows.Forms.Application]::DoEvents()
            Copy-Item $selectedFile $targetPath -Force
        }
        
        # Get version from zip file
        $zipVer = Get-ZipLibVersion $fileName
        if ($zipVer) {
            $lblLibV.Text = "גרסה בקובץ: $zipVer"
        } else {
            $lblLibV.Text = "גרסה: לא ידוע"
        }
        
        $script:LibDownloadCompleted = $true
        $script:LibFileToExtract = $fileName
        $btnLibDL.Text = "חלץ לאוצריא"
        $btnLibDL.Visible = $true
        $btnLibSelectFile.Visible = $false
        $libStatusLbl.Visible = $true
        $libStatusLbl.Text = "קובץ נבחר: $fileName - לחץ לחילוץ"
    }
})

# Library Extract Button (for extracting existing file when newer version available)
$btnLibExtract.Add_Click({
    # Extract existing library file
    $fileToExtract = $script:LibFileToExtract
    if (-not $fileToExtract -or -not (Test-Path $fileToExtract)) {
        $libStatusLbl.Visible = $true
        $libStatusLbl.Text = "לא נמצא קובץ ספרייה לחילוץ"
        return
    }
    
    $libStatusLbl.Visible = $true
    
    # Use the install path found via registry
    $p = $script:InstallPath
    if (-not $p -or -not (Test-Path $p)) {
        $libStatusLbl.Text = "לא נמצאה תיקיית אוצריא"
        return
    }
    
    # Warning before deletion
    $r = Show-RTLMessageBox "שים לב: כעת יימחקו כל הספרים שבספרייה.`n`nהאם להמשיך?" "אישור מחיקה" "YesNo" "Warning"
    if ($r -ne "Yes") {
        $libStatusLbl.Text = "החילוץ בוטל"
        return
    }
    
    $btnLibExtract.Visible = $false
    $btnLibDL.Visible = $false
    $libStatusLbl.Text = "מוחק קבצים ישנים..."
    [System.Windows.Forms.Application]::DoEvents()
    
    # Delete old library files and folders
    try {
        $libFolder = Join-Path $p "אוצריא"
        $linksFolder = Join-Path $p "links"
        $manifestFile = Join-Path $p "files_manifest.json"
        $metadataFile = Join-Path $p "metadata.json"
        
        if (Test-Path $libFolder) { Remove-Item $libFolder -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path $linksFolder) { Remove-Item $linksFolder -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path $manifestFile) { Remove-Item $manifestFile -Force -ErrorAction SilentlyContinue }
        if (Test-Path $metadataFile) { Remove-Item $metadataFile -Force -ErrorAction SilentlyContinue }
    } catch {
        $libStatusLbl.Text = "שגיאה במחיקת קבצים ישנים"
        $btnLibExtract.Visible = $true
        $btnLibDL.Visible = $true
        return
    }
    
    $libStatusLbl.Text = "מחלץ קבצים..."
    $libProgBar.Value = 0
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $encoding = [System.Text.Encoding]::GetEncoding(862)
        $zipPath = (Resolve-Path ".\$fileToExtract").Path
        $archive = [System.IO.Compression.ZipFile]::Open($zipPath, 'Read', $encoding)
        $totalEntries = $archive.Entries.Count
        $currentEntry = 0
        
        foreach ($entry in $archive.Entries) {
            $currentEntry++
            $pct = [math]::Round(($currentEntry / $totalEntries) * 100)
            $libProgBar.Value = $pct
            if ($currentEntry % 100 -eq 0) {
                $libStatusLbl.Text = "מחלץ: $currentEntry / $totalEntries"
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            $targetPath = Join-Path $p $entry.FullName
            $targetDir = [System.IO.Path]::GetDirectoryName($targetPath)
            if (-not (Test-Path $targetDir)) {
                [System.IO.Directory]::CreateDirectory($targetDir) | Out-Null
            }
            if ($entry.Name -ne "") {
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $targetPath, $true)
            }
        }
        $archive.Dispose()
        
        $libProgBar.Value = 100
        $libStatusLbl.Text = "!הספרייה חולצה בהצלחה"
        
        # Ask to delete zip file
        $delR = Show-RTLMessageBox "האם למחוק את קובץ הארכיון?" "מחיקת קובץ" "YesNo" "Question"
        if ($delR -eq "Yes") {
            Remove-Item $fileToExtract -Force -ErrorAction SilentlyContinue
            $libStatusLbl.Text = "הספרייה חולצה והארכיון נמחק"
        }
    } catch {
        $libStatusLbl.Text = "שגיאה בחילוץ: $($_.Exception.Message)"
    }
    $libProgBar.Value = 0
    $btnLibDL.Visible = $true
    # Don't show extract button anymore after successful extraction
})

# Clear Cache Button
$btnClearCache.Add_Click({
    $r = Show-RTLMessageBox "האם אתה בטוח שברצונך למחוק את המטמון של תוכנת אוצריא?`n`nשים לב: כל ההיסטוריה, המועדפים וההעדפות יימחקו ולא יהיה ניתן לשחזרם!" "מחיקת מטמון" "YesNo" "Warning"
    if ($r -eq "Yes") {
        try { Start-Process -FilePath "taskkill" -ArgumentList "/F /IM otzaria.exe" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue } catch { }
        Start-Sleep -Milliseconds 500
        $cachePath = Join-Path $env:APPDATA "com.example"
        if (Test-Path $cachePath) {
            try {
                Remove-Item $cachePath -Recurse -Force -ErrorAction Stop
                Show-RTLMessageBox "המטמון נמחק בהצלחה!" "הצלחה" "OK" "Information"
            } catch {
                Show-RTLMessageBox "שגיאה במחיקת המטמון: $_" "שגיאה" "OK" "Error"
            }
        } else {
            Show-RTLMessageBox "תיקיית המטמון לא נמצאה" "מידע" "OK" "Information"
        }
    }
})

[void]$form.ShowDialog()
