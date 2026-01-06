Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Set working directory to script location
try {
    $scriptPath = $MyInvocation.MyCommand.Path
    if ($scriptPath) {
        $scriptDir = Split-Path -Parent $scriptPath
        if ($scriptDir -and (Test-Path $scriptDir)) {
            Set-Location $scriptDir
        }
    }
} catch { }

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

# DWM API for title bar color (Windows 10/11)
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DwmHelper {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    
    public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
    public const int DWMWA_CAPTION_COLOR = 35;
    
    public static void SetTitleBarColor(IntPtr hwnd, int color) {
        try {
            DwmSetWindowAttribute(hwnd, DWMWA_CAPTION_COLOR, ref color, sizeof(int));
        } catch { }
    }
    
    public static void SetDarkMode(IntPtr hwnd, bool dark) {
        try {
            int value = dark ? 1 : 0;
            DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
        } catch { }
    }
}
"@ -ErrorAction SilentlyContinue
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
# Detect Windows dark mode from registry
$script:DarkMode = $false
$script:DarkModeManual = $false  # Track if user manually changed theme
try {
    $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    $appsUseLightTheme = Get-ItemPropertyValue -Path $regPath -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue
    if ($appsUseLightTheme -eq 0) {
        $script:DarkMode = $true
    }
} catch {
    $script:DarkMode = $false
}

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

# Differential Update Variables
$script:DiffUpdateAvailable = $false
$script:DiffUpdateJob = $null
$script:DiffIsUpdating = $false
$script:DiffFilesToUpdate = @()
$script:DiffFilesToDelete = @()
$script:DiffCurrentFile = 0
$script:DiffTotalFiles = 0
$script:LocalManifest = $null
$script:RemoteManifest = $null
$script:RemoteManifestUrl = "https://raw.githubusercontent.com/Otzaria/otzaria-library/refs/heads/main/files_manifest.json"
$script:GitHubRawBaseUrl = "https://raw.githubusercontent.com/Otzaria/otzaria-library/refs/heads/main/"
$script:LibraryPathPrefix = "sefariaToOtzaria/sefaria_export/ספרים/"  # Prefix to remove from manifest paths
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
    
    # Final fallback: check for אוצריא folder in the current working directory
    $currentDir = (Get-Location).Path
    $localOtzaria = Join-Path $currentDir "אוצריא"
    $localExe = Join-Path $localOtzaria "otzaria.exe"
    if (Test-Path $localExe) {
        return $localOtzaria
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
        # Get up to 100 releases per page, check multiple pages if needed
        $allReleases = @()
        for ($page = 1; $page -le 4; $page++) {
            try {
                $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases?per_page=100&page=$page" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
                if ($r.Count -eq 0) { break }
                $allReleases += $r
            } catch { break }
        }
        
        $rel = $allReleases | Where-Object { $_.prerelease -eq $Pre } | Select-Object -First 1
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
        # Get up to 100 releases per page, check multiple pages if needed
        $allReleases = @()
        for ($page = 1; $page -le 4; $page++) {
            try {
                $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases?per_page=100&page=$page" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
                if ($r.Count -eq 0) { break }
                $allReleases += $r
            } catch { break }
        }
        
        foreach ($rel in $allReleases) {
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
        # Get up to 100 releases per page, check multiple pages if needed
        $allReleases = @()
        for ($page = 1; $page -le 4; $page++) {
            try {
                $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases?per_page=100&page=$page" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
                if ($r.Count -eq 0) { break }
                $allReleases += $r
            } catch { break }
        }
        
        $rel = $allReleases | Where-Object { $_.prerelease -eq $Pre } | Select-Object -First 1
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
        $wc.Headers.Add("User-Agent", "OtzariaUpdater")
        $wc.Headers.Add("Cache-Control", "no-cache")
        $timestamp = [DateTimeOffset]::Now.ToUnixTimeSeconds()
        return $wc.DownloadString("https://raw.githubusercontent.com/Otzaria/otzaria-library/refs/heads/main/MoreBooks/%D7%A1%D7%A4%D7%A8%D7%99%D7%9D/%D7%90%D7%95%D6%B9%D7%A6%D7%A8%D7%99%D7%90/%D7%90%D7%95%D7%93%D7%95%D7%AA%20%D7%94%D7%AA%D7%95%D7%9B%D7%A0%D7%94/%D7%92%D7%99%D7%A8%D7%A1%D7%AA%20%D7%A1%D7%A4%D7%A8%D7%99%D7%94.txt?t=$timestamp").Trim()
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

function Get-ManifestFromZip {
    param($zipPath)
    try {
        if (-not (Test-Path $zipPath)) { return $null }
        
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        
        $encoding = [System.Text.Encoding]::GetEncoding(862)
        $zipArchive = [System.IO.Compression.ZipFile]::Open($zipPath, 'Read', $encoding)
        
        # Look for files_manifest.json inside the ZIP
        $manifestEntry = $zipArchive.Entries | Where-Object { 
            $_.FullName -eq "files_manifest.json" -or 
            $_.FullName -like "*/files_manifest.json"
        } | Select-Object -First 1
        
        if ($manifestEntry) {
            $stream = $manifestEntry.Open()
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            $jsonContent = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()
            $zipArchive.Dispose()
            
            return $jsonContent | ConvertFrom-Json
        }
        
        $zipArchive.Dispose()
        return $null
    } catch {
        return $null
    }
}

function Get-RemoteManifest {
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        $wc = New-Object System.Net.WebClient
        $wc.Encoding = [System.Text.Encoding]::UTF8
        $wc.Headers.Add("User-Agent", "OtzariaUpdater")
        $jsonContent = $wc.DownloadString($script:RemoteManifestUrl)
        return $jsonContent | ConvertFrom-Json
    } catch {
        return $null
    }
}

# Convert manifest path to ZIP path
# manifest: "sefariaToOtzaria/sefaria_export/ספרים/אוצריא/..." -> ZIP: "אוצריא/..."
# manifest: "sefariaToOtzaria/sefaria_export/links/..." -> ZIP: "links/..."
# manifest: "metadata.json" -> ZIP: "metadata.json"
# manifest: "files_manifest.json" -> ZIP: "files_manifest.json"
function Convert-ManifestPathToZipPath {
    param($manifestPath)
    
    if ($manifestPath -eq "metadata.json" -or $manifestPath -eq "files_manifest.json") {
        return $manifestPath
    }
    
    # Handle אוצריא files: sefariaToOtzaria/sefaria_export/ספרים/אוצריא/... -> אוצריא/...
    if ($manifestPath -like "sefariaToOtzaria/sefaria_export/ספרים/*") {
        return $manifestPath.Substring("sefariaToOtzaria/sefaria_export/ספרים/".Length)
    }
    
    # Handle links files: sefariaToOtzaria/sefaria_export/links/... -> links/...
    if ($manifestPath -like "sefariaToOtzaria/sefaria_export/links/*") {
        return $manifestPath.Substring("sefariaToOtzaria/sefaria_export/".Length)
    }
    
    # Default - return as is
    return $manifestPath
}

function Compare-LibraryManifests {
    param($localManifest, $remoteManifest)
    
    $changes = @{
        New = @()      # קבצים חדשים
        Modified = @() # קבצים שהשתנו
        Deleted = @()  # קבצים שנמחקו
    }
    
    if (-not $localManifest -or -not $remoteManifest) {
        return $changes
    }
    
    # Convert to hashtable for easier lookup
    $localHash = @{}
    foreach ($prop in $localManifest.PSObject.Properties) {
        $localHash[$prop.Name] = $prop.Value.hash
    }
    
    $remoteHash = @{}
    foreach ($prop in $remoteManifest.PSObject.Properties) {
        $remoteHash[$prop.Name] = $prop.Value.hash
    }
    
    # Find new and modified files
    foreach ($file in $remoteHash.Keys) {
        if (-not $localHash.ContainsKey($file)) {
            $changes.New += $file
        }
        elseif ($localHash[$file] -ne $remoteHash[$file]) {
            $changes.Modified += $file
        }
    }
    
    # Find deleted files
    foreach ($file in $localHash.Keys) {
        if (-not $remoteHash.ContainsKey($file)) {
            $changes.Deleted += $file
        }
    }
    
    return $changes
}

function Check-LibraryUpdateAvailable {
    # Only check ZIP file - not installed folder
    $localManifest = $null
    $zipFile = $null
    
    # Find ZIP file
    if (Test-Path "otzaria_latest.zip") {
        $zipFile = "otzaria_latest.zip"
    } else {
        $zipFiles = Get-ChildItem "." -Filter "*otzaria_*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
        if ($zipFiles) { $zipFile = $zipFiles[0].Name }
    }
    
    # If no ZIP file exists, can't do differential update
    if (-not $zipFile -or -not (Test-Path $zipFile)) {
        return @{ Available = $false; ZipFile = $null; Changes = $null; NoZipFile = $true }
    }
    
    $localManifest = Get-ManifestFromZip $zipFile
    
    # Get remote manifest first
    $script:RemoteManifest = Get-RemoteManifest
    if (-not $script:RemoteManifest) {
        return @{ Available = $false; ZipFile = $zipFile; Changes = $null; NoNetwork = $true }
    }
    
    # If no local manifest in ZIP, we need full update (all remote files are "new")
    if (-not $localManifest) {
        $allFiles = @()
        foreach ($prop in $script:RemoteManifest.PSObject.Properties) {
            $allFiles += $prop.Name
        }
        return @{
            Available = $true
            ZipFile = $zipFile
            Changes = @{ New = $allFiles; Modified = @(); Deleted = @() }
            NewCount = $allFiles.Count
            ModifiedCount = 0
            DeletedCount = 0
            TotalCount = $allFiles.Count
            NoLocalManifest = $true
        }
    }
    
    $script:LocalManifest = $localManifest
    
    # Compare manifests
    $changes = Compare-LibraryManifests $script:LocalManifest $script:RemoteManifest
    
    $totalChanges = $changes.New.Count + $changes.Modified.Count + $changes.Deleted.Count
    
    return @{
        Available = ($totalChanges -gt 0)
        ZipFile = $zipFile
        Changes = $changes
        NewCount = $changes.New.Count
        ModifiedCount = $changes.Modified.Count
        DeletedCount = $changes.Deleted.Count
        TotalCount = $totalChanges
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
$form.MinimumSize = New-Object System.Drawing.Size(720,860)
$form.Size = New-Object System.Drawing.Size(720,860)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $true
$form.RightToLeft = "Yes"
$form.RightToLeftLayout = $true
$form.FormBorderStyle = "Sizable"
$form.Font = New-Object System.Drawing.Font("Segoe UI",10)

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

# Differential Update Button
$btnLibUpdate = New-Object System.Windows.Forms.Button
$btnLibUpdate.Text = "עדכון ארכיון"
$btnLibUpdate.Location = New-Object System.Drawing.Point(417,72)
$btnLibUpdate.Size = New-Object System.Drawing.Size(130,34)
$btnLibUpdate.FlatStyle = "Flat"
$btnLibUpdate.Visible = $false
$libCard.Controls.Add($btnLibUpdate)
$toolTip.SetToolTip($btnLibUpdate, "עדכון דיפרנציאלי - הורדת רק קבצים שהשתנו")

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

# Function to check if VC++ Redistributable is installed
function Test-VCRedistInstalled {
    $vcKeys = @(
        "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64",
        "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x86",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x86"
    )
    foreach ($key in $vcKeys) {
        if (Test-Path $key) {
            $installed = Get-ItemProperty -Path $key -Name "Installed" -ErrorAction SilentlyContinue
            if ($installed -and $installed.Installed -eq 1) {
                return $true
            }
        }
    }
    # Also check for older versions
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($path in $regPaths) {
        $items = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Visual C++*Redistributable*" }
        if ($items) { return $true }
    }
    return $false
}

# Bottom Buttons
$btnVCRedist = New-Object System.Windows.Forms.Button
$btnVCRedist.Text = "התקנת הרחבה"
$btnVCRedist.Size = New-Object System.Drawing.Size(180,40)
$btnVCRedist.Location = New-Object System.Drawing.Point(490,685)
$btnVCRedist.FlatStyle = "Flat"
$btnVCRedist.Font = New-Object System.Drawing.Font("Segoe UI Semibold",11)
$btnVCRedist.Visible = $false
$toolTip.SetToolTip($btnVCRedist, "התקנת Visual C++ Redistributable")

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

$form.Controls.AddRange(@($header,$mainCard,$libCard,$btnClearCache,$btnVCRedist,$statusDivider,$statusLocation,$statusBar,$lblNetwork))

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
    foreach ($btn in @($btnDL, $btnCancel, $btnPreChangelog, $btnLibDL, $btnLibCancel, $btnLibChangelog, $btnLibExtract, $btnLibSelectFile, $btnClearCache, $btnVCRedist, $btnTheme, $btnRefresh, $btnSelectFile, $btnLibUpdate)) {
        $btn.BackColor = $t.BtnBg
        $btn.ForeColor = $t.TextColor
        $btn.FlatAppearance.BorderColor = $t.BtnBorder
    }
    
    # Set title bar color (Windows 10/11)
    try {
        if ($form.Handle) {
            # Set dark mode for title bar buttons
            [DwmHelper]::SetDarkMode($form.Handle, $script:DarkMode)
            
            # Set title bar caption color (COLORREF format: 0x00BBGGRR)
            if ($script:DarkMode) {
                # Dark mode: use header color (40,36,34 -> #282422)
                $titleBarColor = 0x00222428
            } else {
                # Light mode: use #fff8f4
                $titleBarColor = 0x00f4f8ff
            }
            [DwmHelper]::SetTitleBarColor($form.Handle, $titleBarColor)
        }
    } catch { }
    
    $form.Refresh()
}

# Timer
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 200

# Variable to track last check time for dark mode
$script:LastDarkModeCheck = [DateTime]::MinValue

$timer.Add_Tick({
    # Check Windows dark mode every 2 seconds (only if not manually overridden)
    if (-not $script:DarkModeManual -and ([DateTime]::Now - $script:LastDarkModeCheck).TotalSeconds -ge 2) {
        $script:LastDarkModeCheck = [DateTime]::Now
        try {
            $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            $appsUseLightTheme = Get-ItemPropertyValue -Path $regPath -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue
            $windowsDarkMode = ($appsUseLightTheme -eq 0)
            if ($windowsDarkMode -ne $script:DarkMode) {
                $script:DarkMode = $windowsDarkMode
                Apply-Theme
                # Update theme button tooltip
                if ($script:DarkMode) {
                    $btnTheme.Text = "◐"
                    $toolTip.SetToolTip($btnTheme, "מצב בהיר")
                } else {
                    $btnTheme.Text = "◐"
                    $toolTip.SetToolTip($btnTheme, "מצב כהה")
                }
            }
        } catch { }
    }
    
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
                $libProgBar.Value = 100
                $libStatusLbl.Text = "ההורדה הושלמה! לחץ 'חלץ לאוצריא'"
                $btnLibDL.Visible = $false
                $btnLibExtract.Visible = $true
                $btnLibExtract.Text = "חלץ לאוצריא"
                $btnLibExtract.Size = New-Object System.Drawing.Size(180,34)
                $btnLibExtract.Location = New-Object System.Drawing.Point(237,72)
                $btnLibCancel.Visible = $false
                $script:LibDownloadCompleted = $true
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
    
    # VC++ Redistributable Download
    if ($script:VCDownloadJob -and $script:VCDownloading) {
        $output = Receive-Job -Job $script:VCDownloadJob -ErrorAction SilentlyContinue
        foreach ($line in $output) {
            if ($line -match '^SIZE:(\d+)$') {
                $script:VCTotalSize = [long]$matches[1]
            }
            elseif ($line -match '^PROGRESS:(\d+):(\d+)$') {
                $bytes = [long]$matches[1]; $pct = [int]$matches[2]
                $libProgBar.Value = $pct
                $libProgBar.Refresh()
                $mbDone = [math]::Round($bytes / 1MB, 1); $mbTotal = [math]::Round($script:VCTotalSize / 1MB, 1)
                $libStatusLbl.Text = "הורד $mbDone מתוך $mbTotal מ`"ב :VC++"
                try { [TBProg]::SetVal($form.Handle, $pct, 100) } catch { }
            }
            elseif ($line -eq "DONE") {
                $script:VCDownloading = $false
                $libProgBar.Value = 100
                $libStatusLbl.Text = "מתקין VC++ AIO..."
                [System.Windows.Forms.Application]::DoEvents()
                
                # Install - AIO uses /ai for silent install (auto install all)
                $proc = Start-Process -FilePath $script:VCFile -ArgumentList "/ai" -Wait -PassThru
                
                $libProgBar.Value = 0
                $libStatusLbl.Visible = $false
                $btnVCRedist.Text = "התקנת הרחבה"
                
                if ($proc.ExitCode -eq 0) {
                    Show-RTLMessageBox "ההתקנה הושלמה בהצלחה!" "הצלחה" "OK" "Information"
                    $btnVCRedist.Visible = $false
                    $btnClearCache.Location = New-Object System.Drawing.Point(235,685)
                } else {
                    Show-RTLMessageBox "ההתקנה נכשלה. קוד שגיאה: $($proc.ExitCode)" "שגיאה" "OK" "Error"
                }
                
                # Delete installer file from TEMP
                Remove-Item $script:VCFile -Force -ErrorAction SilentlyContinue
                
                try { [TBProg]::SetState($form.Handle, 0) } catch { }
                Remove-Job -Job $script:VCDownloadJob -Force -ErrorAction SilentlyContinue; $script:VCDownloadJob = $null
            }
            elseif ($line -eq "STOPPED") {
                $script:VCDownloading = $false
                $libProgBar.Value = 0
                $libStatusLbl.Visible = $false
                $btnVCRedist.Text = "התקנת הרחבה"
                try { [TBProg]::SetState($form.Handle, 0) } catch { }
                Remove-Job -Job $script:VCDownloadJob -Force -ErrorAction SilentlyContinue; $script:VCDownloadJob = $null
            }
            elseif ($line -match '^ERROR:(.*)$') {
                $script:VCDownloading = $false
                $libProgBar.Value = 0
                $libStatusLbl.Visible = $false
                $btnVCRedist.Text = "התקנת הרחבה"
                Show-RTLMessageBox "שגיאה בהורדה: $($matches[1])" "שגיאה" "OK" "Error"
                try { [TBProg]::SetState($form.Handle, 0) } catch { }
                Remove-Job -Job $script:VCDownloadJob -Force -ErrorAction SilentlyContinue; $script:VCDownloadJob = $null
            }
        }
    }
    
    # Differential Update Progress
    if ($script:DiffUpdateJob -and $script:DiffIsUpdating) {
        $output = Receive-Job -Job $script:DiffUpdateJob -ErrorAction SilentlyContinue
        foreach ($line in $output) {
            if ($line -match '^PROGRESS:(\d+):(\d+):(.*)$') {
                $current = [int]$matches[1]
                $total = [int]$matches[2]
                $fileName = $matches[3]
                $pct = [math]::Round(($current / $total) * 100)
                $libProgBar.Value = $pct
                $libProgBar.Refresh()
                $libStatusLbl.Text = "מעדכן: $current / $total"
                try { [TBProg]::SetVal($form.Handle, $pct, 100) } catch { }
            }
            elseif ($line -eq "DONE") {
                $script:DiffIsUpdating = $false
                $libProgBar.Value = 100
                $libStatusLbl.Text = "העדכון הושלם בהצלחה!"
                $btnLibUpdate.Visible = $false
                $btnLibDL.Visible = $false
                $script:DiffUpdateAvailable = $false
                try { [TBProg]::SetState($form.Handle, 0) } catch { }
                Remove-Job -Job $script:DiffUpdateJob -Force -ErrorAction SilentlyContinue
                $script:DiffUpdateJob = $null
                
                # Get the updated ZIP file and rename it according to the new version
                $zipFile = Get-ChildItem "." -Filter "*otzaria_*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if (-not $zipFile) { $zipFile = Get-Item "otzaria_latest.zip" -ErrorAction SilentlyContinue }
                if ($zipFile) {
                    $newVer = Get-ZipLibVersion $zipFile.FullName
                    if ($newVer) {
                        # Rename the file to match the new version
                        $newFileName = "otzaria_latest_$newVer.zip"
                        if ($zipFile.Name -ne $newFileName) {
                            try {
                                $newPath = Join-Path $zipFile.DirectoryName $newFileName
                                # Remove existing file with same name if exists
                                if (Test-Path $newPath) {
                                    Remove-Item $newPath -Force
                                }
                                Rename-Item $zipFile.FullName $newFileName -Force
                                $libStatusLbl.Text = "העדכון הושלם! הקובץ שונה ל-$newFileName"
                            } catch {
                                $libStatusLbl.Text = "העדכון הושלם בהצלחה!"
                            }
                        }
                        # Update displayed version
                        $lblLibV.Text = "גרסה בארכיון: $newVer   |   גרסה בשרת: $($script:LibVer)"
                    }
                    
                    # Show extract button centered
                    $btnLibExtract.Text = "חלץ לאוצריא"
                    $btnLibExtract.Size = New-Object System.Drawing.Size(180,34)
                    $btnLibExtract.Location = New-Object System.Drawing.Point(237,72)
                    $btnLibExtract.Visible = $true
                }
                
                # Refresh the UI after a short delay
                Start-Sleep -Milliseconds 500
                $btnRefresh.PerformClick()
            }
            elseif ($line -eq "STOPPED") {
                $script:DiffIsUpdating = $false
                $libProgBar.Value = 0
                $libStatusLbl.Text = "העדכון הופסק"
                $btnLibUpdate.Text = "עדכון ארכיון"
                $btnLibUpdate.Visible = $true
                $btnLibDL.Visible = $true
                try { [TBProg]::SetState($form.Handle, 0) } catch { }
                Remove-Job -Job $script:DiffUpdateJob -Force -ErrorAction SilentlyContinue
                $script:DiffUpdateJob = $null
            }
            elseif ($line -match '^ERROR:(.*)$') {
                $script:DiffIsUpdating = $false
                $libStatusLbl.Text = "שגיאה: $($matches[1])"
                $btnLibUpdate.Text = "עדכון ארכיון"
                $btnLibUpdate.Visible = $true
                $btnLibDL.Visible = $true
                $libProgBar.Value = 0
                try { [TBProg]::SetState($form.Handle, 0) } catch { }
                Remove-Job -Job $script:DiffUpdateJob -Force -ErrorAction SilentlyContinue
                $script:DiffUpdateJob = $null
            }
            elseif ($line -match '^WARN:(.*)$') {
                # Log warning but continue - don't stop the update
                # Could add a warning counter here if needed
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
    
    # Check if VC++ Redistributable is installed
    if (-not (Test-VCRedistInstalled)) {
        $btnVCRedist.Visible = $true
        # Center both buttons - total width = 250 (cache) + 20 (gap) + 180 (VC) = 450
        # Form width is 700, so startX = (700 - 450) / 2 = 125
        $totalWidth = $btnClearCache.Width + 20 + $btnVCRedist.Width
        $startX = [int]((700 - $totalWidth) / 2)
        $btnClearCache.Location = New-Object System.Drawing.Point($startX, 685)
        $btnVCRedist.Location = New-Object System.Drawing.Point(($startX + $btnClearCache.Width + 20), 685)
    }
    
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
    
    # Get main releases - fetch multiple pages
    try {
        $allReleases = @()
        for ($page = 1; $page -le 4; $page++) {
            try {
                $pageData = $wc.DownloadString("https://api.github.com/repos/Otzaria/otzaria/releases?per_page=100&page=$page") | ConvertFrom-Json
                if ($pageData.Count -eq 0) { break }
                $allReleases += $pageData
            } catch { break }
        }
        $json = $allReleases
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
        $wc2 = New-Object System.Net.WebClient
        $wc2.Headers.Add("User-Agent", "OtzariaUpdater")
        $wc2.Headers.Add("Cache-Control", "no-cache, no-store, must-revalidate")
        $wc2.Headers.Add("Pragma", "no-cache")
        $wc2.Encoding = [System.Text.Encoding]::UTF8
        # Try refs/heads/main endpoint first (more reliable)
        $timestamp = [DateTimeOffset]::Now.ToUnixTimeSeconds()
        try {
            $script:LibVer = $wc2.DownloadString("https://raw.githubusercontent.com/Otzaria/otzaria-library/refs/heads/main/MoreBooks/%D7%A1%D7%A4%D7%A8%D7%99%D7%9D/%D7%90%D7%95%D7%A6%D7%A8%D7%99%D7%90/%D7%90%D7%95%D7%93%D7%95%D7%AA%20%D7%94%D7%AA%D7%95%D7%9B%D7%A0%D7%94/%D7%92%D7%99%D7%A8%D7%A1%D7%AA%20%D7%A1%D7%A4%D7%A8%D7%99%D7%94.txt?t=$timestamp").Trim()
        } catch {
            # Fallback to main endpoint
            $script:LibVer = $wc2.DownloadString("https://raw.githubusercontent.com/Otzaria/otzaria-library/main/MoreBooks/%D7%A1%D7%A4%D7%A8%D7%99%D7%9D/%D7%90%D7%95%D7%A6%D7%A8%D7%99%D7%90/%D7%90%D7%95%D7%93%D7%95%D7%AA%20%D7%94%D7%AA%D7%95%D7%9B%D7%A0%D7%94/%D7%92%D7%99%D7%A8%D7%A1%D7%AA%20%D7%A1%D7%A4%D7%A8%D7%99%D7%94.txt?t=$timestamp").Trim()
        }
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
    # Use wildcard that catches files with RTL markers too
    $existingLibZip = Get-ChildItem "." -Filter "*otzaria_latest*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $existingLibZip) {
        # Try also otzaria_*.zip pattern
        $existingLibZip = Get-ChildItem "." -Filter "*otzaria_*.zip" -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -notlike "*Updater*" } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }
    
    if ($existingLibZip) {
        # Library file exists - check if newer version available
        $zipVer = Get-ZipLibVersion $existingLibZip.Name
        $script:LibFileToExtract = $existingLibZip.Name
        
        # Try to get version from filename if not found in ZIP
        if (-not $zipVer -and $existingLibZip.Name -match 'otzaria_latest_(\d+)\.zip') {
            $zipVer = $matches[1]
        }
        
        # Compare versions to see if newer is available on server
        $newerAvailable = $false
        if ($zipVer -and $script:LibVer -and $script:LibVer -ne "?") {
            try {
                $localNum = [int]($zipVer -replace '\D','')
                $serverNum = [int]($script:LibVer -replace '\D','')
                
                # If server version is much lower than local (likely cache issue), ignore it
                if ($serverNum -lt ($localNum - 10)) {
                    # Server returned stale data, treat as up-to-date
                    $script:LibVer = $zipVer
                    $newerAvailable = $false
                } elseif ($serverNum -gt $localNum) {
                    $newerAvailable = $true
                }
            } catch {
                # Fallback to string comparison
                $localParts = $zipVer -split '\.'
                $serverParts = $script:LibVer -split '\.'
                for ($i = 0; $i -lt [Math]::Max($localParts.Count, $serverParts.Count); $i++) {
                    $local = if ($i -lt $localParts.Count) { [int]$localParts[$i] } else { 0 }
                    $server = if ($i -lt $serverParts.Count) { [int]$serverParts[$i] } else { 0 }
                    if ($server -gt $local) { $newerAvailable = $true; break }
                    if ($local -gt $server) { break }
                }
            }
        }
        # If we can't determine version, assume newer is available
        if (-not $zipVer -and $script:LibVer -and $script:LibVer -ne "?") {
            $newerAvailable = $true
            $zipVer = "?"
        }
        
        if ($newerAvailable) {
            # Newer version on server - show "חלץ לאוצריא", "עדכון ארכיון" and "הורדה"
            $lblLibV.Text = "גרסה בארכיון: $zipVer   |   גרסה בשרת: $($script:LibVer)"
            
            $btnLibSelectFile.Visible = $false
            
            # Show 3 buttons: הורדה (left), עדכון ארכיון (center), חלץ לאוצריא (right)
            # Total width ~480, buttons 140 each with 30 spacing, centered in 654 card
            # Start at (654-480)/2 = 87
            $btnLibDL.Text = "הורדה"
            $btnLibDL.Size = New-Object System.Drawing.Size(140,34)
            $btnLibDL.Location = New-Object System.Drawing.Point(87,72)
            $btnLibDL.Visible = $true
            
            $btnLibUpdate.Text = "עדכון ארכיון"
            $btnLibUpdate.Size = New-Object System.Drawing.Size(140,34)
            $btnLibUpdate.Location = New-Object System.Drawing.Point(257,72)
            $btnLibUpdate.Visible = $true
            
            $btnLibExtract.Text = "חלץ לאוצריא"
            $btnLibExtract.Size = New-Object System.Drawing.Size(140,34)
            $btnLibExtract.Location = New-Object System.Drawing.Point(427,72)
            $btnLibExtract.Visible = $true
            
            $script:LibDownloadCompleted = $false
            $libStatusLbl.Visible = $true
            $libStatusLbl.Text = "לחץ לבדיקה"
        } else {
            # Same or newer version locally - library is up to date, show only extract
            $lblLibV.Text = "גרסה בארכיון: $zipVer"
            if ($script:LibVer -and $script:LibVer -ne "?") {
                $lblLibV.Text = "גרסה בארכיון: $zipVer   |   גרסה בשרת: $($script:LibVer)"
            }
            
            # Hide download and update buttons, show only extract
            $btnLibDL.Visible = $false
            $btnLibSelectFile.Visible = $false
            $btnLibUpdate.Visible = $false
            
            # Show extract button centered
            $btnLibExtract.Text = "חלץ לאוצריא"
            $btnLibExtract.Size = New-Object System.Drawing.Size(180,34)
            $btnLibExtract.Location = New-Object System.Drawing.Point(237,72)
            $btnLibExtract.Visible = $true
            
            $script:LibDownloadCompleted = $false
            $libStatusLbl.Visible = $true
            $libStatusLbl.Text = "הספרייה מעודכנת - לחץ לחילוץ לאוצריא"
        }
    } else {
        # No library file detected in standard check - check for ZIP in tool folder
        $folderZip = Get-ChildItem "." -Filter "*otzaria_latest*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $folderZip) {
            $folderZip = Get-ChildItem "." -Filter "*otzaria_*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        }
        
        if ($script:LibVer -and $script:LibVer -ne "?") {
            # Online mode
            $lblLibV.Text = "גרסה בשרת: $($script:LibVer)"
            $libSize = Get-Size $script:LibUrl
            if ($libSize -gt 0) {
                $libSizeMB = [math]::Round($libSize / 1MB, 0)
                $lblLibV.Text = "גרסה בשרת: $($script:LibVer)   |   גודל: $libSizeMB מ`"ב"
            }
            
            # Check if there's a ZIP in folder that's newer than installed library
            $showExtractButton = $false
            if ($folderZip) {
                $folderZipVer = Get-ZipLibVersion $folderZip.Name
                if (-not $folderZipVer -and $folderZip.Name -match 'otzaria_latest_(\d+)\.zip') {
                    $folderZipVer = $matches[1]
                }
                
                if ($folderZipVer) {
                    $script:LibFileToExtract = $folderZip.Name
                    
                    # Check if ZIP version is higher than installed
                    if ($script:InstalledLibVersion -eq "לא נמצא") {
                        $showExtractButton = $true
                    } else {
                        try {
                            $zipNum = [int]($folderZipVer -replace '\D','')
                            $installedNum = [int]($script:InstalledLibVersion -replace '\D','')
                            if ($zipNum -gt $installedNum) {
                                $showExtractButton = $true
                            }
                        } catch { }
                    }
                    
                    if ($showExtractButton) {
                        $lblLibV.Text = "גרסה בשרת: $($script:LibVer)   |   גרסה בארכיון: $folderZipVer"
                    }
                }
            }
            
            if ($showExtractButton) {
                # Show both download and extract buttons
                $btnLibDL.Text = "הורדה"
                $btnLibDL.Size = New-Object System.Drawing.Size(140,34)
                $btnLibDL.Location = New-Object System.Drawing.Point(177,72)
                $btnLibDL.Visible = $true
                
                $btnLibExtract.Text = "חלץ לאוצריא"
                $btnLibExtract.Size = New-Object System.Drawing.Size(140,34)
                $btnLibExtract.Location = New-Object System.Drawing.Point(337,72)
                $btnLibExtract.Visible = $true
                
                $btnLibSelectFile.Visible = $false
                $btnLibUpdate.Visible = $false
                
                $libStatusLbl.Visible = $true
                $libStatusLbl.Text = "נמצא ארכיון בתיקייה - ניתן להוריד או לחלץ"
            } else {
                # Only show download button centered
                $btnLibDL.Text = "הורדת הספרייה"
                $btnLibDL.Size = New-Object System.Drawing.Size(180,34)
                $btnLibDL.Location = New-Object System.Drawing.Point(237,72)
                $btnLibDL.Visible = $true
                $btnLibSelectFile.Visible = $false
                $btnLibExtract.Visible = $false
                $btnLibUpdate.Visible = $false
            }
        } else {
            # Offline and no file detected yet - check for existing ZIP in folder
            $offlineZip = Get-ChildItem "." -Filter "*otzaria_latest*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if (-not $offlineZip) {
                $offlineZip = Get-ChildItem "." -Filter "*otzaria_*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            }
            
            if ($offlineZip) {
                # Found ZIP file in offline mode
                $zipVer = Get-ZipLibVersion $offlineZip.Name
                if (-not $zipVer -and $offlineZip.Name -match 'otzaria_latest_(\d+)\.zip') {
                    $zipVer = $matches[1]
                }
                if (-not $zipVer) { $zipVer = "?" }
                
                $script:LibFileToExtract = $offlineZip.Name
                $lblLibV.Text = "גרסה בארכיון: $zipVer"
                $btnLibDL.Visible = $false
                $btnLibSelectFile.Visible = $false
                $btnLibUpdate.Visible = $false
                
                # Show extract button centered
                $btnLibExtract.Text = "חלץ לאוצריא"
                $btnLibExtract.Size = New-Object System.Drawing.Size(180,34)
                $btnLibExtract.Location = New-Object System.Drawing.Point(237,72)
                $btnLibExtract.Visible = $true
                
                $libStatusLbl.Visible = $true
                $libStatusLbl.Text = "קובץ ספרייה גרסה $zipVer קיים בתיקייה - לחץ לחילוץ לאוצריא"
            } else {
                # Offline and no file - show select file button centered
                $lblLibV.Text = "גרסה בשרת: ?"
                $btnLibDL.Visible = $false
                $btnLibSelectFile.Visible = $true
                $btnLibExtract.Visible = $false
                $btnLibUpdate.Visible = $false
                $libStatusLbl.Visible = $true
                $libStatusLbl.Text = "אין חיבור לרשת - בחר קובץ ספרייה"
            }
        }
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
    
    # Note: Differential update check is now done only when user clicks "עדכון ארכיון"
    
    # Check for downloaded files that don't match server versions
    Check-DownloadedFiles
})

$form.Add_FormClosing({
    $stopFile = "$env:TEMP\otzaria_stop.flag"
    $libStopFile = "$env:TEMP\otzaria_lateststop.flag"
    $vcStopFile = "$env:TEMP\otzaria_vc_stop.flag"
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
    if ($script:VCDownloadJob) {
        "stop" | Out-File $vcStopFile -Force
        Start-Sleep -Milliseconds 500
        Stop-Job -Job $script:VCDownloadJob -ErrorAction SilentlyContinue
        Remove-Job -Job $script:VCDownloadJob -Force -ErrorAction SilentlyContinue
    }
    $timer.Stop()
})

$radioStable.Add_CheckedChanged({ Update-Btn })
$radioPre.Add_CheckedChanged({ Update-Btn })
$radioFull.Add_CheckedChanged({ Update-Btn })
$radioDownloaded.Add_CheckedChanged({ Update-Btn })

# Resize handler - center elements when window size changes
$form.Add_Resize({
    if ($form.WindowState -ne "Minimized") {
        $formWidth = [int]$form.ClientSize.Width
        
        # Center main card
        $cardWidth = [Math]::Min(655, $formWidth - 40)
        $cardX = [Math]::Max(20, [int](($formWidth - $cardWidth) / 2))
        $mainCard.Width = $cardWidth
        $mainCard.Location = New-Object System.Drawing.Point($cardX, $mainCard.Location.Y)
        
        # Center library card
        $libCard.Width = $cardWidth
        $libCard.Location = New-Object System.Drawing.Point($cardX, $libCard.Location.Y)
        
        # Center bottom buttons
        if ($btnVCRedist.Visible) {
            # Both buttons visible - position them side by side centered
            $totalWidth = $btnClearCache.Width + 20 + $btnVCRedist.Width
            $startX = [int](($formWidth - $totalWidth) / 2)
            $btnClearCache.Location = New-Object System.Drawing.Point($startX, $btnClearCache.Location.Y)
            $btnVCRedist.Location = New-Object System.Drawing.Point(($startX + $btnClearCache.Width + 20), $btnVCRedist.Location.Y)
        } else {
            # Only clear cache button - center it
            $btnClearCache.Location = New-Object System.Drawing.Point([int](($formWidth - $btnClearCache.Width) / 2), $btnClearCache.Location.Y)
        }
        
        # Center status elements
        $statusDivider.Width = $cardWidth
        $statusDivider.Location = New-Object System.Drawing.Point($cardX, $statusDivider.Location.Y)
        $statusLocation.Location = New-Object System.Drawing.Point($cardX, $statusLocation.Location.Y)
        $statusBar.Location = New-Object System.Drawing.Point($cardX, $statusBar.Location.Y)
        
        # Keep header buttons on left side (high X values in RTL)
        $btnRefresh.Location = New-Object System.Drawing.Point(($formWidth - 83), 10)
        $btnTheme.Location = New-Object System.Drawing.Point(($formWidth - 130), 10)
    }
})

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

# Library Differential Update Button
$btnLibUpdate.Add_Click({
    # If currently updating - stop
    if ($script:DiffIsUpdating) {
        if ($script:DiffUpdateJob) {
            Stop-Job -Job $script:DiffUpdateJob -ErrorAction SilentlyContinue
            Remove-Job -Job $script:DiffUpdateJob -Force -ErrorAction SilentlyContinue
            $script:DiffUpdateJob = $null
        }
        $script:DiffIsUpdating = $false
        $libProgBar.Value = 0
        $libStatusLbl.Text = "העדכון בוטל"
        $btnLibUpdate.Text = "עדכון ארכיון"
        $btnLibExtract.Visible = $false
        try { [TBProg]::SetState($form.Handle, 0) } catch { }
        return
    }
    
    # Check for ZIP file
    $zipFile = $null
    if (Test-Path "otzaria_latest.zip") {
        $zipFile = "otzaria_latest.zip"
    } else {
        $zipFiles = Get-ChildItem "." -Filter "*otzaria_*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
        if ($zipFiles) { $zipFile = $zipFiles[0].Name }
    }
    
    if (-not $zipFile -or -not (Test-Path $zipFile)) {
        $libStatusLbl.Visible = $true
        $libStatusLbl.Text = "לא נמצא קובץ ארכיון לעדכון"
        return
    }
    
    # Check for update
    $libStatusLbl.Visible = $true
    $libStatusLbl.Text = "בודק עדכונים..."
    [System.Windows.Forms.Application]::DoEvents()
    
    $updateInfo = Check-LibraryUpdateAvailable
    
    if ($updateInfo.NoNetwork) {
        $libStatusLbl.Text = "אין חיבור לרשת"
        return
    }
    
    if (-not $updateInfo.Available) {
        $libStatusLbl.Text = "הספרייה מעודכנת"
        $btnLibUpdate.Visible = $false
        # Show only extract button centered
        $btnLibExtract.Size = New-Object System.Drawing.Size(180,34)
        $btnLibExtract.Location = New-Object System.Drawing.Point(237,72)
        return
    }
    
    # If no local manifest, warn user that this will be a large update
    if ($updateInfo.NoLocalManifest) {
        $msg = "לא נמצא files_manifest.json בארכיון.`n`nהעדכון יוריד את כל הקבצים מהשרת ($($updateInfo.TotalCount) קבצים).`n`nזה עלול לקחת זמן רב. האם להמשיך?"
        $r = Show-RTLMessageBox $msg "עדכון ארכיון" "YesNo" "Warning"
        if ($r -ne "Yes") {
            $libStatusLbl.Text = "העדכון בוטל"
            return
        }
    } else {
        $changes = $updateInfo.Changes
        $newCount = $updateInfo.NewCount
        $modCount = $updateInfo.ModifiedCount
        $delCount = $updateInfo.DeletedCount
        $totalCount = $updateInfo.TotalCount
        
        # Show confirmation dialog
        $msg = "נמצאו $totalCount שינויים:`n"
        if ($newCount -gt 0) { $msg += "• $newCount קבצים חדשים`n" }
        if ($modCount -gt 0) { $msg += "• $modCount קבצים שהשתנו`n" }
        if ($delCount -gt 0) { $msg += "• $delCount קבצים למחיקה`n" }
        $msg += "`nהעדכון יבוצע בקובץ הארכיון.`nהאם להמשיך?"
        
        $r = Show-RTLMessageBox $msg "עדכון ארכיון" "YesNo" "Question"
        if ($r -ne "Yes") {
            $libStatusLbl.Text = "העדכון בוטל"
            return
        }
    }
    
    $changes = $updateInfo.Changes
    
    # Start differential update - ZIP only
    $script:DiffIsUpdating = $true
    $btnLibUpdate.Text = "עצור"
    $btnLibDL.Visible = $false
    $btnLibExtract.Visible = $false
    $libStatusLbl.Text = "מתחיל עדכון..."
    $libProgBar.Value = 0
    
    # Prepare file lists
    $filesToDownload = @()
    $filesToDownload += $changes.New
    $filesToDownload += $changes.Modified
    $filesToDelete = $changes.Deleted
    
    $baseUrl = $script:GitHubRawBaseUrl
    $manifestUrl = $script:RemoteManifestUrl
    $resolvedZipPath = (Resolve-Path $zipFile -ErrorAction SilentlyContinue).Path
    
    $script:DiffUpdateJob = Start-Job -ScriptBlock {
        param($zipPath, $filesToDownload, $filesToDelete, $baseUrl, $manifestUrl)
        
        # Create log file
        $logFile = Join-Path (Split-Path $zipPath -Parent) "update_debug.txt"
        $logContent = @()
        $logContent += "=== Update Debug Log ==="
        $logContent += "Time: $(Get-Date)"
        $logContent += "ZipPath: $zipPath"
        $logContent += "BaseUrl: $baseUrl"
        $logContent += "ManifestUrl: $manifestUrl"
        $logContent += ""
        $logContent += "Files to download: $($filesToDownload.Count)"
        $logContent += "Files to delete: $($filesToDelete.Count)"
        $logContent += ""
        $logContent += "=== Sample paths to download (first 20) ==="
        $sampleCount = 0
        foreach ($f in $filesToDownload) {
            if ($sampleCount -lt 20) {
                $logContent += $f
                $sampleCount++
            }
        }
        $logContent += ""
        $logContent += "=== Sample paths to delete (first 10) ==="
        $sampleCount = 0
        foreach ($f in $filesToDelete) {
            if ($sampleCount -lt 10) {
                $logContent += $f
                $sampleCount++
            }
        }
        
        # Function to convert manifest path to ZIP path (defined inside job)
        # Returns $null for paths that should not be in the ZIP
        function ConvertPath($manifestPath) {
            # Root level files
            if ($manifestPath -eq "metadata.json" -or $manifestPath -eq "files_manifest.json") {
                return $manifestPath
            }
            
            # Pattern: */ספרים/אוצריא/... -> אוצריא/...
            # Matches: Ben-YehudaToOtzaria/ספרים/אוצריא/..., wikisourceToOtzaria/ספרים/אוצריא/..., 
            #          MoreBooks/ספרים/אוצריא/..., sefariaToOtzaria/sefaria_export/ספרים/אוצריא/...
            if ($manifestPath -match '^.*/ספרים/אוצריא/(.+)$') {
                return "אוצריא/$($matches[1])"
            }
            
            # Special case: DictaToOtzaria/ערוך/ספרים/אוצריא/... -> אוצריא/...
            if ($manifestPath -match '^.*/ערוך/ספרים/אוצריא/(.+)$') {
                return "אוצריא/$($matches[1])"
            }
            
            # Links files: */links/... -> links/...
            if ($manifestPath -match '^.*/links/(.+)$') {
                return "links/$($matches[1])"
            }
            
            # Any other path - skip (return null)
            return $null
        }
        
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            Add-Type -AssemblyName System.IO.Compression
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            
            # Filter files - only process files that belong in the ZIP
            $filteredDownloads = @()
            $skippedPaths = @()
            foreach ($f in $filesToDownload) {
                $converted = ConvertPath $f
                if ($converted) {
                    $filteredDownloads += $f
                } else {
                    $skippedPaths += $f
                }
            }
            $filteredDeletes = @()
            foreach ($f in $filesToDelete) {
                $converted = ConvertPath $f
                if ($converted) {
                    $filteredDeletes += $f
                }
            }
            
            $logContent += ""
            $logContent += "=== After filtering ==="
            $logContent += "Filtered downloads: $($filteredDownloads.Count)"
            $logContent += "Filtered deletes: $($filteredDeletes.Count)"
            $logContent += "Skipped paths: $($skippedPaths.Count)"
            $logContent += ""
            $logContent += "=== Skipped paths (first 20) ==="
            $sampleCount = 0
            foreach ($f in $skippedPaths) {
                if ($sampleCount -lt 20) {
                    $logContent += $f
                    $sampleCount++
                }
            }
            
            # Write log so far
            $logContent | Out-File $logFile -Encoding UTF8
            
            $total = $filteredDownloads.Count + $filteredDeletes.Count + 1  # +1 for manifest update
            $current = 0
            $encoding = [System.Text.Encoding]::GetEncoding(862)
            
            # Download and update new/modified files - to ZIP only
            foreach ($manifestFile in $filteredDownloads) {
                $current++
                $zipEntryPath = ConvertPath $manifestFile
                Write-Output "PROGRESS:${current}:${total}:$zipEntryPath"
                
                try {
                    # Properly encode each path segment separately for URL
                    $pathSegments = $manifestFile.Split("/")
                    $encodedSegments = $pathSegments | ForEach-Object { [System.Uri]::EscapeDataString($_) }
                    $encodedPath = $encodedSegments -join "/"
                    $url = "$baseUrl$encodedPath"
                    
                    $wc = New-Object System.Net.WebClient
                    $wc.Encoding = [System.Text.Encoding]::UTF8
                    $wc.Headers.Add("User-Agent", "OtzariaUpdater")
                    $contentBytes = $wc.DownloadData($url)
                    
                    # Update ZIP
                    $zipArchive = [System.IO.Compression.ZipFile]::Open($zipPath, 'Update', $encoding)
                    
                    # Remove existing entry if exists
                    $existingEntry = $zipArchive.Entries | Where-Object { $_.FullName -eq $zipEntryPath } | Select-Object -First 1
                    if ($existingEntry) {
                        $existingEntry.Delete()
                    }
                    
                    # Create parent directories in zip if needed
                    $parentPath = [System.IO.Path]::GetDirectoryName($zipEntryPath)
                    if ($parentPath) {
                        $parentPath = $parentPath.Replace("\", "/")
                        $pathParts = $parentPath.Split("/", [System.StringSplitOptions]::RemoveEmptyEntries)
                        $currentPath = ""
                        foreach ($part in $pathParts) {
                            if ($currentPath -eq "") {
                                $currentPath = $part + "/"
                            } else {
                                $currentPath = $currentPath + $part + "/"
                            }
                            $dirExists = $zipArchive.Entries | Where-Object { $_.FullName -eq $currentPath } | Select-Object -First 1
                            if (-not $dirExists) {
                                $dirEntry = $zipArchive.CreateEntry($currentPath)
                                $dirEntry.Open().Close()
                            }
                        }
                    }
                    
                    # Create new entry with converted path
                    $newEntry = $zipArchive.CreateEntry($zipEntryPath, [System.IO.Compression.CompressionLevel]::Optimal)
                    $entryStream = $newEntry.Open()
                    $entryStream.Write($contentBytes, 0, $contentBytes.Length)
                    $entryStream.Close()
                    
                    $zipArchive.Dispose()
                } catch [System.Net.WebException] {
                    Write-Output "WARN:Failed to download $manifestFile (HTTP error) - Skipping"
                } catch {
                    Write-Output "ERROR:Failed to update $manifestFile - $($_.Exception.Message)"
                    return
                }
            }
            
            # Delete removed files - from ZIP only
            foreach ($manifestFile in $filteredDeletes) {
                $current++
                $zipEntryPath = ConvertPath $manifestFile
                Write-Output "PROGRESS:${current}:${total}:Deleting $zipEntryPath"
                
                try {
                    $zipArchive = [System.IO.Compression.ZipFile]::Open($zipPath, 'Update', $encoding)
                    $existingEntry = $zipArchive.Entries | Where-Object { $_.FullName -eq $zipEntryPath } | Select-Object -First 1
                    if ($existingEntry) {
                        $existingEntry.Delete()
                    }
                    $zipArchive.Dispose()
                } catch {
                    # Ignore delete errors
                }
            }
            
            # Update the manifest file in ZIP
            $current++
            Write-Output "PROGRESS:${current}:${total}:Updating manifest"
            
            try {
                $wc = New-Object System.Net.WebClient
                $wc.Encoding = [System.Text.Encoding]::UTF8
                $wc.Headers.Add("User-Agent", "OtzariaUpdater")
                $manifestContent = $wc.DownloadString($manifestUrl)
                $manifestBytes = [System.Text.Encoding]::UTF8.GetBytes($manifestContent)
                
                $zipArchive = [System.IO.Compression.ZipFile]::Open($zipPath, 'Update', $encoding)
                
                $existingManifest = $zipArchive.Entries | Where-Object { $_.FullName -eq "files_manifest.json" } | Select-Object -First 1
                if ($existingManifest) {
                    $existingManifest.Delete()
                }
                
                $newEntry = $zipArchive.CreateEntry("files_manifest.json", [System.IO.Compression.CompressionLevel]::Optimal)
                $entryStream = $newEntry.Open()
                $entryStream.Write($manifestBytes, 0, $manifestBytes.Length)
                $entryStream.Close()
                
                $zipArchive.Dispose()
            } catch {
                Write-Output "ERROR:Failed to update manifest - $($_.Exception.Message)"
                return
            }
            
            Write-Output "DONE"
        } catch {
            Write-Output "ERROR:$($_.Exception.Message)"
        }
    } -ArgumentList $resolvedZipPath, $filesToDownload, $filesToDelete, $baseUrl, $manifestUrl
    
    try { [TBProg]::SetState($form.Handle, 2) } catch { }
})

# Theme Button
$btnTheme.Add_Click({
    $script:DarkMode = -not $script:DarkMode
    $script:DarkModeManual = $true  # User manually changed theme
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
    
    # Try to load data from server - PARALLEL LOADING
    $statusLbl.Text = "מרענן נתונים..."
    [System.Windows.Forms.Application]::DoEvents()
    
    # Helper function to get all releases
    $getAllReleases = {
        $allReleases = @()
        for ($page = 1; $page -le 4; $page++) {
            try {
                $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases?per_page=100&page=$page" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
                if ($r.Count -eq 0) { break }
                $allReleases += $r
            } catch { break }
        }
        return $allReleases
    }
    
    # Start all API calls in parallel
    $stableJob = Start-Job -ScriptBlock {
        try {
            $allReleases = @()
            for ($page = 1; $page -le 4; $page++) {
                try {
                    $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases?per_page=100&page=$page" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
                    if ($r.Count -eq 0) { break }
                    $allReleases += $r
                } catch { break }
            }
            $rel = $allReleases | Where-Object { $_.prerelease -eq $false } | Select-Object -First 1
            if (-not $rel) { return @{Release=$null; Changelog=""} }
            $a = $rel.assets | Where-Object { $_.name -like "otzaria-*-windows.exe" -and $_.name -notlike "*-full.exe" } | Select-Object -First 1
            if (-not $a) { return @{Release=$null; Changelog=""} }
            $changelog = if ($rel.body) { $rel.body } else { "אין מידע על שינויים" }
            if ($rel.tag_name -match '^([\d\.]+)\+(\d+)$') {
                return @{
                    Release = @{ FullVersion = "$($matches[1]).$($matches[2])"; File = "otzaria-$($matches[1]).$($matches[2])-windows.exe"; Url = $a.browser_download_url }
                    Changelog = $changelog
                }
            }
        } catch { }
        return @{Release=$null; Changelog=""}
    }
    
    $preJob = Start-Job -ScriptBlock {
        try {
            $allReleases = @()
            for ($page = 1; $page -le 4; $page++) {
                try {
                    $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases?per_page=100&page=$page" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
                    if ($r.Count -eq 0) { break }
                    $allReleases += $r
                } catch { break }
            }
            $rel = $allReleases | Where-Object { $_.prerelease -eq $true } | Select-Object -First 1
            if (-not $rel) { return @{Release=$null; Changelog=""} }
            $a = $rel.assets | Where-Object { $_.name -like "otzaria-*-windows.exe" -and $_.name -notlike "*-full.exe" } | Select-Object -First 1
            if (-not $a) { return @{Release=$null; Changelog=""} }
            $changelog = if ($rel.body) { $rel.body } else { "אין מידע על שינויים" }
            if ($rel.tag_name -match '^([\d\.]+)\+(\d+)$') {
                return @{
                    Release = @{ FullVersion = "$($matches[1]).$($matches[2])"; File = "otzaria-$($matches[1]).$($matches[2])-windows.exe"; Url = $a.browser_download_url }
                    Changelog = $changelog
                }
            }
        } catch { }
        return @{Release=$null; Changelog=""}
    }
    
    $fullJob = Start-Job -ScriptBlock {
        try {
            $allReleases = @()
            for ($page = 1; $page -le 4; $page++) {
                try {
                    $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases?per_page=100&page=$page" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
                    if ($r.Count -eq 0) { break }
                    $allReleases += $r
                } catch { break }
            }
            foreach ($rel in $allReleases) {
                $a = $rel.assets | Where-Object { $_.name -like "otzaria-*-windows-full.exe" } | Select-Object -First 1
                if ($a) {
                    if ($rel.tag_name -match '^([\d\.]+)\+(\d+)$') {
                        $fullVer = "$($matches[1]).$($matches[2])"
                        return @{ FullVersion = $fullVer; File = "otzaria-$fullVer-windows-full.exe"; Url = $a.browser_download_url }
                    }
                }
            }
        } catch { }
        return $null
    }
    
    $stableMsixJob = Start-Job -ScriptBlock {
        try {
            $allReleases = @()
            for ($page = 1; $page -le 4; $page++) {
                try {
                    $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases?per_page=100&page=$page" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
                    if ($r.Count -eq 0) { break }
                    $allReleases += $r
                } catch { break }
            }
            $rel = $allReleases | Where-Object { $_.prerelease -eq $false } | Select-Object -First 1
            if (-not $rel) { return $null }
            $a = $rel.assets | Where-Object { $_.name -like "*.msix" } | Select-Object -First 1
            if (-not $a) { return $null }
            if ($rel.tag_name -match '^([\d\.]+)\+(\d+)$') {
                $fullVer = "$($matches[1]).$($matches[2])"
                return @{ FullVersion = $fullVer; File = "otzaria-$fullVer.msix"; Url = $a.browser_download_url }
            }
        } catch { }
        return $null
    }
    
    $preMsixJob = Start-Job -ScriptBlock {
        try {
            $allReleases = @()
            for ($page = 1; $page -le 4; $page++) {
                try {
                    $r = Invoke-RestMethod "https://api.github.com/repos/Otzaria/otzaria/releases?per_page=100&page=$page" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
                    if ($r.Count -eq 0) { break }
                    $allReleases += $r
                } catch { break }
            }
            $rel = $allReleases | Where-Object { $_.prerelease -eq $true } | Select-Object -First 1
            if (-not $rel) { return $null }
            $a = $rel.assets | Where-Object { $_.name -like "*.msix" } | Select-Object -First 1
            if (-not $a) { return $null }
            if ($rel.tag_name -match '^([\d\.]+)\+(\d+)$') {
                $fullVer = "$($matches[1]).$($matches[2])"
                return @{ FullVersion = $fullVer; File = "otzaria-$fullVer.msix"; Url = $a.browser_download_url }
            }
        } catch { }
        return $null
    }
    
    $libVerJob = Start-Job -ScriptBlock {
        try {
            $wc = New-Object System.Net.WebClient
            $wc.Encoding = [System.Text.Encoding]::UTF8
            $wc.Headers.Add("User-Agent", "OtzariaUpdater")
            $wc.Headers.Add("Cache-Control", "no-cache")
            $timestamp = [DateTimeOffset]::Now.ToUnixTimeSeconds()
            return $wc.DownloadString("https://raw.githubusercontent.com/Otzaria/otzaria-library/refs/heads/main/MoreBooks/%D7%A1%D7%A4%D7%A8%D7%99%D7%9D/%D7%90%D7%95%D7%A6%D7%A8%D7%99%D7%90/%D7%90%D7%95%D7%93%D7%95%D7%AA%20%D7%94%D7%AA%D7%95%D7%9B%D7%A0%D7%94/%D7%92%D7%99%D7%A8%D7%A1%D7%AA%20%D7%A1%D7%A4%D7%A8%D7%99%D7%94.txt?t=$timestamp").Trim()
        } catch { }
        return "?"
    }
    
    # Wait for all jobs with UI updates
    $allJobs = @($stableJob, $preJob, $fullJob, $stableMsixJob, $preMsixJob, $libVerJob)
    while ($allJobs | Where-Object { $_.State -eq 'Running' }) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 50
    }
    
    # Collect results
    $stableData = Receive-Job -Job $stableJob -ErrorAction SilentlyContinue
    $preData = Receive-Job -Job $preJob -ErrorAction SilentlyContinue
    $script:FullRelease = Receive-Job -Job $fullJob -ErrorAction SilentlyContinue
    $script:StableMsixRelease = Receive-Job -Job $stableMsixJob -ErrorAction SilentlyContinue
    $script:PreMsixRelease = Receive-Job -Job $preMsixJob -ErrorAction SilentlyContinue
    $script:LibVer = Receive-Job -Job $libVerJob -ErrorAction SilentlyContinue
    
    # Clean up jobs
    $allJobs | Remove-Job -Force -ErrorAction SilentlyContinue
    
    # Process results
    if ($stableData -and $stableData.Release) {
        $script:StableRelease = [pscustomobject]$stableData.Release
        $script:StableChangelog = $stableData.Changelog
    }
    if ($preData -and $preData.Release) {
        $script:PreRelease = [pscustomobject]$preData.Release
        $script:PreChangelog = $preData.Changelog
    }
    if ($script:FullRelease) { $script:FullRelease = [pscustomobject]$script:FullRelease }
    if ($script:StableMsixRelease) { $script:StableMsixRelease = [pscustomobject]$script:StableMsixRelease }
    if ($script:PreMsixRelease) { $script:PreMsixRelease = [pscustomobject]$script:PreMsixRelease }
    if (-not $script:LibVer) { $script:LibVer = "?" }
    
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
        
        $script:LibFinalFile = "otzaria_latest_$($script:LibVer).zip"
        $script:LibTempFile = "otzaria_latest_$($script:LibVer).part"
    
        # Check for existing library files first
        $existingLibZip = Get-ChildItem "." -Filter "*otzaria_latest*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
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
        $existingLibZip = Get-ChildItem "." -Filter "*otzaria_latest*.zip" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
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
    # If download is completed and we're offering to extract - just reset
    if ($script:LibDownloadCompleted) {
        # Reset state - user can use "עדכון ארכיון" button instead
        $script:LibDownloadCompleted = $false
        $script:LibFileToExtract = $null
        $btnLibDL.Text = "הורדת הספרייה"
        $libStatusLbl.Visible = $true
        $libStatusLbl.Text = "לחץ להורדת ספרייה חדשה"
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
        $btnLibExtract.Visible = $false
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
        # File exists - offer to download fresh or use update button
        $r = Show-RTLMessageBox "קובץ ספרייה כבר קיים בתיקייה.`n`nהאם להוריד מחדש? (הקובץ הקיים יימחק)`n`nלחילופין, השתמש בכפתור 'עדכון ארכיון' לעדכון דיפרנציאלי." "קובץ קיים" "YesNo" "Question"
        if ($r -ne "Yes") {
            $libStatusLbl.Visible = $true
            $libStatusLbl.Text = "לחץ 'עדכון ארכיון' לעדכון הארכיון"
            return
        }
        # Delete existing file
        Remove-Item $existingFile -Force -ErrorAction SilentlyContinue
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
    $btnLibExtract.Visible = $false
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
        
        # Show extract button (not download button)
        $btnLibDL.Visible = $false
        $btnLibSelectFile.Visible = $false
        $btnLibUpdate.Visible = $false
        
        $btnLibExtract.Text = "חלץ לאוצריא"
        $btnLibExtract.Size = New-Object System.Drawing.Size(180,34)
        $btnLibExtract.Location = New-Object System.Drawing.Point(237,72)
        $btnLibExtract.Visible = $true
        
        $libStatusLbl.Visible = $true
        $libStatusLbl.Text = "קובץ נבחר: $fileName - לחץ לחילוץ"
    }
})

# Library Extract Button - extracts ZIP to Otzaria install folder
$btnLibExtract.Add_Click({
    # Extract existing library file to install folder
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
    
    # Check if archive version is newer than installed - offer update extraction
    $zipVer = Get-ZipLibVersion $fileToExtract
    $installedVer = $script:InstalledLibVersion
    $offerUpdateExtract = $false
    
    if ($zipVer -and $installedVer -and $installedVer -ne "לא נמצא") {
        try {
            $zipNum = [int]($zipVer -replace '\D','')
            $installedNum = [int]($installedVer -replace '\D','')
            if ($zipNum -gt $installedNum) {
                $offerUpdateExtract = $true
            }
        } catch { }
    }
    
    $extractMode = "full"  # Default to full extraction
    
    if ($offerUpdateExtract) {
        # Create custom dialog for extraction choice
        $choiceForm = New-Object System.Windows.Forms.Form
        $choiceForm.Text = "אפשרויות חילוץ"
        $choiceForm.Size = New-Object System.Drawing.Size(420, 180)
        $choiceForm.StartPosition = "CenterParent"
        $choiceForm.FormBorderStyle = "FixedDialog"
        $choiceForm.MaximizeBox = $false
        $choiceForm.MinimizeBox = $false
        $choiceForm.RightToLeft = "Yes"
        $choiceForm.RightToLeftLayout = $true
        $choiceForm.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        
        $lblQuestion = New-Object System.Windows.Forms.Label
        $lblQuestion.Text = "האם הינך מעוניין לחלץ הכל או לחלץ רק את קבצי העדכון מתוך הארכיון?"
        $lblQuestion.Location = New-Object System.Drawing.Point(20, 25)
        $lblQuestion.Size = New-Object System.Drawing.Size(370, 50)
        $lblQuestion.TextAlign = "MiddleCenter"
        $choiceForm.Controls.Add($lblQuestion)
        
        $btnExtractAll = New-Object System.Windows.Forms.Button
        $btnExtractAll.Text = "חלץ הכל"
        $btnExtractAll.Size = New-Object System.Drawing.Size(120, 40)
        $btnExtractAll.Location = New-Object System.Drawing.Point(60, 85)
        $btnExtractAll.Add_Click({
            $choiceForm.Tag = "full"
            $choiceForm.Close()
        })
        $choiceForm.Controls.Add($btnExtractAll)
        
        $btnExtractUpdate = New-Object System.Windows.Forms.Button
        $btnExtractUpdate.Text = "חלץ עדכון"
        $btnExtractUpdate.Size = New-Object System.Drawing.Size(120, 40)
        $btnExtractUpdate.Location = New-Object System.Drawing.Point(230, 85)
        $btnExtractUpdate.Add_Click({
            $choiceForm.Tag = "update"
            $choiceForm.Close()
        })
        $choiceForm.Controls.Add($btnExtractUpdate)
        
        $choiceForm.ShowDialog() | Out-Null
        
        if ($choiceForm.Tag -eq "update") {
            $extractMode = "update"
        } elseif ($choiceForm.Tag -eq "full") {
            $extractMode = "full"
        } else {
            # User closed the dialog without choosing
            $libStatusLbl.Text = "החילוץ בוטל"
            return
        }
        $choiceForm.Dispose()
    }
    
    if ($extractMode -eq "update") {
        # Update extraction - compare files between ZIP and installed folder
        $libStatusLbl.Text = "בודק קבצים לעדכון..."
        [System.Windows.Forms.Application]::DoEvents()
        
        $btnLibExtract.Visible = $false
        $btnLibUpdate.Visible = $false
        
        try {
            Add-Type -AssemblyName System.IO.Compression
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $encoding = [System.Text.Encoding]::GetEncoding(862)
            $zipPath = (Resolve-Path $fileToExtract).Path
            $archive = [System.IO.Compression.ZipFile]::Open($zipPath, 'Read', $encoding)
            
            # Build list of all files in ZIP (only אוצריא/, links/, files_manifest.json, metadata.json)
            $zipFiles = @{}
            foreach ($entry in $archive.Entries) {
                $entryPath = $entry.FullName
                # Skip directories
                if ($entryPath.EndsWith("/") -or $entryPath.EndsWith("\") -or [string]::IsNullOrEmpty($entry.Name)) {
                    continue
                }
                # Only include relevant paths
                if ($entryPath -like "אוצריא/*" -or $entryPath -like "links/*" -or 
                    $entryPath -eq "files_manifest.json" -or $entryPath -eq "metadata.json") {
                    $zipFiles[$entryPath] = @{
                        Entry = $entry
                        LastWrite = $entry.LastWriteTime.DateTime
                    }
                }
            }
            
            # Build list of all local files
            $localFiles = @{}
            $libFolder = Join-Path $p "אוצריא"
            $linksFolder = Join-Path $p "links"
            
            if (Test-Path $libFolder) {
                Get-ChildItem $libFolder -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
                    $relativePath = "אוצריא/" + $_.FullName.Substring($libFolder.Length + 1).Replace("\", "/")
                    $localFiles[$relativePath] = @{
                        FullPath = $_.FullName
                        LastWrite = $_.LastWriteTime
                    }
                }
            }
            
            if (Test-Path $linksFolder) {
                Get-ChildItem $linksFolder -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
                    $relativePath = "links/" + $_.FullName.Substring($linksFolder.Length + 1).Replace("\", "/")
                    $localFiles[$relativePath] = @{
                        FullPath = $_.FullName
                        LastWrite = $_.LastWriteTime
                    }
                }
            }
            
            $manifestPath = Join-Path $p "files_manifest.json"
            if (Test-Path $manifestPath) {
                $localFiles["files_manifest.json"] = @{
                    FullPath = $manifestPath
                    LastWrite = (Get-Item $manifestPath).LastWriteTime
                }
            }
            
            $metadataPath = Join-Path $p "metadata.json"
            if (Test-Path $metadataPath) {
                $localFiles["metadata.json"] = @{
                    FullPath = $metadataPath
                    LastWrite = (Get-Item $metadataPath).LastWriteTime
                }
            }
            
            # Find files to extract (new or newer in ZIP)
            $filesToExtract = @()
            foreach ($zipPath in $zipFiles.Keys) {
                $zipInfo = $zipFiles[$zipPath]
                if (-not $localFiles.ContainsKey($zipPath)) {
                    # File doesn't exist locally - add it
                    $filesToExtract += $zipPath
                } elseif ($zipInfo.LastWrite -gt $localFiles[$zipPath].LastWrite) {
                    # File in ZIP is newer - update it
                    $filesToExtract += $zipPath
                }
            }
            
            # Find files to delete (exist locally but not in ZIP)
            $filesToDelete = @()
            foreach ($localPath in $localFiles.Keys) {
                if (-not $zipFiles.ContainsKey($localPath)) {
                    $filesToDelete += $localFiles[$localPath].FullPath
                }
            }
            
            $totalOperations = $filesToExtract.Count + $filesToDelete.Count
            
            if ($totalOperations -eq 0) {
                $archive.Dispose()
                $libStatusLbl.Text = "הספרייה מעודכנת - אין קבצים לעדכון"
                $btnLibExtract.Visible = $true
                return
            }
            
            $libStatusLbl.Text = "מעדכן: $($filesToExtract.Count) קבצים להוספה, $($filesToDelete.Count) למחיקה"
            [System.Windows.Forms.Application]::DoEvents()
            
            $libProgBar.Value = 0
            $current = 0
            
            # Delete files that don't exist in ZIP
            foreach ($filePath in $filesToDelete) {
                $current++
                $pct = [math]::Round(($current / $totalOperations) * 100)
                $libProgBar.Value = $pct
                $libStatusLbl.Text = "מוחק: $current / $totalOperations"
                [System.Windows.Forms.Application]::DoEvents()
                
                try {
                    Remove-Item $filePath -Force -ErrorAction SilentlyContinue
                } catch { }
            }
            
            # Extract new/updated files
            foreach ($zipEntryPath in $filesToExtract) {
                $current++
                $pct = [math]::Round(($current / $totalOperations) * 100)
                $libProgBar.Value = $pct
                $libStatusLbl.Text = "מחלץ: $current / $totalOperations"
                [System.Windows.Forms.Application]::DoEvents()
                
                $entry = $zipFiles[$zipEntryPath].Entry
                $targetPath = Join-Path $p $zipEntryPath
                $targetDir = [System.IO.Path]::GetDirectoryName($targetPath)
                
                if (-not [string]::IsNullOrEmpty($targetDir) -and -not (Test-Path $targetDir)) {
                    [System.IO.Directory]::CreateDirectory($targetDir) | Out-Null
                }
                
                try {
                    if (Test-Path $targetPath) {
                        Remove-Item $targetPath -Force
                    }
                    $stream = $entry.Open()
                    $fileStream = [System.IO.File]::Create($targetPath)
                    $stream.CopyTo($fileStream)
                    $fileStream.Close()
                    $stream.Close()
                } catch { }
            }
            
            $archive.Dispose()
            
            # Remove empty directories
            if (Test-Path $libFolder) {
                Get-ChildItem $libFolder -Recurse -Directory -ErrorAction SilentlyContinue | 
                    Where-Object { (Get-ChildItem $_.FullName -ErrorAction SilentlyContinue).Count -eq 0 } |
                    Remove-Item -Force -ErrorAction SilentlyContinue
            }
            if (Test-Path $linksFolder) {
                Get-ChildItem $linksFolder -Recurse -Directory -ErrorAction SilentlyContinue | 
                    Where-Object { (Get-ChildItem $_.FullName -ErrorAction SilentlyContinue).Count -eq 0 } |
                    Remove-Item -Force -ErrorAction SilentlyContinue
            }
            
            $libProgBar.Value = 100
            $libStatusLbl.Text = "העדכון הושלם! ($($filesToExtract.Count) קבצים עודכנו, $($filesToDelete.Count) נמחקו)"
            
            # Update installed library version display
            $script:InstalledLibVersion = Get-InstalledLibVersion $script:InstallPath
            $statusBar.Text = "גרסה מותקנת: $($script:InstalledVersion)   |   גרסת ספרייה מותקנת: $($script:InstalledLibVersion)"
            
        } catch {
            $libStatusLbl.Text = "שגיאה בעדכון: $($_.Exception.Message)"
        }
        
        $libProgBar.Value = 0
        $btnLibExtract.Visible = $true
        return
    }
    
    # Full extraction mode (original code)
    # Warning before deletion
    $r = Show-RTLMessageBox "שים לב: כעת יימחקו כל הספרים שבספרייה.`n`nהאם להמשיך?" "אישור מחיקה" "YesNo" "Warning"
    if ($r -ne "Yes") {
        $libStatusLbl.Text = "החילוץ בוטל"
        return
    }
    
    $btnLibExtract.Visible = $false
    $btnLibUpdate.Visible = $false
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
        return
    }
    
    $libStatusLbl.Text = "מחלץ קבצים..."
    $libProgBar.Value = 0
    [System.Windows.Forms.Application]::DoEvents()
    
    # Create error log
    $errorLog = @()
    $errorLog += "=== Extract Log ==="
    $errorLog += "Time: $(Get-Date)"
    $errorLog += "Target: $p"
    $errorLog += "File: $fileToExtract"
    $errorLog += ""
    
    try {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $encoding = [System.Text.Encoding]::GetEncoding(862)
        $zipPath = (Resolve-Path ".\$fileToExtract").Path
        $errorLog += "ZipPath: $zipPath"
        
        $archive = [System.IO.Compression.ZipFile]::Open($zipPath, 'Read', $encoding)
        $totalEntries = $archive.Entries.Count
        $errorLog += "Total entries: $totalEntries"
        $errorLog += ""
        $currentEntry = 0
        
        foreach ($entry in $archive.Entries) {
            $currentEntry++
            $pct = [math]::Round(($currentEntry / $totalEntries) * 100)
            $libProgBar.Value = $pct
            $libStatusLbl.Text = "מחלץ: $currentEntry / $totalEntries"
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                $entryName = $entry.FullName
                if ([string]::IsNullOrEmpty($entryName)) {
                    continue
                }
                
                $targetPath = Join-Path $p $entryName
                
                # Skip if it's a directory entry (ends with /)
                if ($entryName.EndsWith("/") -or $entryName.EndsWith("\")) {
                    if (-not (Test-Path $targetPath)) {
                        [System.IO.Directory]::CreateDirectory($targetPath) | Out-Null
                    }
                    continue
                }
                
                # Create directory if needed
                $targetDir = [System.IO.Path]::GetDirectoryName($targetPath)
                if (-not [string]::IsNullOrEmpty($targetDir) -and -not (Test-Path $targetDir)) {
                    [System.IO.Directory]::CreateDirectory($targetDir) | Out-Null
                }
                
                # Extract file
                if ($entry.Name -ne "") {
                    if (Test-Path $targetPath) {
                        Remove-Item $targetPath -Force
                    }
                    $stream = $entry.Open()
                    $fileStream = [System.IO.File]::Create($targetPath)
                    $stream.CopyTo($fileStream)
                    $fileStream.Close()
                    $stream.Close()
                }
            } catch {
                $errorLog += "ERROR extracting: $($entry.FullName)"
                $errorLog += "  Target: $targetPath"
                $errorLog += "  Exception: $($_.Exception.Message)"
            }
        }
        $archive.Dispose()
        
        # Save error log if there were errors
        $errorLog += ""
        $errorLog += "=== Extraction completed ==="
        $errorLog | Out-File ".\extract_log.txt" -Encoding UTF8
        
        $libProgBar.Value = 100
        $libStatusLbl.Text = "הספרייה חולצה בהצלחה!"
        
        # Update installed library version display
        $script:InstalledLibVersion = Get-InstalledLibVersion $script:InstallPath
        $statusBar.Text = "גרסה מותקנת: $($script:InstalledVersion)   |   גרסת ספרייה מותקנת: $($script:InstalledLibVersion)"
        
        # Ask to delete zip file
        $delR = Show-RTLMessageBox "האם למחוק את קובץ הארכיון?" "מחיקת קובץ" "YesNo" "Question"
        if ($delR -eq "Yes") {
            Remove-Item $fileToExtract -Force -ErrorAction SilentlyContinue
            $libStatusLbl.Text = "הספרייה חולצה והארכיון נמחק"
            $script:LibFileToExtract = $null
        }
    } catch {
        $errorLog += ""
        $errorLog += "=== FATAL ERROR ==="
        $errorLog += $_.Exception.Message
        $errorLog += $_.ScriptStackTrace
        $errorLog | Out-File ".\extract_log.txt" -Encoding UTF8
        $libStatusLbl.Text = "שגיאה בחילוץ: $($_.Exception.Message)"
    }
    $libProgBar.Value = 0
    $btnLibExtract.Visible = $true
})

# VC++ Redistributable Button
$btnVCRedist.Add_Click({
    # If downloading - cancel
    if ($script:VCDownloading -and $btnVCRedist.Text -eq "ביטול") {
        "stop" | Out-File $script:VCStopFile -Force
        return
    }
    
    # Start download
    $r = Show-RTLMessageBox "האם להתקין את Visual C++ Redistributable?`n`nזוהי חבילה נדרשת להפעלת תוכנת אוצריא." "התקנת הרחבה" "YesNo" "Question"
    if ($r -eq "Yes") {
        $vcFile = Join-Path $env:TEMP "VisualCppRedist_AIO_x86_x64.exe"
        $vcStopFile = "$env:TEMP\otzaria_vc_stop.flag"
        
        # Remove stop flag if exists
        Remove-Item $vcStopFile -Force -ErrorAction SilentlyContinue
        
        # Show cancel button, hide VC button
        $btnVCRedist.Text = "ביטול"
        $script:VCDownloading = $true
        $script:VCStopFile = $vcStopFile
        
        # Get latest release URL from GitHub
        $libStatusLbl.Visible = $true
        $libStatusLbl.Text = "מחפש גרסה אחרונה..."
        [System.Windows.Forms.Application]::DoEvents()
        
        $vcUrl = $null
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            $releases = Invoke-RestMethod "https://api.github.com/repos/abbodi1406/vcredist/releases/latest" -Headers @{"User-Agent"="PS"} -TimeoutSec 15
            $asset = $releases.assets | Where-Object { $_.name -eq "VisualCppRedist_AIO_x86_x64.exe" } | Select-Object -First 1
            if ($asset) {
                $vcUrl = $asset.browser_download_url
            }
        } catch {
            $libStatusLbl.Text = "שגיאה בחיפוש גרסה: $_"
            $btnVCRedist.Text = "התקנת הרחבה"
            $script:VCDownloading = $false
            return
        }
        
        if (-not $vcUrl) {
            $libStatusLbl.Text = "לא נמצא קובץ להורדה"
            $btnVCRedist.Text = "התקנת הרחבה"
            $script:VCDownloading = $false
            return
        }
        
        # Start download job
        $script:VCDownloadJob = Start-Job -ScriptBlock {
            param($url, $file, $stopFile)
            try {
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                
                # Get file size
                $request = [System.Net.HttpWebRequest]::Create($url)
                $request.Method = "HEAD"
                $request.UserAgent = "PS"
                $request.AllowAutoRedirect = $true
                $response = $request.GetResponse()
                $totalSize = $response.ContentLength
                $response.Close()
                
                Write-Output "SIZE:$totalSize"
                
                # Download with large buffer for speed
                $webRequest = [System.Net.HttpWebRequest]::Create($url)
                $webRequest.UserAgent = "PS"
                $webRequest.AllowAutoRedirect = $true
                $webResponse = $webRequest.GetResponse()
                $stream = $webResponse.GetResponseStream()
                $fileStream = [System.IO.File]::Create($file)
                
                $buffer = New-Object byte[] 262144  # 256KB buffer for speed
                $totalRead = 0
                $lastReport = 0
                
                while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    # Check for stop
                    if (Test-Path $stopFile) {
                        $fileStream.Close()
                        $stream.Close()
                        $webResponse.Close()
                        Write-Output "STOPPED"
                        return
                    }
                    
                    $fileStream.Write($buffer, 0, $read)
                    $totalRead += $read
                    
                    if (($totalRead - $lastReport) -gt 102400) {
                        $pct = [math]::Min([math]::Round(($totalRead / $totalSize) * 100), 100)
                        Write-Output "PROGRESS:$totalRead`:$pct"
                        $lastReport = $totalRead
                    }
                }
                
                $fileStream.Close()
                $stream.Close()
                $webResponse.Close()
                Write-Output "DONE"
            } catch {
                Write-Output "ERROR:$_"
            }
        } -ArgumentList $vcUrl, $vcFile, $vcStopFile
        
        $script:VCFile = $vcFile
        $script:VCTotalSize = 0
        
        $libStatusLbl.Visible = $true
        $libStatusLbl.Text = "מוריד VC++ AIO..."
        $libProgBar.Value = 0
    }
})

# Timer tick handler for VC++ download - add to existing timer
$script:VCDownloading = $false
$script:VCDownloadJob = $null
$script:VCStopFile = ""

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
