#Requires -Version 5.1
# =====================================================================
#  אוצריא - הורדה, התקנה ועדכון (גרסה 2.0)
#  מותאם למבנה ההפצה החדש של אוצריא (0.9.98 ומעלה):
#   - otzaria-release-manifest.json, קבצים מפוצלים (part-NNN) עם SHA256
#   - ספרייה במסד נתונים seforim.db (מאגר Otzaria/SeforimLibrary)
#   - חנות תוספים (otzaria.org/plugins)
# =====================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

$script:UpdaterVersion = '2.0'
$script:AppRepo        = 'Otzaria/otzaria'
$script:LibRepo        = 'Otzaria/SeforimLibrary'
$script:AppId          = '{EEC4F712-CD05-4D15-A753-509E840A51A5}_is1'
$script:PluginsSite    = 'https://otzaria.org'
$script:UserAgent      = 'OtzariaUpdater/2.0'

# ---------------------------------------------------------------------
# Working directory (script folder, or EXE folder when compiled with PS2EXE)
# ---------------------------------------------------------------------
$script:WorkDir = $null
try { if ($PSCommandPath) { $script:WorkDir = Split-Path -Parent $PSCommandPath } } catch { }
if (-not $script:WorkDir) {
    try {
        $selfExe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($selfExe -notmatch '\\(powershell|pwsh|powershell_ise)\.exe$') { $script:WorkDir = Split-Path -Parent $selfExe }
    } catch { }
}
if (-not $script:WorkDir) {
    # הורץ דרך "irm ... | iex" - אין תיקיית סקריפט; ההורדות לתיקייה קבועה ולא לתיקייה שבה פתוח הטרמינל
    $script:WorkDir = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads\Otzaria'
    try { New-Item -ItemType Directory -Force -Path $script:WorkDir | Out-Null } catch { $script:WorkDir = (Get-Location).Path }
}
try {
    $probe = Join-Path $script:WorkDir '.otzaria_write_test'
    [IO.File]::WriteAllText($probe, 'x'); Remove-Item -LiteralPath $probe -Force
} catch {
    $script:WorkDir = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads\Otzaria'
    New-Item -ItemType Directory -Force -Path $script:WorkDir | Out-Null
}
Set-Location -LiteralPath $script:WorkDir

# ---------------------------------------------------------------------
# Native helpers
# ---------------------------------------------------------------------
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DpiHelper { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); }
public class DwmHelper {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    public static void SetTitleBarColor(IntPtr hwnd, int color) { try { DwmSetWindowAttribute(hwnd, 35, ref color, sizeof(int)); } catch { } }
    public static void SetDarkMode(IntPtr hwnd, bool dark) { try { int v = dark ? 1 : 0; DwmSetWindowAttribute(hwnd, 20, ref v, sizeof(int)); } catch { } }
}
"@ -ErrorAction SilentlyContinue
    [void][DpiHelper]::SetProcessDPIAware()
} catch { }

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

Add-Type -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
public class ModernProgressBar : Control {
    private int val = 0; private int max = 100;
    private Color barColor = Color.FromArgb(200, 160, 120);
    private Color bgColor = Color.FromArgb(240, 230, 220);
    private Color textColor = Color.FromArgb(100, 70, 50);
    private bool isRTL = true; private bool showText = true;
    public int Value { get { return val; } set { val = Math.Min(Math.Max(value, 0), max); Invalidate(); } }
    public int Maximum { get { return max; } set { max = value; Invalidate(); } }
    public Color BarColor { get { return barColor; } set { barColor = value; Invalidate(); } }
    public Color BgColor { get { return bgColor; } set { bgColor = value; Invalidate(); } }
    public Color TextColor { get { return textColor; } set { textColor = value; Invalidate(); } }
    public bool IsRTL { get { return isRTL; } set { isRTL = value; Invalidate(); } }
    public bool ShowText { get { return showText; } set { showText = value; Invalidate(); } }
    public ModernProgressBar() {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        Height = 22; Font = new Font("Segoe UI", 9, FontStyle.Bold);
    }
    protected override void OnPaint(PaintEventArgs e) {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        using (GraphicsPath bgPath = RoundedRect(new Rectangle(0, 0, Width - 1, Height - 1), 8))
        using (SolidBrush bgBrush = new SolidBrush(bgColor)) { e.Graphics.FillPath(bgBrush, bgPath); }
        if (val > 0 && max > 0) {
            int w = (int)((float)val / max * Width);
            if (w > 16) {
                Rectangle r = isRTL ? new Rectangle(Width - w, 0, w - 1, Height - 1) : new Rectangle(0, 0, w - 1, Height - 1);
                using (GraphicsPath p = RoundedRect(r, 8)) using (SolidBrush b = new SolidBrush(barColor)) { e.Graphics.FillPath(b, p); }
            }
        }
        if (showText && val > 0) {
            using (SolidBrush tb = new SolidBrush(textColor)) {
                StringFormat sf = new StringFormat(); sf.Alignment = StringAlignment.Center; sf.LineAlignment = StringAlignment.Center;
                e.Graphics.DrawString(val.ToString() + "%", Font, tb, new RectangleF(0, 0, Width, Height), sf);
            }
        }
    }
    private GraphicsPath RoundedRect(Rectangle rect, int radius) {
        GraphicsPath path = new GraphicsPath(); int d = radius * 2;
        path.AddArc(rect.X, rect.Y, d, d, 180, 90); path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90); path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
        path.CloseFigure(); return path;
    }
}
"@ -ReferencedAssemblies System.Windows.Forms, System.Drawing -ErrorAction SilentlyContinue

# Minimal SQLite access through winsqlite3.dll (built into Windows 10/11).
# Shared with the plugin-install background job.
$script:SqliteCs = @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class OtzSql {
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_open_v2")] static extern int open_v2(byte[] f, out IntPtr db, int flags, IntPtr vfs);
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_close_v2")] static extern int close_v2(IntPtr db);
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_busy_timeout")] static extern int busy_timeout(IntPtr db, int ms);
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_prepare_v2")] static extern int prepare_v2(IntPtr db, byte[] sql, int n, out IntPtr st, IntPtr tail);
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_step")] static extern int step(IntPtr st);
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_finalize")] static extern int finalize(IntPtr st);
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_column_count")] static extern int column_count(IntPtr st);
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_column_text")] static extern IntPtr column_text(IntPtr st, int c);
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_bind_text")] static extern int bind_text(IntPtr st, int i, byte[] v, int n, IntPtr destructor);
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_bind_int64")] static extern int bind_int64(IntPtr st, int i, long v);
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_bind_null")] static extern int bind_null(IntPtr st, int i);
    [DllImport("winsqlite3.dll", EntryPoint="sqlite3_errmsg")] static extern IntPtr errmsg(IntPtr db);

    static byte[] U8(string s) { return Encoding.UTF8.GetBytes(s + "\0"); }
    static string Str(IntPtr p) {
        if (p == IntPtr.Zero) return null;
        int len = 0; while (Marshal.ReadByte(p, len) != 0) len++;
        byte[] b = new byte[len]; Marshal.Copy(p, b, 0, len); return Encoding.UTF8.GetString(b);
    }
    static string Err(IntPtr db) { return Str(errmsg(db)) ?? "sqlite error"; }

    public static IntPtr Open(string path, bool write) {
        IntPtr db;
        int rc = open_v2(U8(path), out db, write ? 0x2 : 0x1, IntPtr.Zero);
        if (rc != 0) { string m = db != IntPtr.Zero ? Err(db) : "open failed"; if (db != IntPtr.Zero) close_v2(db); throw new Exception(m); }
        busy_timeout(db, 8000);
        return db;
    }
    public static void Close(IntPtr db) { if (db != IntPtr.Zero) close_v2(db); }

    static IntPtr Prepare(IntPtr db, string sql, object[] args) {
        IntPtr st;
        if (prepare_v2(db, U8(sql), -1, out st, IntPtr.Zero) != 0) throw new Exception(Err(db));
        if (args != null) {
            for (int i = 0; i < args.Length; i++) {
                object a = args[i];
                if (a == null) bind_null(st, i + 1);
                else if (a is int || a is long || a is short) bind_int64(st, i + 1, Convert.ToInt64(a));
                else if (a is bool) bind_int64(st, i + 1, ((bool)a) ? 1 : 0);
                else { byte[] b = Encoding.UTF8.GetBytes(a.ToString()); bind_text(st, i + 1, b, b.Length, new IntPtr(-1)); }
            }
        }
        return st;
    }
    public static List<string[]> Query(IntPtr db, string sql, object[] args) {
        var rows = new List<string[]>();
        IntPtr st = Prepare(db, sql, args);
        try {
            while (true) {
                int rc = step(st);
                if (rc == 100) {
                    int n = column_count(st); var r = new string[n];
                    for (int c = 0; c < n; c++) r[c] = Str(column_text(st, c));
                    rows.Add(r);
                } else if (rc == 101) break;
                else throw new Exception(Err(db));
            }
        } finally { finalize(st); }
        return rows;
    }
    public static void Exec(IntPtr db, string sql, object[] args) {
        IntPtr st = Prepare(db, sql, args);
        try {
            int rc;
            while ((rc = step(st)) == 100) { }
            if (rc != 101) throw new Exception(Err(db));
        } finally { finalize(st); }
    }
    public static string Scalar(string path, string sql) {
        IntPtr db = IntPtr.Zero;
        try {
            db = Open(path, false);
            var rows = Query(db, sql, null);
            return (rows.Count > 0 && rows[0].Length > 0) ? rows[0][0] : null;
        } catch { return null; }
        finally { Close(db); }
    }
}
"@
try { Add-Type -TypeDefinition $script:SqliteCs -ErrorAction Stop } catch { }

try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class UxTheme {
    [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)] public static extern int SetWindowTheme(IntPtr hWnd, string app, string idList);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr w, IntPtr l);
    public static void Apply(IntPtr h, string theme) { try { SetWindowTheme(h, theme, null); } catch { } }
    public static IntPtr ListViewHeader(IntPtr lv) { return SendMessage(lv, 0x101F, IntPtr.Zero, IntPtr.Zero); }
}
"@ -ErrorAction SilentlyContinue
} catch { }

# פסי גלילה, כותרות עמודות ותיבות בחירה מצוירים בידי Windows - מבקשים את העיצוב הכהה/הבהיר שלו
function Set-NativeTheme($c) {
    if (-not $c.IsHandleCreated) { return }
    try {
        $dark = $script:DarkMode
        if ($c -is [Windows.Forms.ComboBox]) { [UxTheme]::Apply($c.Handle, $(if ($dark) { 'DarkMode_CFD' } else { [NullString]::Value })) }
        elseif ($c -is [Windows.Forms.ListView]) {
            [UxTheme]::Apply($c.Handle, $(if ($dark) { 'DarkMode_Explorer' } else { 'Explorer' }))
            $hdr = [UxTheme]::ListViewHeader($c.Handle)
            if ($hdr -ne [IntPtr]::Zero) { [UxTheme]::Apply($hdr, $(if ($dark) { 'DarkMode_ItemsView' } else { [NullString]::Value })) }
        }
        else { [UxTheme]::Apply($c.Handle, $(if ($dark) { 'DarkMode_Explorer' } else { [NullString]::Value })) }
        $c.Invalidate($true)
    } catch { }
}

# ---------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------
$script:DarkMode = $false
$script:DarkModeManual = $false
function Get-WindowsDarkMode {
    try { return ((Get-ItemPropertyValue -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'AppsUseLightTheme' -ErrorAction Stop) -eq 0) } catch { return $false }
}
$script:DarkMode = Get-WindowsDarkMode

$script:LightTheme = @{
    BgColor = [Drawing.Color]::FromArgb(255,248,244); HeaderColor = [Drawing.Color]::FromArgb(249,236,223)
    CardColor = [Drawing.Color]::White; TextColor = [Drawing.Color]::FromArgb(100,70,50)
    TextLight = [Drawing.Color]::FromArgb(150,120,100); Divider = [Drawing.Color]::FromArgb(240,230,220)
    BtnBg = [Drawing.Color]::FromArgb(255,241,229); BtnBorder = [Drawing.Color]::FromArgb(235,215,195)
    ProgBg = [Drawing.Color]::FromArgb(240,230,220); ProgBar = [Drawing.Color]::FromArgb(200,160,120)
    Good = [Drawing.Color]::FromArgb(40,130,60); Warn = [Drawing.Color]::FromArgb(190,90,20)
}
$script:DarkTheme = @{
    BgColor = [Drawing.Color]::FromArgb(30,30,35); HeaderColor = [Drawing.Color]::FromArgb(40,40,48)
    CardColor = [Drawing.Color]::FromArgb(45,45,52); TextColor = [Drawing.Color]::FromArgb(230,225,220)
    TextLight = [Drawing.Color]::FromArgb(160,155,150); Divider = [Drawing.Color]::FromArgb(60,60,68)
    BtnBg = [Drawing.Color]::FromArgb(55,55,65); BtnBorder = [Drawing.Color]::FromArgb(80,80,95)
    ProgBg = [Drawing.Color]::FromArgb(35,35,40); ProgBar = [Drawing.Color]::FromArgb(110,100,95)
    Good = [Drawing.Color]::FromArgb(110,200,120); Warn = [Drawing.Color]::FromArgb(240,160,80)
}
function Get-Theme { if ($script:DarkMode) { $script:DarkTheme } else { $script:LightTheme } }

# Tags drive the generic theming: card, header, divider, strong, good, warn, network
function Apply-ThemeTo($root) {
    $t = Get-Theme
    foreach ($c in $root.Controls) {
        $tag = [string]$c.Tag
        if ($c -is [ModernProgressBar]) { $c.BgColor = $t.ProgBg; $c.BarColor = $t.ProgBar; $c.TextColor = $t.TextColor }
        elseif ($c -is [Windows.Forms.Button]) { $c.BackColor = $t.BtnBg; $c.ForeColor = $t.TextColor; $c.FlatAppearance.BorderColor = $t.BtnBorder }
        elseif ($c -is [Windows.Forms.ComboBox] -or $c -is [Windows.Forms.TextBoxBase] -or $c -is [Windows.Forms.ListView] -or $c -is [Windows.Forms.ListBox]) { $c.BackColor = $t.CardColor; $c.ForeColor = $t.TextColor; Set-NativeTheme $c }
        elseif ($c -is [Windows.Forms.RadioButton] -or $c -is [Windows.Forms.CheckBox]) { $c.ForeColor = $t.TextColor; Set-NativeTheme $c }
        elseif ($c -is [Windows.Forms.Panel]) {
            if ($tag -eq 'card') { $c.BackColor = $t.CardColor } elseif ($tag -eq 'header') { $c.BackColor = $t.HeaderColor }
        }
        elseif ($c -is [Windows.Forms.Label]) {
            switch ($tag) {
                'divider' { $c.BackColor = $t.Divider }
                'strong'  { $c.ForeColor = $t.TextColor }
                'good'    { $c.ForeColor = $t.Good }
                'warn'    { $c.ForeColor = $t.Warn }
                'network' { }
                default   { $c.ForeColor = $t.TextLight }
            }
        }
        if ($c.Controls.Count -gt 0) { Apply-ThemeTo $c }
    }
}
function Apply-FormTheme($f) {
    $t = Get-Theme
    $f.BackColor = $t.BgColor
    Apply-ThemeTo $f
    Set-NativeTheme $f   # פסי הגלילה של החלון עצמו (כשהוא קטן מדי)
    try {
        if ($f.IsHandleCreated) {
            [DwmHelper]::SetDarkMode($f.Handle, $script:DarkMode)
            if ($script:DarkMode) { [DwmHelper]::SetTitleBarColor($f.Handle, 0x00302828) } else { [DwmHelper]::SetTitleBarColor($f.Handle, 0x00f4f8ff) }
        }
    } catch { }
    $f.Refresh()
}

# ---------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------
function Show-RTLMessageBox {
    param([string]$Message, [string]$Title = 'אוצריא',
          [Windows.Forms.MessageBoxButtons]$Buttons = [Windows.Forms.MessageBoxButtons]::OK,
          [Windows.Forms.MessageBoxIcon]$Icon = [Windows.Forms.MessageBoxIcon]::Information)
    $opt = [Windows.Forms.MessageBoxOptions]::RtlReading -bor [Windows.Forms.MessageBoxOptions]::RightAlign
    return [Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons, $Icon, [Windows.Forms.MessageBoxDefaultButton]::Button1, $opt)
}

function Format-Size([double]$b) {
    if ($b -ge 1GB) { return ('{0:N2} ג"ב' -f ($b / 1GB)) }
    if ($b -ge 1MB) { return ('{0:N1} מ"ב' -f ($b / 1MB)) }
    if ($b -ge 1KB) { return ('{0:N0} ק"ב' -f ($b / 1KB)) }
    return ('{0} בתים' -f [long]$b)
}

function ConvertTo-LocalDate($v) {
    if (-not $v) { return $null }
    if ($v -is [datetime]) { return $v.ToLocalTime() }
    try { return [DateTimeOffset]::Parse([string]$v, [Globalization.CultureInfo]::InvariantCulture).LocalDateTime } catch { return $null }
}
function Format-Date($v) { $d = ConvertTo-LocalDate $v; if ($d) { return $d.ToString('dd/MM/yyyy') } else { return '' } }

# Compares dotted versions numerically ("0.9.98.3" vs "0.9.98"); returns -1/0/1
function Compare-Ver([string]$a, [string]$b) {
    $pa = @(([string]$a -split '[-+ ]')[0] -split '\.'); $pb = @(([string]$b -split '[-+ ]')[0] -split '\.')
    $n = [Math]::Max($pa.Count, $pb.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $x = 0L; $y = 0L
        if ($i -lt $pa.Count) { [void][long]::TryParse(($pa[$i] -replace '\D.*$',''), [ref]$x) }
        if ($i -lt $pb.Count) { [void][long]::TryParse(($pb[$i] -replace '\D.*$',''), [ref]$y) }
        if ($x -lt $y) { return -1 }; if ($x -gt $y) { return 1 }
    }
    return 0
}

function Get-WebText([string]$Url) {
    $wc = New-Object System.Net.WebClient
    $wc.Encoding = [Text.Encoding]::UTF8
    $wc.Headers.Add('User-Agent', $script:UserAgent)
    if ($Url -like 'https://api.github.com/*') { $wc.Headers.Add('Accept', 'application/vnd.github+json') }
    try { return $wc.DownloadString($Url) }
    catch {
        $ex = $_.Exception
        while ($ex -and -not ($ex -is [System.Net.WebException])) { $ex = $ex.InnerException }
        $code = 0; if ($ex -and $ex.Response) { try { $code = [int]$ex.Response.StatusCode } catch { } }
        if (($code -eq 403 -or $code -eq 429) -and $Url -like 'https://api.github.com/*') {
            throw 'חרגת ממגבלת הבקשות של GitHub (60 בשעה). נסה שוב מאוחר יותר.'
        }
        if ($ex) { throw $ex.Message }
        throw
    }
    finally { $wc.Dispose() }
}

function Read-TextFileTrim([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try { $s = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8); return $s.Trim().TrimStart([char]0xFEFF).Trim() } catch { return $null }
}
function Read-JsonFile([string]$Path) {
    $s = Read-TextFileTrim $Path
    if (-not $s) { return $null }
    try { return ConvertFrom-Json -InputObject $s } catch { return $null }
}

function Get-WinArch {
    $a = $env:PROCESSOR_ARCHITEW6432; if (-not $a) { $a = $env:PROCESSOR_ARCHITECTURE }
    if ($a -eq 'ARM64') { return 'arm64' } else { return 'x64' }
}

function Test-OtzariaRunning { return [bool](Get-Process -Name 'otzaria' -ErrorAction SilentlyContinue) }

# Asks the user to close Otzaria; returns $true when it is no longer running
function Confirm-CloseOtzaria([string]$Why) {
    if (-not (Test-OtzariaRunning)) { return $true }
    $r = Show-RTLMessageBox "אוצריא פתוחה כעת.`n$Why`n`nלסגור את אוצריא עכשיו?`n(יש לשמור קודם עבודה פתוחה)" 'אוצריא פתוחה' 'YesNo' 'Warning'
    if ($r -ne 'Yes') { return $false }
    Get-Process -Name 'otzaria' -ErrorAction SilentlyContinue | ForEach-Object { try { $_.CloseMainWindow() | Out-Null } catch { } }
    for ($i = 0; $i -lt 30 -and (Test-OtzariaRunning); $i++) { Start-Sleep -Milliseconds 200; [Windows.Forms.Application]::DoEvents() }
    if (Test-OtzariaRunning) { Get-Process -Name 'otzaria' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue; Start-Sleep -Milliseconds 800 }
    return -not (Test-OtzariaRunning)
}

function Convert-MarkdownToText($md) {
    if (-not $md) { return '' }
    $t = [string]$md
    $t = $t -replace '<[^>]+>', ''
    $t = $t -replace '(?m)^\s*>\s?', ''
    $t = $t -replace '(?m)^#{1,6}\s*(.+)$', '## $1'
    $t = $t -replace '__([^_]+)__', '$1'
    $t = $t -replace '(?m)^\s*[-\*]\s+', '• '
    $t = $t -replace '```[^\n]*\n?', ''
    $t = $t -replace '`([^`]+)`', '$1'
    $t = $t -replace '\[([^\]]+)\]\([^\)]+\)', '$1'
    $t = $t -replace "`r?`n", "`r`n"
    $t = $t -replace '(\r\n){3,}', "`r`n`r`n"
    return $t.Trim()
}

# ---------------------------------------------------------------------
# Display scaling. The layout is designed for 96 DPI. The whole window (positions,
# sizes and fonts) is scaled by one factor: the display scaling, capped so the
# window still fits the screen. Prevents clipped/overlapping text on 125%/150% screens.
# ---------------------------------------------------------------------
function Get-SystemDpi {
    $g = [Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
    try { return [double]$g.DpiX } finally { $g.Dispose() }
}
function Set-ControlFontScale($ctl, [double]$k) {
    foreach ($c in $ctl.Controls) { Set-ControlFontScale $c $k }   # children first: avoid double scaling of inherited fonts
    $fnt = $ctl.Font
    $ctl.Font = New-Object Drawing.Font($fnt.FontFamily, [float]($fnt.SizeInPoints * $k), $fnt.Style)
}
# $DesignClient: גודל אזור התוכן המקורי. Windows חותך חלון שגבוה מהמסך עוד לפני ההתאמה,
# ולכן הגודל נקבע מחדש לפיו. אם גם אחרי ההקטנה החלון גבוה מהמסך (מסך קטן, למשל 1024x768)
# - הגובה מוגבל לגובה המסך ויש פס גלילה, במקום להקטין עד שהטקסט לא קריא.
function Set-UiScale($f, [int]$designW, [int]$designH, [double]$MinScale = 0.75, $DesignClient = $null) {
    $dpi = Get-SystemDpi
    $wa = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $s = [Math]::Min($dpi / 96.0, [Math]::Min(($wa.Height - 50) / [double]$designH, ($wa.Width - 30) / [double]$designW))
    $s = [Math]::Max($s, $MinScale)
    Set-ControlFontScale $f ($s * 96.0 / $dpi)
    if ([Math]::Abs($s - 1.0) -gt 0.01) { $f.Scale((New-Object Drawing.SizeF([float]$s, [float]$s))) }
    if ($DesignClient) {
        $wantW = [int]($DesignClient.Width * $s); $wantH = [int]($DesignClient.Height * $s)
        $chromeH = [Math]::Max(30, $f.Height - $f.ClientSize.Height)
        $maxH = $wa.Height - $chromeH - 8
        if ($wantH -gt $maxH) {
            $f.AutoScroll = $true
            $f.AutoScrollMinSize = New-Object Drawing.Size($wantW, $wantH)
            $f.ClientSize = New-Object Drawing.Size(($wantW + [Windows.Forms.SystemInformation]::VerticalScrollBarWidth), $maxH)
        } else {
            $f.ClientSize = New-Object Drawing.Size($wantW, $wantH)
        }
    }
    return $s
}

# טקסט עם סימון פשוט: **מודגש**, ושורה שמתחילה ב-"## " היא כותרת
function Write-RichText($rtb, [string]$Text) {
    $base = $rtb.Font
    $bold = New-Object Drawing.Font($base, [Drawing.FontStyle]::Bold)
    $head = New-Object Drawing.Font($base.FontFamily, [float]($base.Size * 1.2), [Drawing.FontStyle]::Bold)
    $rtb.Clear()
    $first = $true
    foreach ($line in ($Text -split "`r?`n")) {
        if (-not $first) { $rtb.SelectionStart = $rtb.TextLength; $rtb.SelectionFont = $base; $rtb.SelectedText = "`n" }
        $first = $false
        if ($line.StartsWith('## ')) {
            $rtb.SelectionStart = $rtb.TextLength; $rtb.SelectionFont = $head; $rtb.SelectedText = $line.Substring(3)
            continue
        }
        # סימן כיוון מימין-לשמאל (RLM) - שורה שמתחילה באנגלית לא תשבש את מיקום הנקודה
        $rlm = [string][char]0x200F
        $line = $rlm + ($line -replace '^(\s*•\s*)', ('$1' + $rlm))
        $parts = $line -split '\*\*'
        for ($k = 0; $k -lt $parts.Count; $k++) {
            if ($parts[$k] -eq '') { continue }
            $rtb.SelectionStart = $rtb.TextLength
            $rtb.SelectionFont = $(if ($k % 2 -eq 1) { $bold } else { $base })
            $rtb.SelectedText = $parts[$k]
        }
    }
    $rtb.SelectionStart = 0; $rtb.SelectionLength = 0
}

function Show-TextDialog([string]$Title, [string]$Text) {
    $dlg = New-Object Windows.Forms.Form
    $dlg.Text = $Title; $dlg.Size = New-Object Drawing.Size(640, 560); $dlg.StartPosition = 'CenterParent'
    $dlg.RightToLeft = 'Yes'; $dlg.RightToLeftLayout = $true; $dlg.MinimizeBox = $false
    $dlg.Font = New-Object Drawing.Font('Segoe UI', 10)
    $tb = New-Object Windows.Forms.RichTextBox
    $tb.ReadOnly = $true; $tb.ScrollBars = 'Vertical'; $tb.WordWrap = $true; $tb.DetectUrls = $false
    $tb.Dock = 'Fill'; $tb.Font = New-Object Drawing.Font('Segoe UI', 11); $tb.BorderStyle = 'None'; $tb.RightToLeft = 'Yes'
    $pnl = New-Object Windows.Forms.Panel; $pnl.Dock = 'Bottom'; $pnl.Height = 52
    $btn = New-Object Windows.Forms.Button; $btn.Text = 'סגור'; $btn.Size = New-Object Drawing.Size(110, 34); $btn.FlatStyle = 'Flat'
    $btn.Location = New-Object Drawing.Point(255, 9); $btn.Add_Click({ $dlg.Close() })
    $pnl.Controls.Add($btn)
    $wrap = New-Object Windows.Forms.Panel; $wrap.Dock = 'Fill'; $wrap.Padding = New-Object Windows.Forms.Padding(14); $wrap.Tag = 'card'
    $wrap.Controls.Add($tb)
    $dlg.Controls.Add($wrap); $dlg.Controls.Add($pnl)
    $dlg.Add_Shown({ Apply-FormTheme $dlg; Write-RichText $tb $Text; $btn.Focus() })
    $dlg.AutoScaleMode = 'None'
    [void](Set-UiScale $dlg 640 560)
    [void]$dlg.ShowDialog()
}

# ---------------------------------------------------------------------
# Local detection: installed app, library, downloaded files
# ---------------------------------------------------------------------
function Get-OtzariaInstall {
    $caps = Read-JsonFile (Join-Path $env:APPDATA 'otzaria\otzaria_capabilities.json')
    $dirs = @()
    $hives = @(
        @{ P = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\' },
        @{ P = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\' },
        @{ P = 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\' })
    foreach ($h in $hives) {
        try {
            $p = Get-ItemProperty -LiteralPath ($h.P + $script:AppId) -ErrorAction Stop
            if ($p.InstallLocation) { $dirs += $p.InstallLocation.TrimEnd('\') }
        } catch { }
    }
    if ($caps -and $caps.exePath) { $dirs += (Split-Path -Parent ([string]$caps.exePath)) }
    $dirs += (Join-Path $env:ProgramFiles 'Otzaria'), (Join-Path $env:LOCALAPPDATA 'Programs\Otzaria')
    foreach ($d in (Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object { $_.Root -match '^[A-Z]:\\$' })) {
        $dirs += (Join-Path $d.Root 'אוצריא')
    }
    foreach ($dir in $dirs) {
        if (-not $dir) { continue }
        $exe = Join-Path $dir 'otzaria.exe'
        if (-not (Test-Path -LiteralPath $exe)) { continue }
        $fv = ''; try { $fv = [string](Get-Item -LiteralPath $exe).VersionInfo.FileVersion } catch { }
        $build = $null; if ($fv -match '\+(\d+)') { $build = [int]$matches[1] }
        $ver = ($fv -split '\+')[0]
        if ($caps -and $caps.appVersion -and $build -and ([string]$caps.buildNumber -eq [string]$build)) { $ver = [string]$caps.appVersion }
        elseif ($build) {
            # מספר הבנייה של אוצריא: <גירסה בלי נקודות><שתי ספרות תיקון>, למשל 99803 = 0.9.98 תיקון 3
            $vp = @($ver -split '\.')
            if ($vp.Count -eq 3 -and [int][Math]::Floor($build / 100) -eq ([int]$vp[1] * 100 + [int]$vp[2]) -and ($build % 100) -gt 0) { $ver = "$ver.$($build % 100)" }
        }
        $portable = Test-Path -LiteralPath (Join-Path $dir 'portable.marker')
        $system = (Test-Path -LiteralPath (Join-Path $dir 'system_install.marker')) -or
                  ($dir.ToLower().StartsWith($env:ProgramFiles.ToLower())) -or
                  (${env:ProgramFiles(x86)} -and $dir.ToLower().StartsWith(${env:ProgramFiles(x86)}.ToLower()))
        $mode = if ($portable) { 'portable' } elseif ($system) { 'system' } else { 'user' }
        $modeText = switch ($mode) { 'portable' { 'גרסה ניידת' } 'system' { 'לכל המשתמשים' } default { 'למשתמש הנוכחי' } }
        return [pscustomobject]@{ Dir = $dir; Exe = $exe; FileVersion = $fv; Version = $ver; Build = $build; Mode = $mode; ModeText = $modeText; Caps = $caps }
    }
    return $null
}

function Get-DataRoot {
    if ($script:Install -and $script:Install.Mode -eq 'portable') { return (Join-Path $script:Install.Dir 'otzaria_data') }
    return (Join-Path $env:APPDATA 'otzaria')
}

function Get-LocalLibrary {
    $cands = New-Object System.Collections.Generic.List[string]
    $caps = if ($script:Install) { $script:Install.Caps } else { $null }
    if ($caps -and $caps.libraryDbPath) { $cands.Add((Split-Path -Parent ([string]$caps.libraryDbPath))) }
    $roots = @((Get-DataRoot), (Join-Path $env:ProgramData 'otzaria'))
    foreach ($r in $roots) {
        foreach ($f in 'library_database_path.txt', 'library_path.txt') {
            $v = Read-TextFileTrim (Join-Path $r $f)
            if ($v) { $cands.Add($v); $cands.Add((Join-Path $v 'books')) }
        }
    }
    if ($script:Install -and $script:Install.Mode -eq 'system') { $cands.Add((Join-Path $env:ProgramData 'otzaria\books')) }
    $cands.Add((Join-Path (Get-DataRoot) 'books'))
    $cands.Add((Join-Path $env:ProgramData 'otzaria\books'))
    foreach ($c in $cands) {
        if (-not $c) { continue }
        $db = Join-Path $c 'seforim.db'
        if (Test-Path -LiteralPath $db) {
            $ver = [OtzSql]::Scalar($db, "SELECT value FROM schema_meta WHERE key='db_version'")
            $schema = [OtzSql]::Scalar($db, "SELECT value FROM schema_meta WHERE key='db_schema_version'")
            return [pscustomobject]@{ Dir = $c; Db = $db; Version = $ver; Schema = $schema }
        }
    }
    return $null
}

$script:InstallerPattern = '^otzaria-([\d\.]+)-windows(_arm64)?(-full)?\.exe$'
function Get-LocalInstallers {
    @(Get-ChildItem -LiteralPath $script:WorkDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match $script:InstallerPattern -or $_.Name -match '^otzaria-[\d\.]+-windows(_arm64)?-portable\.zip$' } |
        Sort-Object LastWriteTime -Descending)
}

# Library parts in the folder grouped by app version: @{ '0.9.98' = @(files) }
function Get-LocalLibraryParts {
    $g = @{}
    foreach ($f in (Get-ChildItem -LiteralPath $script:WorkDir -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -match '^otzaria-([\d\.]+)-library\.tar\.zst\.part-\d+$') {
            if (-not $g.ContainsKey($matches[1])) { $g[$matches[1]] = @() }
            $g[$matches[1]] += $f
        }
    }
    return $g
}

# ---------------------------------------------------------------------
# Remote data: app releases, library releases
# ---------------------------------------------------------------------
function New-AssetUrl([string]$Repo, [string]$Tag, [string]$Name) {
    return "https://github.com/$Repo/releases/download/$([uri]::EscapeDataString($Tag))/$([uri]::EscapeDataString($Name))"
}

function Convert-ManifestComponent($c, [string]$tag) {
    $a = @($c.assets)[0]
    if (-not $a) { return $null }
    $repo = if ($a.repository) { [string]$a.repository } else { $script:AppRepo }
    $atag = if ($a.releaseTag) { [string]$a.releaseTag } else { $tag }
    $parts = @()
    foreach ($p in @($a.parts)) {
        if (-not $p) { continue }
        $parts += [pscustomobject]@{ Name = [string]$p.name; Size = [long]$p.size; Sha256 = [string]$p.sha256; Url = (New-AssetUrl $repo $atag $p.name) }
    }
    $manUrl = $null; if ($a.manifestAsset) { $manUrl = New-AssetUrl $repo $atag $a.manifestAsset }
    return [pscustomobject]@{
        Id = [string]$c.id; Kind = [string]$a.kind; Name = [string]$a.name; Size = [long]$a.size
        Sha256 = [string]$a.sha256; Url = (New-AssetUrl $repo $atag $a.name); Parts = $parts
        ManifestAsset = [string]$a.manifestAsset; ManifestUrl = $manUrl
    }
}

# Fallback for releases without otzaria-release-manifest.json
function Get-ComponentsFromAssets($r) {
    $comps = @{}
    foreach ($a in @($r.assets)) {
        $n = [string]$a.name; $id = $null
        if ($n -match '^otzaria-[\d\.]+-windows\.exe$') { $id = 'otzaria-windows-x64' }
        elseif ($n -match '^otzaria-[\d\.]+-windows_arm64\.exe$') { $id = 'otzaria-windows-arm64' }
        elseif ($n -match '^otzaria-[\d\.]+-windows-full\.exe$') { $id = 'otzaria-windows-full' }
        elseif ($n -match '^otzaria-[\d\.]+-windows_arm64-full\.exe$') { $id = 'otzaria-windows-full-arm64' }
        elseif ($n -eq 'otzaria-windows.zip') { $id = 'otzaria-windows-portable-x64' }
        elseif ($n -eq 'otzaria-windows_arm64.zip') { $id = 'otzaria-windows-portable-arm64' }
        if (-not $id) { continue }
        $sha = ''; if ($a.digest -and ([string]$a.digest) -like 'sha256:*') { $sha = ([string]$a.digest).Substring(7) }
        $comps[$id] = [pscustomobject]@{ Id = $id; Kind = 'single'; Name = $n; Size = [long]$a.size; Sha256 = $sha; Url = [string]$a.browser_download_url; Parts = @(); ManifestAsset = ''; ManifestUrl = $null }
    }
    return $comps
}

function Convert-Release($r) {
    $tag = [string]$r.tag_name
    $build = $null
    foreach ($a in @($r.assets)) { if ([string]$a.name -match '^otzaria-[\d\.]+\+(\d+)-linux\.deb$') { $build = [int]$matches[1]; break } }
    $comps = @{}
    $man = @($r.assets) | Where-Object { $_.name -eq 'otzaria-release-manifest.json' } | Select-Object -First 1
    if ($man) {
        try {
            $m = ConvertFrom-Json -InputObject (Get-WebText ([string]$man.browser_download_url))
            foreach ($c in @($m.components)) { $cc = Convert-ManifestComponent $c $tag; if ($cc) { $comps[$cc.Id] = $cc } }
        } catch { }
    }
    if ($comps.Count -eq 0) { $comps = Get-ComponentsFromAssets $r }
    return [pscustomobject]@{
        Tag = $tag; Version = ($tag -split '\+')[0]; Build = $build; Prerelease = [bool]$r.prerelease
        Body = [string]$r.body; Date = (ConvertTo-LocalDate $r.published_at); Components = $comps
    }
}

function Get-AppReleases {
    $list = ConvertFrom-Json -InputObject (Get-WebText "https://api.github.com/repos/$($script:AppRepo)/releases?per_page=8")
    $stable = $null; $pre = $null
    foreach ($r in @($list)) {
        if ($r.draft) { continue }
        if (-not $r.prerelease) { if (-not $stable) { $stable = $r } }
        elseif (-not $pre) { $pre = $r }
    }
    $res = @()
    if ($stable) { $res += (Convert-Release $stable) }
    if ($pre -and (-not $stable -or (ConvertTo-LocalDate $pre.published_at) -gt (ConvertTo-LocalDate $stable.published_at))) { $res += (Convert-Release $pre) }
    return $res
}

function Get-LatestLibraryRelease {
    $r = ConvertFrom-Json -InputObject (Get-WebText "https://api.github.com/repos/$($script:LibRepo)/releases/latest")
    $ver = $null; if ([string]$r.tag_name -match '^v(\d+)') { $ver = $matches[1] }
    $changes = @($r.assets) | Where-Object { $_.name -eq 'library-changes.md' } | Select-Object -First 1
    return [pscustomobject]@{ Tag = [string]$r.tag_name; Version = $ver; Date = (ConvertTo-LocalDate $r.published_at); ChangesUrl = $(if ($changes) { [string]$changes.browser_download_url } else { $null }) }
}

# ---------------------------------------------------------------------
# Background download worker (runs in a separate process via Start-Job)
# Plan JSON: { Items: [{Url,Path,Size,Sha256}], Assemble: [{Output,Parts,Sha256}] }
# Output lines: PROGRESS:<bytes>:<pct> | STAGE:<text> | DONE | STOPPED | ERROR:<msg>
# ---------------------------------------------------------------------
$script:DownloadWorker = {
    param([string]$PlanJson, [string]$StopFlag)
    $ErrorActionPreference = 'Stop'
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
    function HashOf([string]$p) { return (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToLowerInvariant() }
    try {
        $plan = ConvertFrom-Json -InputObject $PlanJson
        $items = @($plan.Items | Where-Object { $_ })
        $total = 0L; foreach ($it in $items) { $total += [long]$it.Size }
        $done = 0L
        $last = [DateTime]::MinValue
        $idx = 0
        foreach ($it in $items) {
            $idx++
            $final = [string]$it.Path; $tmp = "$final.download"; $size = [long]$it.Size; $sha = ([string]$it.Sha256).ToLowerInvariant()
            $leaf = Split-Path -Leaf $final
            if (Test-Path -LiteralPath $final) {
                $len = (Get-Item -LiteralPath $final).Length
                $ok = ($size -le 0 -or $len -eq $size)
                if ($ok -and $sha) { Write-Output "STAGE:בודק קובץ קיים: $leaf"; $ok = ((HashOf $final) -eq $sha) }
                if ($ok) { $done += $len; Write-Output "PROGRESS:$($done):$([int][Math]::Min(100, [Math]::Floor($done * 100.0 / [Math]::Max([double]1, [double]$total))))"; continue }
                Remove-Item -LiteralPath $final -Force
            }
            $start = 0L
            if (Test-Path -LiteralPath $tmp) {
                $start = (Get-Item -LiteralPath $tmp).Length
                if ($size -gt 0 -and $start -gt $size) { Remove-Item -LiteralPath $tmp -Force; $start = 0L }
            }
            if ($size -le 0 -or $start -lt $size) {
                Write-Output "STAGE:מוריד ($idx/$($items.Count)): $leaf"
                $req = [Net.HttpWebRequest]::Create([string]$it.Url)
                $req.UserAgent = 'OtzariaUpdater/2.0'; $req.Timeout = 60000; $req.ReadWriteTimeout = 120000; $req.AllowAutoRedirect = $true
                if ($start -gt 0) { $req.AddRange($start) }
                $resp = $req.GetResponse()
                try {
                    if ($start -gt 0 -and [int]$resp.StatusCode -ne 206) { $start = 0L }
                    $mode = if ($start -gt 0) { [IO.FileMode]::Append } else { [IO.FileMode]::Create }
                    $fs = [IO.File]::Open($tmp, $mode, [IO.FileAccess]::Write)
                    try {
                        $stream = $resp.GetResponseStream()
                        $buf = New-Object byte[] 262144
                        $cur = $start
                        while (($n = $stream.Read($buf, 0, $buf.Length)) -gt 0) {
                            $fs.Write($buf, 0, $n); $cur += $n
                            if (([DateTime]::Now - $last).TotalMilliseconds -ge 300) {
                                $last = [DateTime]::Now
                                if (Test-Path -LiteralPath $StopFlag) {
                                    Remove-Item -LiteralPath $StopFlag -Force -ErrorAction SilentlyContinue
                                    Write-Output 'STOPPED'
                                    return
                                }
                                $b = $done + $cur
                                Write-Output "PROGRESS:$($b):$([int][Math]::Min(100, [Math]::Floor($b * 100.0 / [Math]::Max([double]1, [double]$total))))"
                            }
                        }
                    } finally { $fs.Close() }
                } finally { $resp.Close() }
            }
            $len = (Get-Item -LiteralPath $tmp).Length
            if ($size -gt 0 -and $len -ne $size) { throw "ההורדה של $leaf לא הושלמה (התקבלו $len מתוך $size בתים). נסה שוב." }
            if ($sha) {
                Write-Output "STAGE:מאמת תקינות: $leaf"
                if ((HashOf $tmp) -ne $sha) { Remove-Item -LiteralPath $tmp -Force; throw "בדיקת התקינות (SHA256) של $leaf נכשלה. הקובץ נמחק - נסה שוב." }
            }
            Move-Item -LiteralPath $tmp -Destination $final -Force
            $done += $len
            Write-Output "PROGRESS:$($done):$([int][Math]::Min(100, [Math]::Floor($done * 100.0 / [Math]::Max([double]1, [double]$total))))"
        }
        foreach ($as in @($plan.Assemble)) {
            if (-not $as) { continue }
            $out = [string]$as.Output; $tmpOut = "$out.download"
            Write-Output "STAGE:מרכיב את $(Split-Path -Leaf $out)..."
            $o = [IO.File]::Create($tmpOut)
            try { foreach ($p in @($as.Parts)) { $in = [IO.File]::OpenRead([string]$p); try { $in.CopyTo($o, 4MB) } finally { $in.Close() } } }
            finally { $o.Close() }
            if ($as.Sha256) {
                Write-Output "STAGE:מאמת את $(Split-Path -Leaf $out)..."
                if ((HashOf $tmpOut) -ne ([string]$as.Sha256).ToLowerInvariant()) { Remove-Item -LiteralPath $tmpOut -Force; throw 'בדיקת התקינות של הקובץ המורכב נכשלה.' }
            }
            Move-Item -LiteralPath $tmpOut -Destination $out -Force
            foreach ($p in @($as.Parts)) { Remove-Item -LiteralPath ([string]$p) -Force -ErrorAction SilentlyContinue }
        }
        Write-Output 'DONE'
    } catch {
        Write-Output ("ERROR:" + $_.Exception.Message)
    }
}

$script:Active = $null       # current running download
$script:Paused = @{}         # Kind -> plan, for resume
$script:CancelPending = @{}

function Get-StopFlag([string]$Kind) { return (Join-Path $env:TEMP "otzaria_updater_stop_$Kind.flag") }

function Start-Download {
    param([string]$Kind, [object[]]$Items, [object[]]$Assemble = @(), $Bar, $Label, [scriptblock]$OnDone, $Context)
    if ($script:Active) { Show-RTLMessageBox 'הורדה אחרת כבר פעילה. יש להמתין לסיומה או לעצור אותה.' 'הורדה פעילה' 'OK' 'Warning' | Out-Null; return }
    $total = 0L; foreach ($i in $Items) { $total += [long]$i.Size }
    $plan = @{ Items = @($Items); Assemble = @($Assemble) }
    $json = ConvertTo-Json -InputObject $plan -Depth 6 -Compress
    $flag = Get-StopFlag $Kind
    Remove-Item -LiteralPath $flag -Force -ErrorAction SilentlyContinue
    $job = Start-Job -ScriptBlock $script:DownloadWorker -ArgumentList $json, $flag
    $script:Paused.Remove($Kind)
    $script:CancelPending[$Kind] = $false
    $script:Active = @{ Kind = $Kind; Job = $job; Total = $total; Bar = $Bar; Label = $Label; OnDone = $OnDone; Context = $Context; Items = $Items; Assemble = $Assemble }
    $Bar.Value = 0; $Label.Visible = $true; $Label.Text = 'מתחיל הורדה...'
    try { [TBProg]::SetState($form.Handle, 2) } catch { }
    Update-AllUI
}

function Stop-Download([string]$Kind) {
    if ($script:Active -and $script:Active.Kind -eq $Kind) { 'stop' | Out-File -LiteralPath (Get-StopFlag $Kind) -Force }
}

function Remove-PlanTemps($plan) {
    foreach ($i in @($plan.Items)) { if ($i) { Remove-Item -LiteralPath ("$($i.Path).download") -Force -ErrorAction SilentlyContinue } }
    foreach ($a in @($plan.Assemble)) { if ($a) { Remove-Item -LiteralPath ("$($a.Output).download") -Force -ErrorAction SilentlyContinue } }
}

function Cancel-Download([string]$Kind) {
    $r = Show-RTLMessageBox "לבטל את ההורדה?`nקבצים חלקיים יימחקו (קבצים שכבר הושלמו יישארו)." 'ביטול הורדה' 'YesNo' 'Question'
    if ($r -ne 'Yes') { return }
    if ($script:Active -and $script:Active.Kind -eq $Kind) {
        $script:CancelPending[$Kind] = $true
        Stop-Download $Kind
    } elseif ($script:Paused.ContainsKey($Kind)) {
        Remove-PlanTemps $script:Paused[$Kind]
        $script:Paused.Remove($Kind)
        Update-AllUI
    }
}

function Complete-Download([string]$Result, [string]$Message) {
    $a = $script:Active
    $script:Active = $null
    try { Remove-Job -Job $a.Job -Force -ErrorAction SilentlyContinue } catch { }
    try { [TBProg]::SetState($form.Handle, 0) } catch { }
    switch ($Result) {
        'done' {
            $a.Bar.Value = 100
            $a.Label.Text = 'ההורדה הושלמה'
            if ($a.OnDone) { & $a.OnDone $a.Context }
        }
        'stopped' {
            if ($script:CancelPending[$a.Kind]) {
                Remove-PlanTemps $a; $a.Bar.Value = 0; $a.Label.Text = 'ההורדה בוטלה'
            } else {
                $script:Paused[$a.Kind] = $a; $a.Label.Text = 'ההורדה הושהתה - לחץ "המשך" כדי להמשיך מאותה נקודה'
                try { [TBProg]::SetState($form.Handle, 8) } catch { }
            }
        }
        default {
            $script:Paused[$a.Kind] = $a
            $a.Label.Text = 'שגיאה בהורדה'
            Show-RTLMessageBox "שגיאה בהורדה:`n$Message`n`nאפשר ללחוץ על ""המשך"" כדי לנסות שוב." 'שגיאה' 'OK' 'Error' | Out-Null
        }
    }
    Update-AllUI
}

function Process-ActiveJob {
    $a = $script:Active
    if (-not $a) { return }
    $terminal = $a.Job.State -in @('Completed', 'Failed', 'Stopped')
    $lines = @(Receive-Job -Job $a.Job -ErrorAction SilentlyContinue)
    foreach ($line in $lines) {
        $s = [string]$line
        if ($s -match '^PROGRESS:(\d+):(\d+)$') {
            $bytes = [long]$matches[1]; $pct = [int]$matches[2]
            $a.Bar.Value = $pct
            $a.Label.Text = "הורדו $(Format-Size $bytes) מתוך $(Format-Size $a.Total)"
            try { [TBProg]::SetVal($form.Handle, $pct, 100) } catch { }
        }
        elseif ($s.StartsWith('STAGE:')) { $a.Label.Text = $s.Substring(6) }
        elseif ($s -eq 'DONE') { Complete-Download 'done' ''; return }
        elseif ($s -eq 'STOPPED') { Complete-Download 'stopped' ''; return }
        elseif ($s.StartsWith('ERROR:')) { Complete-Download 'error' $s.Substring(6); return }
    }
    if ($script:Active -and $terminal) {
        $reason = ''; try { $reason = $a.Job.ChildJobs[0].JobStateInfo.Reason.Message } catch { }
        Complete-Download 'error' ("תהליך ההורדה הסתיים באופן לא צפוי. $reason")
    }
}

# ---------------------------------------------------------------------
# Install / extract actions
# ---------------------------------------------------------------------
function Invoke-Installer([string]$Path, [bool]$Silent) {
    if (-not (Test-Path -LiteralPath $Path)) { Show-RTLMessageBox "הקובץ לא נמצא:`n$Path" 'שגיאה' 'OK' 'Error' | Out-Null; return }
    if (-not (Confirm-CloseOtzaria 'יש לסגור אותה לפני ההתקנה.')) { $statusLbl.Text = 'ההתקנה בוטלה - אוצריא עדיין פתוחה'; return }
    if ($Silent) {
        $statusLbl.Text = 'מתקין... (התקנה שקטה)'
        $btnDL.Enabled = $false
        [Windows.Forms.Application]::DoEvents()
        try {
            $p = Start-Process -FilePath $Path -ArgumentList '/SILENT', '/SUPPRESSMSGBOXES', '/NORESTART' -PassThru -WorkingDirectory (Split-Path -Parent $Path)
            while (-not $p.HasExited) { [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 150 }
            if ($p.ExitCode -eq 0) {
                $statusLbl.Text = 'אוצריא הותקנה בהצלחה!'
                Show-RTLMessageBox 'אוצריא הותקנה בהצלחה!' 'ההתקנה הושלמה' | Out-Null
            } else {
                $statusLbl.Text = "ההתקנה הסתיימה עם קוד $($p.ExitCode)"
                Show-RTLMessageBox "ההתקנה לא הושלמה (קוד $($p.ExitCode)).`nנסה להתקין שוב ללא ""התקנה שקטה"" כדי לראות את ההודעות." 'התקנה' 'OK' 'Warning' | Out-Null
            }
        } catch { $statusLbl.Text = "שגיאה: $($_.Exception.Message)" }
        $btnDL.Enabled = $true
        Refresh-LocalInfo
    } else {
        Start-Process -FilePath $Path -WorkingDirectory (Split-Path -Parent $Path)
        $statusLbl.Text = 'תוכנת ההתקנה הופעלה'
    }
}

function Expand-PortableZip([string]$Zip) {
    $fbd = New-Object Windows.Forms.FolderBrowserDialog
    $fbd.Description = 'בחר תיקייה לחילוץ הגרסה הניידת (מומלץ תיקייה ריקה, למשל בכונן נייד)'
    $fbd.ShowNewFolderButton = $true
    if ($fbd.ShowDialog() -ne 'OK') { return }
    $dest = $fbd.SelectedPath
    if ((Test-Path -LiteralPath (Join-Path $dest 'otzaria.exe')) -and (Test-OtzariaRunning)) {
        if (-not (Confirm-CloseOtzaria 'יש לסגור אותה לפני החלפת הקבצים.')) { return }
    }
    $destFull = [IO.Path]::GetFullPath($dest).TrimEnd('\') + '\'
    $zipArc = $null
    try {
        $zipArc = [IO.Compression.ZipFile]::OpenRead($Zip)
        $entries = @($zipArc.Entries); $i = 0
        foreach ($e in $entries) {
            $i++
            if ($i % 10 -eq 0) { $progBar.Value = [int]($i * 100 / $entries.Count); $statusLbl.Text = "מחלץ $i / $($entries.Count)"; [Windows.Forms.Application]::DoEvents() }
            $target = [IO.Path]::GetFullPath((Join-Path $dest $e.FullName))
            if (-not $target.StartsWith($destFull, [StringComparison]::OrdinalIgnoreCase)) { continue }
            if ([string]::IsNullOrEmpty($e.Name)) { New-Item -ItemType Directory -Force -Path $target | Out-Null; continue }
            New-Item -ItemType Directory -Force -Path (Split-Path -Parent $target) | Out-Null
            [IO.Compression.ZipFileExtensions]::ExtractToFile($e, $target, $true)
        }
    } catch {
        Show-RTLMessageBox "שגיאה בחילוץ: $($_.Exception.Message)" 'שגיאה' 'OK' 'Error' | Out-Null; return
    } finally { if ($zipArc) { $zipArc.Dispose() } }
    $progBar.Value = 100
    $r = Show-RTLMessageBox "הקבצים חולצו ל:`n$dest`n`nלהגדיר כגרסה ניידת אמיתית?`n(כן = כל ההגדרות והספרייה יישמרו בתיקייה otzaria_data שליד התוכנה, ולא בפרופיל המשתמש)" 'גרסה ניידת' 'YesNo' 'Question'
    if ($r -eq 'Yes') { try { New-Item -ItemType File -Force -Path (Join-Path $dest 'portable.marker') | Out-Null } catch { } }
    $statusLbl.Text = 'החילוץ הושלם'
    Start-Process explorer.exe -ArgumentList "`"$dest`""
}

function Remove-OldInstallers([string]$KeepName, [string]$KeepVersion) {
    $old = @()
    foreach ($f in (Get-LocalInstallers)) {
        if ($f.Name -eq $KeepName) { continue }
        if ($f.Name -match '^otzaria-([\d\.]+)-' -and (Compare-Ver $matches[1] $KeepVersion) -lt 0) { $old += $f }
    }
    if ($old.Count -eq 0) { return }
    $msg = "נמצאו $($old.Count) קבצי התקנה ישנים בתיקייה:`n`n" + (($old | ForEach-Object { '• ' + $_.Name }) -join "`n") + "`n`nלמחוק אותם?"
    if ((Show-RTLMessageBox $msg 'מחיקת גרסאות ישנות' 'YesNo' 'Question') -eq 'Yes') {
        foreach ($f in $old) { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue }
    }
}

# =====================================================================
# Main form
# =====================================================================
$form = New-Object Windows.Forms.Form
$form.Text = "אוצריא - הורדה ועדכון  (v$($script:UpdaterVersion))"
$form.ClientSize = New-Object Drawing.Size(705, 868)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.RightToLeft = 'Yes'
$form.RightToLeftLayout = $true
$form.AutoScroll = $true
$form.AutoScaleMode = 'None'
$form.Font = New-Object Drawing.Font('Segoe UI', 10)
try { $form.Icon = [Drawing.Icon]::ExtractAssociatedIcon([Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) } catch { }
$toolTip = New-Object Windows.Forms.ToolTip

function New-Lbl($parent, [string]$text, [int]$x, [int]$y, [int]$w, [int]$h = 24, [string]$align = 'MiddleCenter', [string]$tag = '', [float]$size = 10, [string]$font = 'Segoe UI') {
    $l = New-Object Windows.Forms.Label
    $l.Text = $text; $l.Location = New-Object Drawing.Point($x, $y); $l.Size = New-Object Drawing.Size($w, $h)
    $l.TextAlign = $align; $l.Tag = $tag; $l.Font = New-Object Drawing.Font($font, $size)
    $l.AutoEllipsis = $true   # overflowing text ends with "..." instead of being cut
    $parent.Controls.Add($l); return $l
}
function New-Btn($parent, [string]$text, [int]$x, [int]$y, [int]$w, [int]$h = 34, [float]$size = 10) {
    $b = New-Object Windows.Forms.Button
    $b.Text = $text; $b.Location = New-Object Drawing.Point($x, $y); $b.Size = New-Object Drawing.Size($w, $h)
    $b.FlatStyle = 'Flat'; $b.Cursor = 'Hand'; $b.Font = New-Object Drawing.Font('Segoe UI', $size)
    $parent.Controls.Add($b); return $b
}
function New-Card([int]$y, [int]$h) {
    $p = New-Object Windows.Forms.Panel
    $p.Location = New-Object Drawing.Point(25, $y); $p.Size = New-Object Drawing.Size(655, $h); $p.Tag = 'card'
    $form.Controls.Add($p); return $p
}
# Inside cards coordinates are left-based (not mirrored); TextAlign is flipped by RightToLeft:
#  'MiddleRight' renders on the left of the box, 'MiddleLeft' renders on the right.

# --- Header ---
$header = New-Object Windows.Forms.Panel
$header.Dock = 'Top'; $header.Height = 78; $header.Tag = 'header'
$title = New-Lbl $header 'אוצריא' 0 4 705 46 'MiddleCenter' 'strong' 22 'Segoe UI'
$title.Font = New-Object Drawing.Font('Segoe UI', 22, [Drawing.FontStyle]::Bold)
$title.Dock = 'Top'
$sub = New-Lbl $header 'הורדה, התקנה ועדכון של התוכנה, הספרייה והתוספים' 0 50 705 26 'MiddleCenter' '' 10.5
$sub.Dock = 'Bottom'
$form.Controls.Add($header)

$btnRefresh = New-Btn $form '⟳' 640 12 40 40 15
$toolTip.SetToolTip($btnRefresh, 'רענון נתונים')
$btnTheme = New-Btn $form '◐' 594 12 40 40 15
$toolTip.SetToolTip($btnTheme, 'מצב כהה / בהיר')
$btnRefresh.BringToFront(); $btnTheme.BringToFront()

# --- Main card: application ---
$mainCard = New-Card 90 382
$lblSec = New-Lbl $mainCard 'הורדת והתקנת אוצריא' 20 10 615 28 'MiddleCenter' 'strong' 13 'Segoe UI Semibold'
$div1 = New-Lbl $mainCard '' 20 42 615 1 'MiddleCenter' 'divider'

$lblInstT = New-Lbl $mainCard 'מותקן במחשב:' 475 50 160 24 'MiddleRight'
$lblInstV = New-Lbl $mainCard 'בודק...' 20 50 450 24 'MiddleLeft' 'strong' 10 'Segoe UI Semibold'
$lblLatestT = New-Lbl $mainCard 'הגרסה שנבחרה:' 475 76 160 24 'MiddleRight'
$lblLatestV = New-Lbl $mainCard 'טוען...' 128 76 342 24 'MiddleLeft' 'strong' 10 'Segoe UI Semibold'
$btnChangelog = New-Btn $mainCard 'מה חדש?' 20 75 100 27 9

$lblChannel = New-Lbl $mainCard 'ערוץ:' 560 108 75 26 'MiddleRight'
$cmbChannel = New-Object Windows.Forms.ComboBox
$cmbChannel.DropDownStyle = 'DropDownList'; $cmbChannel.Location = New-Object Drawing.Point(345, 108); $cmbChannel.Size = New-Object Drawing.Size(210, 26)
$cmbChannel.Font = New-Object Drawing.Font('Segoe UI', 9.5)
$mainCard.Controls.Add($cmbChannel)
$lblArch = New-Lbl $mainCard 'מעבד:' 255 108 80 26 'MiddleRight'
$cmbArch = New-Object Windows.Forms.ComboBox
$cmbArch.DropDownStyle = 'DropDownList'; $cmbArch.Location = New-Object Drawing.Point(150, 108); $cmbArch.Size = New-Object Drawing.Size(100, 26)
$cmbArch.Font = New-Object Drawing.Font('Segoe UI', 9.5)
[void]$cmbArch.Items.AddRange(@('x64', 'ARM64'))
$cmbArch.SelectedIndex = $(if ((Get-WinArch) -eq 'arm64') { 1 } else { 0 })
$mainCard.Controls.Add($cmbArch)

function New-Radio([string]$text, [int]$y) {
    $r = New-Object Windows.Forms.RadioButton
    $r.Text = $text; $r.Location = New-Object Drawing.Point(345, $y); $r.Size = New-Object Drawing.Size(290, 26)
    $r.Font = New-Object Drawing.Font('Segoe UI', 10)
    $mainCard.Controls.Add($r); return $r
}
$radioInstaller = New-Radio 'מתקין (ללא ספרייה)' 144
$radioPortable  = New-Radio 'גרסה ניידת (ZIP)' 172
$radioFull      = New-Radio 'מתקין מלא כולל הספרייה' 200
$radioLocal     = New-Radio 'קובץ שנמצא בתיקייה' 228
$radioInstaller.Checked = $true
$lblSizeInstaller = New-Lbl $mainCard '' 20 144 320 26 'MiddleLeft'
$lblSizePortable  = New-Lbl $mainCard '' 20 172 320 26 'MiddleLeft'
$lblSizeFull      = New-Lbl $mainCard '' 20 200 320 26 'MiddleLeft'
$lblLocalFile     = New-Lbl $mainCard '' 20 228 320 26 'MiddleLeft'
$toolTip.SetToolTip($radioFull, 'קובץ התקנה אחד (כ-2.5 ג"ב) שמכיל את התוכנה ואת הספרייה המלאה - מתאים למחשב ללא אינטרנט')
$toolTip.SetToolTip($radioPortable, 'גרסה שאינה דורשת התקנה - מחלצים לתיקייה (למשל בכונן נייד) ומפעילים')

$chkSilent = New-Object Windows.Forms.CheckBox
$chkSilent.Text = 'התקנה שקטה (בלי אשף)'; $chkSilent.Location = New-Object Drawing.Point(215, 258); $chkSilent.Size = New-Object Drawing.Size(225, 24)
$chkSilent.Font = New-Object Drawing.Font('Segoe UI', 9.5)
$mainCard.Controls.Add($chkSilent)

$btnSelectFile = New-Btn $mainCard 'בחר קובץ...' 495 288 140 36
$toolTip.SetToolTip($btnSelectFile, 'בחירת קובץ התקנה של אוצריא ממקום אחר')
$btnDL = New-Btn $mainCard 'הורדה' 215 288 225 36 10.5
$btnDL.Font = New-Object Drawing.Font('Segoe UI Semibold', 10.5)
$btnCancel = New-Btn $mainCard 'ביטול' 20 288 100 36
$btnCancel.Visible = $false

$progBar = New-Object ModernProgressBar
$progBar.Location = New-Object Drawing.Point(20, 332); $progBar.Size = New-Object Drawing.Size(615, 20)
$mainCard.Controls.Add($progBar)
$statusLbl = New-Lbl $mainCard '' 20 354 615 24 'MiddleCenter'

# --- Library card ---
$libCard = New-Card 482 262
$lblLib = New-Lbl $libCard 'ספריית אוצריא' 20 10 615 28 'MiddleCenter' 'strong' 13 'Segoe UI Semibold'
$div2 = New-Lbl $libCard '' 20 42 615 1 'MiddleCenter' 'divider'
$lblLibLocalT = New-Lbl $libCard 'במחשב:' 475 50 160 24 'MiddleRight'
$lblLibLocalV = New-Lbl $libCard 'בודק...' 20 50 450 24 'MiddleLeft' 'strong' 10 'Segoe UI Semibold'
$lblLibLatestT = New-Lbl $libCard 'גרסה אחרונה:' 475 76 160 24 'MiddleRight'
$lblLibLatestV = New-Lbl $libCard 'טוען...' 128 76 342 24 'MiddleLeft' 'strong' 10 'Segoe UI Semibold'
$btnLibChangelog = New-Btn $libCard 'מה חדש?' 20 75 100 27 9
$lblLibNote = New-Lbl $libCard ("הספרייה מתעדכנת אוטומטית מתוך אוצריא.`nלמחשב ללא אינטרנט: הורד חבילת ספרייה - המתקין יחלץ אותה.") 20 102 615 42 'MiddleCenter' '' 9
$chkIndex = New-Object Windows.Forms.CheckBox
$chkIndex.Text = 'כולל אינדקס חיפוש מוכן (למחשבים איטיים)'; $chkIndex.Location = New-Object Drawing.Point(120, 146); $chkIndex.Size = New-Object Drawing.Size(415, 24)
$chkIndex.Font = New-Object Drawing.Font('Segoe UI', 9.5)
$libCard.Controls.Add($chkIndex)
$toolTip.SetToolTip($chkIndex, 'בדרך כלל אין צורך: אוצריא בונה את אינדקס החיפוש בעצמה אחרי ההתקנה')
$btnLibDL = New-Btn $libCard 'הורדת חבילת ספרייה' 160 174 335 36 10.5
$btnLibDL.Font = New-Object Drawing.Font('Segoe UI Semibold', 10.5)
$btnLibCancel = New-Btn $libCard 'ביטול' 20 174 100 36
$btnLibCancel.Visible = $false
$libProgBar = New-Object ModernProgressBar
$libProgBar.Location = New-Object Drawing.Point(20, 216); $libProgBar.Size = New-Object Drawing.Size(615, 20)
$libCard.Controls.Add($libProgBar)
$libStatusLbl = New-Lbl $libCard '' 20 237 615 22 'MiddleCenter'

# --- Bottom buttons (form level: x is measured from the right) ---
$btnPlugins = New-Btn $form 'תוספים לאוצריא' 25 756 210 40 10.5
$btnPlugins.Font = New-Object Drawing.Font('Segoe UI Semibold', 10.5)
$btnOpenLib = New-Btn $form 'תיקיית הספרייה' 248 756 209 40 10.5
$btnReset   = New-Btn $form 'איפוס הגדרות' 470 756 210 40 10.5
$toolTip.SetToolTip($btnPlugins, 'התקנה ועדכון של תוספים מחנות התוספים של אוצריא')
$toolTip.SetToolTip($btnReset, 'מחיקת הגדרות, היסטוריה ומטמון של אוצריא (הספרייה והגיבויים נשמרים)')

$statusDivider = New-Lbl $form '' 25 806 655 1 'MiddleCenter' 'divider'
$statusLocation = New-Lbl $form '' 25 811 655 20 'MiddleCenter' '' 8.5
$statusBar = New-Lbl $form '' 25 830 655 20 'MiddleCenter' '' 8.5
$lblNetwork = New-Lbl $form 'אין חיבור לרשת' 25 848 655 20 'MiddleCenter' 'network' 9.5 'Segoe UI Semibold'
$lblNetwork.ForeColor = [Drawing.Color]::FromArgb(200, 50, 50); $lblNetwork.Visible = $false

# =====================================================================
# State & UI logic
# =====================================================================
$script:Install = $null
$script:Library = $null
$script:Releases = @()
$script:CurrentRelease = $null
$script:LibLatest = $null
$script:Online = $false
$script:SelectedLocalFile = $null

function Get-SelectedArch { if ($cmbArch.SelectedIndex -eq 1) { 'arm64' } else { 'x64' } }

function Get-VersionedName([string]$AssetName, [string]$Version) { return ($AssetName -replace '^otzaria-[\d\.]+-', "otzaria-$Version-") }

function Get-AppTarget {
    if ($radioLocal.Checked) {
        if (-not $script:SelectedLocalFile) { return $null }
        $f = $script:SelectedLocalFile
        $kind = if ($f.Name -like '*.zip') { 'portable' } else { 'installer' }
        $ver = ''; if ($f.Name -match '^otzaria-([\d\.]+)-') { $ver = $matches[1] }
        return [pscustomobject]@{ Kind = $kind; Path = $f.FullName; Name = $f.Name; Size = $f.Length; Items = @(); Assemble = @(); Version = $ver }
    }
    $rel = $script:CurrentRelease
    if (-not $rel) { return $null }
    $arch = Get-SelectedArch
    $archSuffix = if ($arch -eq 'arm64') { '_arm64' } else { '' }
    if ($radioInstaller.Checked) {
        $c = $rel.Components["otzaria-windows-$arch"]; if (-not $c) { return $null }
        $name = Get-VersionedName $c.Name $rel.Version; $path = Join-Path $script:WorkDir $name
        return [pscustomobject]@{ Kind = 'installer'; Path = $path; Name = $name; Size = $c.Size; Version = $rel.Version
            Items = @(@{ Url = $c.Url; Path = $path; Size = $c.Size; Sha256 = $c.Sha256 }); Assemble = @() }
    }
    if ($radioPortable.Checked) {
        $c = $rel.Components["otzaria-windows-portable-$arch"]; if (-not $c) { return $null }
        $name = "otzaria-$($rel.Version)-windows$archSuffix-portable.zip"; $path = Join-Path $script:WorkDir $name
        return [pscustomobject]@{ Kind = 'portable'; Path = $path; Name = $name; Size = $c.Size; Version = $rel.Version
            Items = @(@{ Url = $c.Url; Path = $path; Size = $c.Size; Sha256 = $c.Sha256 }); Assemble = @() }
    }
    if ($radioFull.Checked) {
        $id = if ($arch -eq 'arm64') { 'otzaria-windows-full-arm64' } else { 'otzaria-windows-full' }
        $c = $rel.Components[$id]; if (-not $c) { return $null }
        $name = Get-VersionedName $c.Name $rel.Version; $path = Join-Path $script:WorkDir $name
        if ($c.Kind -eq 'split' -and @($c.Parts).Count -gt 0) {
            $items = @(); $paths = @()
            foreach ($p in $c.Parts) { $pp = Join-Path $script:WorkDir $p.Name; $items += @{ Url = $p.Url; Path = $pp; Size = $p.Size; Sha256 = $p.Sha256 }; $paths += $pp }
            return [pscustomobject]@{ Kind = 'installer'; Path = $path; Name = $name; Size = $c.Size; Version = $rel.Version
                Items = $items; Assemble = @(@{ Output = $path; Parts = $paths; Sha256 = $c.Sha256 }) }
        }
        return [pscustomobject]@{ Kind = 'installer'; Path = $path; Name = $name; Size = $c.Size; Version = $rel.Version
            Items = @(@{ Url = $c.Url; Path = $path; Size = $c.Size; Sha256 = $c.Sha256 }); Assemble = @() }
    }
    return $null
}

# Library package = library parts (+ optional index) + matching installer of the same release
function Get-LibTarget {
    $rel = $script:CurrentRelease
    $arch = Get-SelectedArch
    if ($rel -and $rel.Components['library-full']) {
        $inst = $rel.Components["otzaria-windows-$arch"]
        if (-not $inst) { return $null }
        $instName = Get-VersionedName $inst.Name $rel.Version; $instPath = Join-Path $script:WorkDir $instName
        $items = @(@{ Url = $inst.Url; Path = $instPath; Size = $inst.Size; Sha256 = $inst.Sha256 })
        $libSize = 0L; $parts = @()
        $ids = @('library-full'); if ($chkIndex.Checked -and $rel.Components['library-index']) { $ids += 'library-index' }
        foreach ($id in $ids) {
            $c = $rel.Components[$id]
            foreach ($p in @($c.Parts)) { $pp = Join-Path $script:WorkDir $p.Name; $items += @{ Url = $p.Url; Path = $pp; Size = $p.Size; Sha256 = $p.Sha256 }; $parts += $pp; $libSize += $p.Size }
            if (@($c.Parts).Count -eq 0) { $pp = Join-Path $script:WorkDir $c.Name; $items += @{ Url = $c.Url; Path = $pp; Size = $c.Size; Sha256 = $c.Sha256 }; $parts += $pp; $libSize += $c.Size }
            if ($c.ManifestUrl) { $items += @{ Url = $c.ManifestUrl; Path = (Join-Path $script:WorkDir $c.ManifestAsset); Size = 0; Sha256 = '' } }
        }
        $ready = (Test-Path -LiteralPath $instPath)
        foreach ($it in $items) { if ($it.Size -gt 0 -and -not ((Test-Path -LiteralPath $it.Path) -and (Get-Item -LiteralPath $it.Path).Length -eq $it.Size)) { $ready = $false } }
        return [pscustomobject]@{ Installer = $instPath; Items = $items; LibSize = $libSize; Ready = $ready; Version = $rel.Version; Offline = $false }
    }
    # Offline: library parts + an installer of the same base version already in the folder
    $groups = Get-LocalLibraryParts
    foreach ($ver in ($groups.Keys | Sort-Object { $_ } -Descending)) {
        $inst = Get-LocalInstallers | Where-Object { $_.Name -match "^otzaria-$([regex]::Escape($ver))(\.\d+)?-windows(_arm64)?\.exe$" } | Select-Object -First 1
        if ($inst) {
            $size = 0L; foreach ($f in $groups[$ver]) { $size += $f.Length }
            return [pscustomobject]@{ Installer = $inst.FullName; Items = @(); LibSize = $size; Ready = $true; Version = $ver; Offline = $true }
        }
    }
    return $null
}

function Update-AppUI {
    $active = $script:Active -and $script:Active.Kind -eq 'App'
    $paused = $script:Paused.ContainsKey('App')
    $busy = $active -or $paused
    foreach ($c in @($cmbChannel, $cmbArch, $radioInstaller, $radioPortable, $radioFull, $radioLocal, $btnSelectFile)) { $c.Enabled = -not $busy }

    $rel = $script:CurrentRelease
    $arch = Get-SelectedArch
    $fmt = { param($c) if ($c) { Format-Size $c.Size } else { 'לא זמין בגרסה זו' } }
    if ($rel) {
        $lblSizeInstaller.Text = & $fmt $rel.Components["otzaria-windows-$arch"]
        $lblSizePortable.Text  = & $fmt $rel.Components["otzaria-windows-portable-$arch"]
        $fid = if ($arch -eq 'arm64') { 'otzaria-windows-full-arm64' } else { 'otzaria-windows-full' }
        $lblSizeFull.Text = & $fmt $rel.Components[$fid]
        $lblLatestV.Text = "$($rel.Version)  ·  $(Format-Date $rel.Date)" + $(if ($rel.Prerelease) { '  (טרום-שחרור)' } else { '' })
        $lblLatestV.Tag = 'strong'
        if ($script:Install -and $script:Install.Build -and $rel.Build) {
            if ($script:Install.Build -ge $rel.Build) { $lblLatestV.Text += '   ✓ מותקנת'; $lblLatestV.Tag = 'good' }
            else { $lblLatestV.Text += '   ← עדכון זמין'; $lblLatestV.Tag = 'warn' }
        }
    } else {
        $lblSizeInstaller.Text = ''; $lblSizePortable.Text = ''; $lblSizeFull.Text = ''
        $lblLatestV.Text = $(if ($script:Online) { 'לא נמצאה גרסה' } else { 'אין חיבור לרשת' }); $lblLatestV.Tag = 'warn'
    }
    foreach ($r in @($radioInstaller, $radioPortable, $radioFull)) { if (-not $busy) { $r.Enabled = [bool]$rel } }
    $radioLocal.Visible = [bool]$script:SelectedLocalFile
    $lblLocalFile.Visible = [bool]$script:SelectedLocalFile
    if ($script:SelectedLocalFile) { $lblLocalFile.Text = "$(Format-Size $script:SelectedLocalFile.Length)  ·  $($script:SelectedLocalFile.Name)" }
    if (-not $rel -and $script:SelectedLocalFile -and -not $radioLocal.Checked) { $radioLocal.Checked = $true }

    $t = Get-AppTarget
    $exists = $t -and (Test-Path -LiteralPath $t.Path)
    $btnCancel.Visible = $busy
    if ($active) { $btnDL.Text = 'עצור'; $btnDL.Enabled = $true }
    elseif ($paused) { $btnDL.Text = 'המשך'; $btnDL.Enabled = -not $script:Active }
    elseif ($exists) { $btnDL.Text = $(if ($t.Kind -eq 'portable') { 'חילוץ לתיקייה' } else { 'התקנה' }); $btnDL.Enabled = $true }
    elseif ($t) { $btnDL.Text = "הורדה ($(Format-Size $t.Size))"; $btnDL.Enabled = -not $script:Active }
    else { $btnDL.Text = 'הורדה'; $btnDL.Enabled = $false }
    $chkSilent.Visible = $exists -and -not $busy -and $t.Kind -eq 'installer'
    if (-not $busy -and -not $script:Active) {
        if ($exists) { $statusLbl.Text = "הקובץ קיים בתיקייה: $($t.Name)"; $progBar.Value = 100 }
        elseif ($statusLbl.Text -eq '' -or $progBar.Value -eq 100) { $statusLbl.Text = ''; $progBar.Value = 0 }
    }
    Apply-ThemeTo $mainCard
}

function Update-LibUI {
    $active = $script:Active -and $script:Active.Kind -eq 'Lib'
    $paused = $script:Paused.ContainsKey('Lib')
    $busy = $active -or $paused
    $lib = $script:Library
    if ($lib) {
        $v = if ($lib.Version) { "גרסה $($lib.Version)" } else { 'גרסה לא ידועה' }
        if ($lib.Schema) { $v += " (סכמה $($lib.Schema))" }
        $lblLibLocalV.Text = $v
        $toolTip.SetToolTip($lblLibLocalV, $lib.Db)
    } else { $lblLibLocalV.Text = 'לא נמצאה ספרייה' }
    $L = $script:LibLatest
    if ($L -and $L.Version) {
        $lblLibLatestV.Text = "גרסה $($L.Version)  ·  $(Format-Date $L.Date)"; $lblLibLatestV.Tag = 'strong'
        if ($lib -and $lib.Version) {
            if ([int]$lib.Version -ge [int]$L.Version) { $lblLibLatestV.Text += '   ✓ מעודכנת'; $lblLibLatestV.Tag = 'good' }
            else { $lblLibLatestV.Text += '   ← אוצריא תעדכן אוטומטית'; $lblLibLatestV.Tag = 'warn' }
        }
    } else { $lblLibLatestV.Text = $(if ($script:Online) { 'לא ידוע' } else { 'אין חיבור לרשת' }); $lblLibLatestV.Tag = '' }
    $btnLibChangelog.Enabled = [bool]($L -and $L.ChangesUrl)

    $t = Get-LibTarget
    $chkIndex.Enabled = -not $busy
    $btnLibCancel.Visible = $busy
    if ($active) { $btnLibDL.Text = 'עצור'; $btnLibDL.Enabled = $true }
    elseif ($paused) { $btnLibDL.Text = 'המשך'; $btnLibDL.Enabled = -not $script:Active }
    elseif ($t -and $t.Ready) { $btnLibDL.Text = 'התקנה עם הספרייה'; $btnLibDL.Enabled = $true }
    elseif ($t) { $btnLibDL.Text = "הורדת חבילת ספרייה ($(Format-Size $t.LibSize))"; $btnLibDL.Enabled = -not $script:Active }
    else { $btnLibDL.Text = 'הורדת חבילת ספרייה'; $btnLibDL.Enabled = $false }
    if (-not $busy -and -not $script:Active -and $t -and $t.Ready) {
        $libStatusLbl.Text = "חבילת הספרייה ($($t.Version)) מוכנה בתיקייה, ליד קובץ ההתקנה"; $libProgBar.Value = 100
    }
    Apply-ThemeTo $libCard
}

function Update-StatusBar {
    $loc = if ($script:Install) { "מותקן ב: $($script:Install.Dir)" } else { 'אוצריא לא מותקנת במחשב' }
    if ($script:Library) { $loc += "   |   ספרייה: $($script:Library.Dir)" }
    $statusLocation.Text = $loc
    $statusBar.Text = "תיקיית הורדות: $($script:WorkDir)"
    $toolTip.SetToolTip($statusLocation, $loc)
    $toolTip.SetToolTip($statusBar, $statusBar.Text)
}

function Update-AllUI { Update-AppUI; Update-LibUI }

function Refresh-LocalInfo {
    $script:Install = Get-OtzariaInstall
    $script:Library = Get-LocalLibrary
    if ($script:Install) {
        $lblInstV.Text = "$($script:Install.Version)  ·  $($script:Install.ModeText)"
    } else { $lblInstV.Text = 'לא מותקנת' }
    $loc = @(Get-LocalInstallers)
    if (-not $script:SelectedLocalFile -or -not (Test-Path -LiteralPath $script:SelectedLocalFile.FullName)) {
        $script:SelectedLocalFile = $(if ($loc.Count -gt 0) { $loc[0] } else { $null })
    }
    Update-StatusBar
    Update-AllUI
}

function Refresh-Data {
    $form.Cursor = 'WaitCursor'
    $statusLbl.Text = 'טוען נתונים...'; [Windows.Forms.Application]::DoEvents()
    Refresh-LocalInfo
    $netError = $null
    try { $script:Releases = @(Get-AppReleases); $script:Online = $true }
    catch { $script:Releases = @(); $script:Online = $false; $netError = $_.Exception.Message }
    try { if ($script:Online) { $script:LibLatest = Get-LatestLibraryRelease } } catch { $script:LibLatest = $null }
    $lblNetwork.Visible = -not $script:Online
    if (-not $script:Online -and $netError) { $lblNetwork.Text = "אין חיבור לשרת: $netError" }
    $cmbChannel.Items.Clear()
    foreach ($r in $script:Releases) {
        $label = if ($r.Prerelease) { "טרום-שחרור  $($r.Version)" } else { "יציבה  $($r.Version)" }
        [void]$cmbChannel.Items.Add($label)
    }
    if ($cmbChannel.Items.Count -gt 0) { $cmbChannel.SelectedIndex = 0; $script:CurrentRelease = $script:Releases[0] }
    else { $script:CurrentRelease = $null; [void]$cmbChannel.Items.Add('אין נתונים'); $cmbChannel.SelectedIndex = 0 }
    $statusLbl.Text = ''
    Update-AllUI
    $form.Cursor = 'Default'
}

# =====================================================================
# Event handlers - application card
# =====================================================================
$cmbChannel.Add_SelectedIndexChanged({
    $i = $cmbChannel.SelectedIndex
    if ($i -ge 0 -and $i -lt $script:Releases.Count) { $script:CurrentRelease = $script:Releases[$i] }
    Update-AllUI
})
$cmbArch.Add_SelectedIndexChanged({ Update-AllUI })
foreach ($r in @($radioInstaller, $radioPortable, $radioFull, $radioLocal)) { $r.Add_CheckedChanged({ if ($this.Checked) { $statusLbl.Text = ''; Update-AppUI } }) }
$chkIndex.Add_CheckedChanged({ Update-LibUI })

# "יומן שינויים" של אוצריא - הקובץ assets/יומן שינויים.md שבתוך התוכנה.
# מבנה: "* **0.9.98.3**" ואחריו שורות "  - חידוש". מחזיר רשימת @{ Version; Items }.
$script:ChangelogFile = 'assets/' + [string][char]0x05D9 + [char]0x05D5 + [char]0x05DE + [char]0x05DF + ' ' + [char]0x05E9 + [char]0x05D9 + [char]0x05E0 + [char]0x05D5 + [char]0x05D9 + [char]0x05D9 + [char]0x05DD + '.md'
function ConvertFrom-AppChangelog([string]$md) {
    $sections = @(); $cur = $null
    foreach ($line in ($md -split "`r?`n")) {
        if ($line -match '^\*\s+\*\*\s*v?([\d\.]+)\s*\*\*') {
            $cur = [pscustomobject]@{ Version = $matches[1]; Items = New-Object System.Collections.Generic.List[string] }
            $sections += $cur
        } elseif ($cur -and $line -match '^\s+[-\*]\s+(.+)$') {
            $item = $matches[1] -replace '`([^`]+)`', '$1' -replace '\[([^\]]+)\]\([^\)]+\)', '$1'
            $cur.Items.Add($item.Trim())
        }
    }
    return ,$sections
}
function Get-AppChangelogText([string]$Tag) {
    # 1) מהאינטרנט - הקובץ של הגרסה המדויקת שנבחרה; 2) העותק שבתיקיית ההתקנה
    $md = $null
    $script:ChangelogLocal = $false
    $path = ($script:ChangelogFile -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
    foreach ($ref in @($Tag, 'dev')) {
        if (-not $ref) { continue }
        try { $md = Get-WebText "https://raw.githubusercontent.com/$($script:AppRepo)/$([uri]::EscapeDataString($ref))/$path"; break } catch { }
    }
    if (-not $md -and $script:Install) {
        $dir = Join-Path $script:Install.Dir 'data\flutter_assets\assets'
        $f = Get-ChildItem -LiteralPath $dir -Filter '*.md' -ErrorAction SilentlyContinue |
             Where-Object { [uri]::UnescapeDataString($_.Name) -eq (Split-Path -Leaf $script:ChangelogFile) } | Select-Object -First 1
        if ($f) { $md = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8); $script:ChangelogLocal = $true }
    }
    return $md
}

$btnChangelog.Add_Click({
    $rel = $script:CurrentRelease
    $target = if ($rel) { $rel.Version } elseif ($script:Install) { $script:Install.Version } else { $null }
    $form.Cursor = 'WaitCursor'
    $md = Get-AppChangelogText $(if ($rel) { $rel.Tag } else { $null })
    $form.Cursor = 'Default'
    if (-not $md) { Show-RTLMessageBox 'לא ניתן לטעון את יומן השינויים (אין חיבור לרשת ואוצריא אינה מותקנת).' 'יומן שינויים' 'OK' 'Warning' | Out-Null; return }
    $all = ConvertFrom-AppChangelog $md
    if ($all.Count -eq 0) { Show-TextDialog 'יומן שינויים' (Convert-MarkdownToText $md); return }
    $inst = if ($script:Install) { $script:Install.Version } else { $null }
    # מה חדש מאז הגרסה המותקנת ועד הגרסה שנבחרה; אם אין הבדל - החידושים של הגרסה שנבחרה
    $show = @()
    if ($inst -and $target -and (Compare-Ver $target $inst) -gt 0) {
        $show = @($all | Where-Object { (Compare-Ver $_.Version $inst) -gt 0 -and (Compare-Ver $_.Version $target) -le 0 })
        $title = "מה חדש מאז הגרסה שלך ($inst)"
    }
    if ($show.Count -eq 0) {
        $show = @($all | Where-Object { $target -and (Compare-Ver $_.Version $target) -eq 0 })
        if ($show.Count -eq 0) { $show = @($all[0]) }
        $title = "מה חדש בגרסה $($show[0].Version)"
    }
    $sb = New-Object System.Text.StringBuilder
    foreach ($sec in $show) {
        [void]$sb.AppendLine("## גרסה $($sec.Version)")
        foreach ($i in $sec.Items) { [void]$sb.AppendLine("•  $i") }
        [void]$sb.AppendLine('')
    }
    $older = @($all | Where-Object { -not ($show -contains $_) } | Select-Object -First 3)
    if ($older.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('## גרסאות קודמות')
        [void]$sb.AppendLine('')
        foreach ($sec in $older) {
            [void]$sb.AppendLine("**גרסה $($sec.Version)**")
            foreach ($i in $sec.Items) { [void]$sb.AppendLine("•  $i") }
            [void]$sb.AppendLine('')
        }
    }
    if ($script:ChangelogLocal) { $title += '  (אין חיבור לרשת - מתוך אוצריא המותקנת)' }
    Show-TextDialog $title $sb.ToString().TrimEnd()
})

$btnSelectFile.Add_Click({
    $ofd = New-Object Windows.Forms.OpenFileDialog
    $ofd.Filter = 'קבצי אוצריא (*.exe;*.zip)|*.exe;*.zip|כל הקבצים (*.*)|*.*'
    $ofd.InitialDirectory = $script:WorkDir
    if ($ofd.ShowDialog() -eq 'OK') {
        $script:SelectedLocalFile = Get-Item -LiteralPath $ofd.FileName
        Update-AppUI
        $radioLocal.Checked = $true
        Update-AppUI
    }
})

$btnDL.Add_Click({
    if ($script:Active -and $script:Active.Kind -eq 'App') { Stop-Download 'App'; return }
    if ($script:Paused.ContainsKey('App')) {
        $p = $script:Paused['App']
        Start-Download -Kind 'App' -Items $p.Items -Assemble $p.Assemble -Bar $progBar -Label $statusLbl -OnDone $p.OnDone -Context $p.Context
        return
    }
    $t = Get-AppTarget
    if (-not $t) { return }
    if (Test-Path -LiteralPath $t.Path) {
        if ($t.Kind -eq 'portable') { Expand-PortableZip $t.Path } else { Invoke-Installer $t.Path $chkSilent.Checked }
        return
    }
    $need = 0L; foreach ($i in $t.Items) { $need += [long]$i.Size }
    if (@($t.Assemble).Count -gt 0) { $need *= 2 }
    try {
        $free = (New-Object IO.DriveInfo([IO.Path]::GetPathRoot($script:WorkDir))).AvailableFreeSpace
        if ($free -lt $need) {
            if ((Show-RTLMessageBox "אין מספיק מקום פנוי בכונן.`nנדרש: $(Format-Size $need)`nפנוי: $(Format-Size $free)`n`nלהמשיך בכל זאת?" 'מקום בדיסק' 'YesNo' 'Warning') -ne 'Yes') { return }
        }
    } catch { }
    Start-Download -Kind 'App' -Items $t.Items -Assemble $t.Assemble -Bar $progBar -Label $statusLbl -Context $t -OnDone {
        param($ctx)
        $statusLbl.Text = "ההורדה הושלמה: $($ctx.Name)"
        Remove-OldInstallers $ctx.Name $ctx.Version
    }
})
$btnCancel.Add_Click({ Cancel-Download 'App' })

# =====================================================================
# Event handlers - library card
# =====================================================================
$btnLibChangelog.Add_Click({
    $L = $script:LibLatest
    if (-not $L -or -not $L.ChangesUrl) { return }
    $form.Cursor = 'WaitCursor'
    try { $md = Get-WebText $L.ChangesUrl } catch { $md = "שגיאה בטעינת רשימת השינויים: $($_.Exception.Message)" }
    $form.Cursor = 'Default'
    Show-TextDialog "עדכוני ספרייה - גרסה $($L.Version)" (Convert-MarkdownToText $md)
})

$btnLibDL.Add_Click({
    if ($script:Active -and $script:Active.Kind -eq 'Lib') { Stop-Download 'Lib'; return }
    if ($script:Paused.ContainsKey('Lib')) {
        $p = $script:Paused['Lib']
        Start-Download -Kind 'Lib' -Items $p.Items -Assemble $p.Assemble -Bar $libProgBar -Label $libStatusLbl -OnDone $p.OnDone -Context $p.Context
        return
    }
    $t = Get-LibTarget
    if (-not $t) { return }
    if ($t.Ready) {
        $msg = "המתקין של אוצריא $($t.Version) יופעל ויחלץ את הספרייה מהקבצים שבתיקייה.`n`n" +
               "שים לב: הספרייה הקיימת במחשב תוחלף בספרייה שבחבילה. אם הספרייה במחשב חדשה יותר, אוצריא תעדכן אותה שוב כשתהיה מחוברת לרשת.`n`nלהמשיך?"
        if ((Show-RTLMessageBox $msg 'התקנה עם הספרייה' 'YesNo' 'Question') -eq 'Yes') { Invoke-Installer $t.Installer $chkSilent.Checked }
        return
    }
    $need = 0L; foreach ($i in $t.Items) { $need += [long]$i.Size }
    $msg = "יורדו לתיקייה:`n• קובץ ההתקנה של אוצריא $($t.Version)`n• חבילת הספרייה" + $(if ($chkIndex.Checked) { ' ואינדקס החיפוש' } else { '' }) +
           "`n`nסך הכל: $(Format-Size $need)`nתיקייה: $($script:WorkDir)`n`nאת התיקייה כולה אפשר להעביר למחשב ללא אינטרנט ולהפעיל שם את קובץ ההתקנה.`n`nלהתחיל?"
    if ((Show-RTLMessageBox $msg 'הורדת חבילת ספרייה' 'YesNo' 'Question') -ne 'Yes') { return }
    try {
        $free = (New-Object IO.DriveInfo([IO.Path]::GetPathRoot($script:WorkDir))).AvailableFreeSpace
        if ($free -lt $need -and (Show-RTLMessageBox "אין מספיק מקום פנוי (נדרש $(Format-Size $need), פנוי $(Format-Size $free)). להמשיך?" 'מקום בדיסק' 'YesNo' 'Warning') -ne 'Yes') { return }
    } catch { }
    Start-Download -Kind 'Lib' -Items $t.Items -Bar $libProgBar -Label $libStatusLbl -Context $t.Version -OnDone {
        param($ver)
        $libStatusLbl.Text = 'חבילת הספרייה מוכנה. לחץ "התקנה עם הספרייה".'
        $other = @(Get-ChildItem -LiteralPath $script:WorkDir -File -ErrorAction SilentlyContinue | Where-Object {
            $_.Name -match '^otzaria-([\d\.]+)-library' -and $matches[1] -ne (($ver -split '\.')[0..2] -join '.') })
        if ($other.Count -gt 0 -and (Show-RTLMessageBox "נמצאו בתיקייה $($other.Count) קבצי ספרייה של גרסאות אחרות. למחוק אותם?" 'ניקוי' 'YesNo' 'Question') -eq 'Yes') {
            $other | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
        }
    }
})
$btnLibCancel.Add_Click({ Cancel-Download 'Lib' })

# =====================================================================
# Bottom buttons
# =====================================================================
$btnOpenLib.Add_Click({
    $d = if ($script:Library) { $script:Library.Dir } else { Get-DataRoot }
    if (Test-Path -LiteralPath $d) { Start-Process explorer.exe -ArgumentList "`"$d`"" } else { Show-RTLMessageBox 'תיקיית הספרייה לא נמצאה.' | Out-Null }
})

$btnReset.Add_Click({
    $root = Get-DataRoot
    $msg = "פעולה זו מוחקת את ההגדרות, ההיסטוריה, הסימניות, ההערות האישיות, התוספים ונתוני התוספים של אוצריא.`n" +
           "הספרייה (books) ותיקיית הגיבויים (backups) לא יימחקו.`n`nתיקייה: $root`n`nלא ניתן לבטל פעולה זו. להמשיך?"
    if ((Show-RTLMessageBox $msg 'איפוס הגדרות אוצריא' 'YesNo' 'Warning') -ne 'Yes') { return }
    if (-not (Confirm-CloseOtzaria 'יש לסגור אותה לפני האיפוס.')) { return }
    $keepNames = @('books', 'backups', 'library_path.txt', 'library_database_path.txt')
    $protected = if ($script:Library) { [IO.Path]::GetFullPath($script:Library.Dir).TrimEnd('\') } else { '' }
    $errors = 0
    foreach ($dir in @($root, (Join-Path $env:LOCALAPPDATA 'otzaria'), (Join-Path $env:APPDATA 'com.example\otzaria'))) {
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        foreach ($e in (Get-ChildItem -LiteralPath $dir -Force -ErrorAction SilentlyContinue)) {
            if ($dir -eq $root -and $keepNames -contains $e.Name.ToLower()) { continue }
            $full = $e.FullName.TrimEnd('\')
            if ($protected -and ($protected -eq $full -or $protected.StartsWith($full + '\', [StringComparison]::OrdinalIgnoreCase))) { continue }
            try { Remove-Item -LiteralPath $e.FullName -Recurse -Force -ErrorAction Stop } catch { $errors++ }
        }
    }
    if ($errors -eq 0) { Show-RTLMessageBox 'ההגדרות אופסו בהצלחה.' 'איפוס' | Out-Null }
    else { Show-RTLMessageBox "האיפוס הושלם, אך $errors פריטים לא נמחקו (ייתכן שהם בשימוש)." 'איפוס' 'OK' 'Warning' | Out-Null }
    Refresh-LocalInfo
})

$btnTheme.Add_Click({
    $script:DarkMode = -not $script:DarkMode; $script:DarkModeManual = $true
    Apply-FormTheme $form; Update-AllUI
})
$btnRefresh.Add_Click({ if (-not $script:Active) { Refresh-Data } })

# =====================================================================
# Plugins window
# =====================================================================
$script:PluginBaseline = @('plugin.storage.read', 'plugin.storage.write', 'app.info.read', 'ui.feedback', 'notifications.send', 'events.subscribe:theme.changed')
# Off by default in Otzaria's own permission dialog (pluginPermissionDefaultGrant)
$script:PluginDefaultOff = @('app.run_on_startup', 'app.startup_contributions', 'app.background_keep_alive', 'clipboard.read')

# Background worker: download + verify + extract + register (plugins_host.db) each plugin.
# Output: ITEM:<i>:<text> | OK:<i>:<installed|updated>:<version> | FAIL:<i>:<msg> | DONE
$script:PluginWorker = {
    param([string]$PlanJson, [string]$SqliteCs)
    $ErrorActionPreference = 'Stop'
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    Add-Type -TypeDefinition $SqliteCs
    function CmpVer([string]$a, [string]$b) {
        $pa = @(([string]$a -split '[-+ ]')[0] -split '\.'); $pb = @(([string]$b -split '[-+ ]')[0] -split '\.')
        for ($i = 0; $i -lt [Math]::Max($pa.Count, $pb.Count); $i++) {
            $x = 0L; $y = 0L
            if ($i -lt $pa.Count) { [void][long]::TryParse(($pa[$i] -replace '\D.*$',''), [ref]$x) }
            if ($i -lt $pb.Count) { [void][long]::TryParse(($pb[$i] -replace '\D.*$',''), [ref]$y) }
            if ($x -lt $y) { return -1 }; if ($x -gt $y) { return 1 }
        }
        return 0
    }
    $plan = ConvertFrom-Json -InputObject $PlanJson
    $dataRoot = [string]$plan.DataRoot
    $dbPath = Join-Path $dataRoot 'databases\plugins_host.db'
    $installedRoot = Join-Path $dataRoot 'plugins\installed'
    $tmpDir = Join-Path $env:TEMP ('otzaria_plugins_' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
    $baseline = @($plan.Baseline); $defaultOff = @($plan.DefaultOff)
    $allPerms = [bool]$plan.AllPermissions
    $i = -1
    try {
        foreach ($it in @($plan.Items)) {
            $i++
            $release = $null
            try {
                $pkg = Join-Path $tmpDir "$i.otzplugin"
                if ($it.LocalPath) {
                    Write-Output "ITEM:$($i):מעתיק מהתיקייה"
                    Copy-Item -LiteralPath ([string]$it.LocalPath) -Destination $pkg -Force
                } else {
                    Write-Output "ITEM:$($i):מוריד"
                    # עד 3 ניסיונות - תקלת רשת רגעית לא תכשיל את ההתקנה
                    for ($try = 1; $try -le 3; $try++) {
                        $wc = New-Object Net.WebClient; $wc.Headers.Add('User-Agent', 'OtzariaUpdater/2.0')
                        try { $wc.DownloadFile([string]$it.Url, $pkg); $dlErr = $null }
                        catch { $dlErr = $_.Exception; while ($dlErr.InnerException) { $dlErr = $dlErr.InnerException } }
                        finally { $wc.Dispose() }
                        $okSize = (Test-Path -LiteralPath $pkg) -and ((-not $it.Size) -or (Get-Item -LiteralPath $pkg).Length -eq [long]$it.Size)
                        if (-not $dlErr -and $okSize) { break }
                        if ($try -eq 3) {
                            if (-not $dlErr) { throw 'ההורדה לא הושלמה (הקובץ שהתקבל חלקי)' }
                            $code = 0; if ($dlErr -is [Net.WebException] -and $dlErr.Response) { try { $code = [int]$dlErr.Response.StatusCode } catch { } }
                            if ($code -eq 404) { throw 'הקובץ לא נמצא בשרת (ייתכן שהתוסף הוסר מהחנות)' }
                            elseif ($code -ge 500) { throw "שרת התוספים לא זמין כרגע (שגיאה $code). נסה שוב מאוחר יותר." }
                            elseif ($code -gt 0) { throw "השרת דחה את ההורדה (שגיאה $code)" }
                            else { throw 'אין חיבור לשרת התוספים. בדוק את החיבור לאינטרנט ונסה שוב.' }
                        }
                        Write-Output "ITEM:$($i):מנסה שוב ($($try + 1)/3)"
                        Start-Sleep -Seconds (2 * $try)
                    }
                }
                if ($it.Sha256) {
                    $h = (Get-FileHash -LiteralPath $pkg -Algorithm SHA256).Hash.ToLowerInvariant()
                    if ($h -ne ([string]$it.Sha256).ToLowerInvariant()) { throw 'בדיקת התקינות (SHA256) נכשלה' }
                }
                Write-Output "ITEM:$($i):מתקין"
                $zip = [IO.Compression.ZipFile]::OpenRead($pkg)
                try {
                    $manEntry = $zip.Entries | Where-Object { $_.FullName -eq 'manifest.json' } | Select-Object -First 1
                    if (-not $manEntry) { throw 'manifest.json לא נמצא בחבילה' }
                    $sr = New-Object IO.StreamReader($manEntry.Open(), [Text.Encoding]::UTF8)
                    try { $manText = $sr.ReadToEnd().TrimStart([char]0xFEFF) } finally { $sr.Close() }
                    $m = ConvertFrom-Json -InputObject $manText
                    $id = [string]$m.id
                    if (-not $id -or $id -match '[\\/:*?"<>|]' -or $id.Contains('..')) { throw "מזהה תוסף לא תקין: $id" }
                    if ($it.Uid -and $id -ne [string]$it.Uid) { throw "מזהה התוסף בחבילה ($id) שונה מהקטלוג" }
                    if (-not $m.version -or -not $m.entrypoint) { throw 'המניפסט חסר שדות חובה' }
                    $entry = ([string]$m.entrypoint) -replace '\\', '/' -replace '^\./', ''
                    if (-not ($zip.Entries | Where-Object { $_.FullName -eq $entry })) { throw "קובץ הכניסה $entry חסר בחבילה" }
                    if ($plan.AppVersion) {
                        if ($m.minAppVersion -and (CmpVer $plan.AppVersion $m.minAppVersion) -lt 0) { throw "דורש אוצריא $($m.minAppVersion) ומעלה" }
                        if ($m.maxAppVersion -and (CmpVer $plan.AppVersion $m.maxAppVersion) -gt 0) { throw "תומך עד אוצריא $($m.maxAppVersion) בלבד" }
                    }
                    $parent = Join-Path $installedRoot $id
                    New-Item -ItemType Directory -Force -Path $parent | Out-Null
                    $suffix = -join ((1..8) | ForEach-Object { '{0:x}' -f (Get-Random -Maximum 16) })
                    $release = Join-Path $parent ".release-$suffix"
                    New-Item -ItemType Directory -Force -Path $release | Out-Null
                    $relFull = [IO.Path]::GetFullPath($release).TrimEnd('\') + '\'
                    foreach ($e in $zip.Entries) {
                        $target = [IO.Path]::GetFullPath((Join-Path $release $e.FullName))
                        if (-not $target.StartsWith($relFull, [StringComparison]::OrdinalIgnoreCase)) { throw 'נתיב לא חוקי בחבילה' }
                        if ([string]::IsNullOrEmpty($e.Name)) { New-Item -ItemType Directory -Force -Path $target | Out-Null; continue }
                        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $target) | Out-Null
                        [IO.Compression.ZipFileExtensions]::ExtractToFile($e, $target, $true)
                    }
                } finally { $zip.Dispose() }

                # Register in plugins_host.db, mirroring Otzaria's own install transaction
                $now = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.ffffff')
                $db = [OtzSql]::Open($dbPath, $true)
                $oldPath = $null; $outcome = 'installed'
                try {
                    $ex = @([OtzSql]::Query($db, 'SELECT install_path, installed_at, enabled, pinned, pinned_to_nav_rail, hidden_from_tools, user_order, source_type FROM plugin_installation WHERE plugin_id = ?', [object[]]@($id)))
                    $prev = @{}
                    foreach ($g in [OtzSql]::Query($db, 'SELECT permission, granted FROM plugin_permission_grant WHERE plugin_id = ?', [object[]]@($id))) { $prev[$g[0]] = ($g[1] -eq '1') }
                    $pinnedDefault = 1
                    if ($m.toolTab -and $null -ne $m.toolTab.defaultPinned) { $pinnedDefault = $(if ($m.toolTab.defaultPinned) { 1 } else { 0 }) }
                    if ($ex.Count -gt 0) {
                        $r = $ex[0]
                        if ($r[7] -eq 'development') { throw 'תוסף בפיתוח מקומי - לא מעודכן' }
                        $outcome = 'updated'; $oldPath = $r[0]
                        $installedAt = $r[1]; $enabled = [int]$r[2]; $pinned = [int]$r[3]; $nav = [int]$r[4]; $hidden = [int]$r[5]
                        $userOrder = if ($r[6]) { [long]$r[6] } else { $null }
                    } else {
                        $installedAt = $now; $enabled = 1; $pinned = $pinnedDefault; $nav = 0; $hidden = 0
                        $mx = [OtzSql]::Query($db, 'SELECT MAX(user_order) FROM plugin_installation', $null)
                        $userOrder = if ($mx.Count -gt 0 -and $mx[0][0]) { [long]$mx[0][0] + 1 } else { $null }
                    }
                    $perms = New-Object System.Collections.Generic.List[string]
                    foreach ($p in @($m.permissions)) { $ps = [string]$p; if ($ps -and $baseline -notcontains $ps -and -not $perms.Contains($ps)) { $perms.Add($ps) } }
                    $icon = if ($m.icon) { [string]$m.icon } else { $null }
                    [OtzSql]::Exec($db, 'BEGIN IMMEDIATE', $null)
                    try {
                        [OtzSql]::Exec($db, 'INSERT OR REPLACE INTO plugin_installation (plugin_id, name, version, install_path, entrypoint_path, icon_path, enabled, pinned, pinned_to_nav_rail, hidden_from_tools, manifest_json, installed_at, updated_at, source_type, dev_root_path, user_order) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                            [object[]]@($id, [string]$m.name, [string]$m.version, $release, [string]$m.entrypoint, $icon, $enabled, $pinned, $nav, $hidden, $manText.Trim(), $installedAt, $now, 'packaged', $null, $userOrder))
                        [OtzSql]::Exec($db, 'DELETE FROM plugin_permission_grant WHERE plugin_id = ?', [object[]]@($id))
                        foreach ($p in $perms) {
                            if ($allPerms) { $g = 1 }
                            elseif ($prev.ContainsKey($p)) { $g = $(if ($prev[$p]) { 1 } else { 0 }) }
                            elseif ($defaultOff -contains $p) { $g = 0 }
                            else { $g = 1 }
                            [OtzSql]::Exec($db, 'INSERT OR REPLACE INTO plugin_permission_grant (plugin_id, permission, granted, granted_at) VALUES (?,?,?,?)', [object[]]@($id, $p, $g, $now))
                        }
                        [OtzSql]::Exec($db, 'COMMIT', $null)
                    } catch { try { [OtzSql]::Exec($db, 'ROLLBACK', $null) } catch { }; throw }
                } finally { [OtzSql]::Close($db) }
                $release = $null   # committed - keep it

                if ($oldPath) {
                    $parentFull = [IO.Path]::GetFullPath($parent).TrimEnd('\') + '\'
                    try { $oldFull = [IO.Path]::GetFullPath($oldPath) } catch { $oldFull = '' }
                    if ($oldFull -and $oldFull.StartsWith($parentFull, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $oldFull)) {
                        Remove-Item -LiteralPath $oldFull -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
                $cg = Join-Path $dataRoot 'plugin_crash_guard.json'
                if (Test-Path -LiteralPath $cg) {
                    try {
                        $j = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($cg))
                        $bl = @($j.blocked | Where-Object { $_ -and ([string]$_ -ne $id) -and ([string]$_.pluginId -ne $id) -and ([string]$_.id -ne $id) })
                        if ($bl.Count -ne @($j.blocked).Count) {
                            $j.blocked = $bl
                            [IO.File]::WriteAllText($cg, (ConvertTo-Json -InputObject $j -Depth 6 -Compress), (New-Object Text.UTF8Encoding($false)))
                        }
                    } catch { }
                }
                Write-Output "OK:$($i):$($outcome):$($m.version):$id"
            } catch {
                if ($release -and (Test-Path -LiteralPath $release)) { Remove-Item -LiteralPath $release -Recurse -Force -ErrorAction SilentlyContinue }
                Write-Output ("FAIL:$($i):" + $_.Exception.Message)
            }
        }
    } finally { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Output 'DONE'
}

function Get-InstalledPlugins {
    $h = @{}
    $db = Join-Path (Get-DataRoot) 'databases\plugins_host.db'
    if (-not (Test-Path -LiteralPath $db)) { return $h }
    $c = [IntPtr]::Zero
    try {
        $c = [OtzSql]::Open($db, $false)
        foreach ($r in [OtzSql]::Query($c, 'SELECT plugin_id, version, name, source_type FROM plugin_installation', $null)) {
            $h[$r[0]] = [pscustomobject]@{ Version = $r[1]; Name = $r[2]; Source = $r[3] }
        }
    } catch { } finally { [OtzSql]::Close($c) }
    return $h
}

# מחיקת תוספים - כמו uninstallPlugin של אוצריא: רישום, הרשאות, נתונים שמורים,
# ותיקיות התוסף (קבצים, data, cache). אוצריא חייבת להיות סגורה.
function Remove-OtzariaPlugins([string[]]$Ids) {
    $dataRoot = Get-DataRoot
    $pluginsRoot = [IO.Path]::GetFullPath((Join-Path $dataRoot 'plugins')).TrimEnd('\') + '\'
    $res = @{ Ok = 0; Fails = @() }
    $db = [OtzSql]::Open((Join-Path $dataRoot 'databases\plugins_host.db'), $true)
    try {
        foreach ($id in $Ids) {
            try {
                if (-not $id -or $id -match '[\\/:*?"<>|]' -or $id.Contains('..')) { throw 'מזהה לא תקין' }
                $rows = @([OtzSql]::Query($db, 'SELECT install_path, source_type FROM plugin_installation WHERE plugin_id = ?', [object[]]@($id)))
                [OtzSql]::Exec($db, 'BEGIN IMMEDIATE', $null)
                try {
                    foreach ($sql in 'DELETE FROM plugin_installation WHERE plugin_id = ?', 'DELETE FROM plugin_permission_grant WHERE plugin_id = ?',
                                     'DELETE FROM plugin_kv_store WHERE plugin_id = ?', 'DELETE FROM plugin_published_record WHERE plugin_id = ?',
                                     'DELETE FROM plugin_runtime_log WHERE plugin_id = ?') {
                        [OtzSql]::Exec($db, $sql, [object[]]@($id))
                    }
                    [OtzSql]::Exec($db, 'COMMIT', $null)
                } catch { try { [OtzSql]::Exec($db, 'ROLLBACK', $null) } catch { }; throw }
                $dirs = @((Join-Path $dataRoot "plugins\installed\$id"), (Join-Path $dataRoot "plugins\data\$id"), (Join-Path $dataRoot "plugins\cache\$id"))
                if ($rows.Count -gt 0 -and $rows[0][0] -and $rows[0][1] -ne 'development') { $dirs += $rows[0][0] }
                foreach ($d in $dirs) {
                    try { $full = [IO.Path]::GetFullPath($d) } catch { continue }
                    if ($full.StartsWith($pluginsRoot, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $full)) {
                        Remove-Item -LiteralPath $full -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
                $res.Ok++
            } catch { $res.Fails += "• ${id}: $($_.Exception.Message)" }
        }
    } finally { [OtzSql]::Close($db) }
    return $res
}

function Get-VersionSortKey([string]$v) {
    $parts = @((([string]$v -split '[-+ ]')[0]) -split '\.')
    return (($parts | ForEach-Object { $n = 0L; [void][long]::TryParse(($_ -replace '\D.*$', ''), [ref]$n); '{0:D8}' -f $n }) -join '.')
}

# ---------------------------------------------------------------------
# סדר התוספים בחלון "כלים ותוספים" של אוצריא.
# אוצריא ממיינת לפי: user_order (אם קיים, 1000+ערך) אחרת contributes.toolTab.order
# (ברירת מחדל 900), ואז לפי תאריך התקנה ומזהה. שמירה = user_order רציף לכל התוספים.
# ---------------------------------------------------------------------
function Get-PluginOrderList {
    $dbPath = Join-Path (Get-DataRoot) 'databases\plugins_host.db'
    $list = @()
    $db = [OtzSql]::Open($dbPath, $false)
    try {
        foreach ($r in [OtzSql]::Query($db, 'SELECT plugin_id, name, user_order, installed_at, manifest_json FROM plugin_installation', $null)) {
            $title = $r[1]; $mOrder = 900; $headless = $false
            try {
                $m = ConvertFrom-Json -InputObject $r[4]
                if ($m.contributes -and $m.contributes.toolTab) {
                    if ($m.contributes.toolTab.title) { $title = [string]$m.contributes.toolTab.title }
                    if ($null -ne $m.contributes.toolTab.order) { $mOrder = [int]$m.contributes.toolTab.order }
                }
                if ($m.headless -eq $true) { $headless = $true }
            } catch { }
            $eff = if ($r[2]) { 1000 + [int]$r[2] } else { $mOrder }
            $list += [pscustomobject]@{ Id = $r[0]; Title = $title; Eff = $eff; Installed = [string]$r[3]; Headless = $headless }
        }
    } finally { [OtzSql]::Close($db) }
    return @($list | Sort-Object Eff, Installed, Id)
}

function Save-PluginOrder([string[]]$Ids) {
    $db = [OtzSql]::Open((Join-Path (Get-DataRoot) 'databases\plugins_host.db'), $true)
    try {
        [OtzSql]::Exec($db, 'BEGIN IMMEDIATE', $null)
        try {
            for ($i = 0; $i -lt $Ids.Count; $i++) {
                [OtzSql]::Exec($db, 'UPDATE plugin_installation SET user_order = ? WHERE plugin_id = ?', [object[]]@($i, $Ids[$i]))
            }
            [OtzSql]::Exec($db, 'COMMIT', $null)
        } catch { try { [OtzSql]::Exec($db, 'ROLLBACK', $null) } catch { }; throw }
    } finally { [OtzSql]::Close($db) }
}

function Show-PluginOrderDialog {
    if (-not (Test-Path -LiteralPath (Join-Path (Get-DataRoot) 'databases\plugins_host.db'))) {
        Show-RTLMessageBox 'לא נמצא מסד התוספים של אוצריא.' 'סידור תוספים' 'OK' 'Warning' | Out-Null; return
    }
    $items = @(Get-PluginOrderList | Where-Object { -not $_.Headless })
    $od = New-Object Windows.Forms.Form
    $od.Text = 'סידור התוספים באוצריא'
    $od.ClientSize = New-Object Drawing.Size(560, 640); $od.MinimumSize = New-Object Drawing.Size(480, 480)
    $od.StartPosition = 'CenterParent'; $od.RightToLeft = 'Yes'; $od.RightToLeftLayout = $true
    $od.MinimizeBox = $false; $od.Font = New-Object Drawing.Font('Segoe UI', 10); $od.Icon = $form.Icon

    $hint = New-Object Windows.Forms.Label
    $hint.Text = "זה הסדר שבו התוספים מופיעים בחלון 'כלים ותוספים' באוצריא.`nגרור תוסף למקום אחר, או השתמש בכפתורים."
    $hint.Dock = 'Top'; $hint.Height = 50; $hint.TextAlign = 'MiddleCenter'; $hint.Padding = New-Object Windows.Forms.Padding(8, 4, 8, 0)

    $lb = New-Object Windows.Forms.ListBox
    $lb.Dock = 'Fill'; $lb.Font = New-Object Drawing.Font('Segoe UI', 11); $lb.IntegralHeight = $false
    $lb.AllowDrop = $true; $lb.BorderStyle = 'None'; $lb.SelectionMode = 'One'
    $lb.DisplayMember = 'Title'
    foreach ($it in $items) { [void]$lb.Items.Add($it) }
    $wrap = New-Object Windows.Forms.Panel; $wrap.Dock = 'Fill'; $wrap.Padding = New-Object Windows.Forms.Padding(12, 4, 12, 4); $wrap.Tag = 'card'
    $wrap.Controls.Add($lb)

    $bar = New-Object Windows.Forms.FlowLayoutPanel
    $bar.Dock = 'Bottom'; $bar.Height = 104; $bar.Padding = New-Object Windows.Forms.Padding(8, 6, 8, 6); $bar.WrapContents = $true
    $mk = { param($text, $w) $b = New-Object Windows.Forms.Button; $b.Text = $text; $b.Size = New-Object Drawing.Size($w, 36); $b.FlatStyle = 'Flat'; $b.Cursor = 'Hand'; $b.Margin = New-Object Windows.Forms.Padding(4); $bar.Controls.Add($b); $b }
    $bUp = & $mk '▲ למעלה' 100
    $bDown = & $mk '▼ למטה' 100
    $bTop = & $mk 'לראש הרשימה' 120
    $bEnd = & $mk 'לסוף הרשימה' 120
    $bAbc = & $mk 'סדר לפי א-ב' 140
    $bSave = & $mk 'שמירה' 120
    $bSave.Font = New-Object Drawing.Font('Segoe UI Semibold', 10)
    $bCancel = & $mk 'ביטול' 100

    $move = {
        param([int]$to)
        $i = $lb.SelectedIndex
        if ($i -lt 0) { return }
        $to = [Math]::Max(0, [Math]::Min($lb.Items.Count - 1, $to))
        if ($to -eq $i) { return }
        $x = $lb.Items[$i]; $lb.Items.RemoveAt($i); $lb.Items.Insert($to, $x); $lb.SelectedIndex = $to
    }
    $bUp.Add_Click({ & $move ($lb.SelectedIndex - 1) })
    $bDown.Add_Click({ & $move ($lb.SelectedIndex + 1) })
    $bTop.Add_Click({ & $move 0 })
    $bEnd.Add_Click({ & $move ($lb.Items.Count - 1) })
    $bAbc.Add_Click({
        $arr = @(); foreach ($x in $lb.Items) { $arr += $x }
        $objs = @($arr | Sort-Object -Property @{ Expression = { [string]$_.Title } } -Culture 'he-IL')
        $sel = $lb.SelectedItem
        $lb.BeginUpdate(); $lb.Items.Clear(); foreach ($x in $objs) { [void]$lb.Items.Add($x) }; $lb.EndUpdate()
        if ($sel) { $lb.SelectedItem = $sel }
    })
    # גרירה ושחרור
    $drag = @{ Index = -1 }
    $lb.Add_MouseDown({ param($s, $e) $drag.Index = $lb.IndexFromPoint($e.Location) })
    $lb.Add_MouseMove({
        param($s, $e)
        if ($e.Button -eq 'Left' -and $drag.Index -ge 0) { [void]$lb.DoDragDrop($lb.Items[$drag.Index], 'Move') }
    })
    $lb.Add_DragOver({ param($s, $e) $e.Effect = 'Move' })
    $lb.Add_DragDrop({
        param($s, $e)
        $pt = $lb.PointToClient((New-Object Drawing.Point($e.X, $e.Y)))
        $to = $lb.IndexFromPoint($pt); if ($to -lt 0) { $to = $lb.Items.Count - 1 }
        $lb.SelectedIndex = $drag.Index
        & $move $to
        $drag.Index = -1
    })
    $lb.Add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq 'Up') { & $move ($lb.SelectedIndex - 1); $e.Handled = $true }
        elseif ($e.Control -and $e.KeyCode -eq 'Down') { & $move ($lb.SelectedIndex + 1); $e.Handled = $true }
    })

    $bCancel.Add_Click({ $od.Close() })
    $bSave.Add_Click({
        if (-not (Confirm-CloseOtzaria 'יש לסגור אותה כדי לשמור את הסדר.')) { return }
        $ids = @(); foreach ($x in $lb.Items) { $ids += [string]$x.Id }
        # תוספים ללא ממשק (רקע בלבד) נשארים בסוף
        foreach ($x in (Get-PluginOrderList | Where-Object { $_.Headless })) { $ids += [string]$x.Id }
        try {
            Save-PluginOrder $ids
            Show-RTLMessageBox 'הסדר נשמר. הוא יופיע בפעם הבאה שאוצריא תיפתח.' 'סידור תוספים' | Out-Null
            $od.Close()
        } catch { Show-RTLMessageBox "שגיאה בשמירת הסדר: $($_.Exception.Message)" 'סידור תוספים' 'OK' 'Error' | Out-Null }
    })

    $od.Controls.Add($wrap); $od.Controls.Add($bar); $od.Controls.Add($hint)
    $od.Add_Shown({ Apply-FormTheme $od; if ($lb.Items.Count -gt 0) { $lb.SelectedIndex = 0 } })
    $od.AutoScaleMode = 'None'
    [void](Set-UiScale $od 560 680)
    [void]$od.ShowDialog()
}

# ---------------------------------------------------------------------
# תיקיית תוספים - להעברה למחשב ללא אינטרנט
# ---------------------------------------------------------------------
$script:PluginFolder = $null

# שם קובץ חוקי: מסירים תווים אסורים ב-Windows (\ / : * ? " < > |) ותווי בקרה
function Get-SafeFileName([string]$Name) {
    $n = $Name -replace '[\\/:*?"<>|\x00-\x1F]', '' -replace '\s+', ' '
    return $n.Trim().TrimEnd('.').Trim()
}

# קורא את פרטי התוסף מתוך הקובץ עצמו (manifest.json) - לא מסתמך על שם הקובץ
function Read-PluginPackage([string]$Path) {
    $z = $null
    try {
        $z = [IO.Compression.ZipFile]::OpenRead($Path)
        $e = $z.Entries | Where-Object { $_.FullName -eq 'manifest.json' } | Select-Object -First 1
        if (-not $e) { return $null }
        $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
        try { $txt = $sr.ReadToEnd().TrimStart([char]0xFEFF) } finally { $sr.Close() }
        $m = ConvertFrom-Json -InputObject $txt
        if (-not $m.id) { return $null }
        return [pscustomobject]@{ Id = [string]$m.id; Version = [string]$m.version; Name = [string]$m.name; Path = $Path; File = (Split-Path -Leaf $Path) }
    } catch { return $null } finally { if ($z) { $z.Dispose() } }
}

function Get-FolderPlugins([string]$Folder) {
    return @(Get-ChildItem -LiteralPath $Folder -Filter '*.otzplugin' -File -ErrorAction SilentlyContinue |
             ForEach-Object { Read-PluginPackage $_.FullName } | Where-Object { $_ })
}

# משאיר בתיקייה רק את הגירסה החדשה ביותר של כל תוסף; מחזיר כמה קבצים נמחקו
function Remove-OldFolderVersions([string]$Folder) {
    $removed = 0
    foreach ($g in (Get-FolderPlugins $Folder | Group-Object Id)) {
        if ($g.Count -lt 2) { continue }
        $best = $null
        foreach ($k in $g.Group) { if (-not $best -or (Compare-Ver $k.Version $best.Version) -gt 0) { $best = $k } }
        foreach ($k in $g.Group) { if ($k.Path -ne $best.Path) { Remove-Item -LiteralPath $k.Path -Force -ErrorAction SilentlyContinue; $removed++ } }
    }
    return $removed
}

function Select-PluginFolder([string]$Description, [bool]$AllowNew) {
    $fbd = New-Object Windows.Forms.FolderBrowserDialog
    $fbd.Description = $Description; $fbd.ShowNewFolderButton = $AllowNew
    $def = $script:PluginFolder
    if (-not $def) {
        $cand = Join-Path $script:WorkDir ('תוספים לאוצריא')
        $def = if (Test-Path -LiteralPath $cand) { $cand } else { $script:WorkDir }
    }
    if (Test-Path -LiteralPath $def) { $fbd.SelectedPath = $def }
    if ($fbd.ShowDialog() -ne 'OK') { return $null }
    $script:PluginFolder = $fbd.SelectedPath
    return $fbd.SelectedPath
}

# חלון בחירת תוספים להתקנה מתיקייה. מחזיר את החבילות שסומנו (או $null)
function Show-FolderInstallDialog([object[]]$Pkgs, $Installed) {
    $dlg = New-Object Windows.Forms.Form
    $dlg.Text = 'התקנה מתיקייה'
    $dlg.ClientSize = New-Object Drawing.Size(620, 600); $dlg.MinimumSize = New-Object Drawing.Size(520, 420)
    $dlg.StartPosition = 'CenterParent'; $dlg.RightToLeft = 'Yes'; $dlg.RightToLeftLayout = $true
    $dlg.MinimizeBox = $false; $dlg.Font = New-Object Drawing.Font('Segoe UI', 10); $dlg.Icon = $form.Icon
    $hint = New-Object Windows.Forms.Label
    $hint.Text = "התוספים שנמצאו בתיקייה. תוספים חדשים ועדכונים מסומנים מראש."
    $hint.Dock = 'Top'; $hint.Height = 40; $hint.TextAlign = 'MiddleCenter'
    $cl = New-Object Windows.Forms.CheckedListBox
    $cl.Dock = 'Fill'; $cl.Font = New-Object Drawing.Font('Segoe UI', 10.5); $cl.IntegralHeight = $false; $cl.BorderStyle = 'None'
    $cl.CheckOnClick = $true; $cl.DisplayMember = 'Text'
    $rows = @()
    foreach ($k in ($Pkgs | Sort-Object Name)) {
        $loc = if ($Installed.ContainsKey($k.Id)) { $Installed[$k.Id].Version } else { $null }
        if (-not $loc) { $state = 'חדש'; $def = $true }
        else {
            $c = Compare-Ver $k.Version $loc
            if ($c -gt 0) { $state = "עדכון (מותקנת $loc)"; $def = $true }
            elseif ($c -eq 0) { $state = 'כבר מותקן'; $def = $false }
            else { $state = "ישן מהמותקן ($loc)"; $def = $false }
        }
        $row = [pscustomobject]@{ Text = "$($k.Name)   ·   $($k.Version)   ·   $state"; Pkg = $k }
        [void]$cl.Items.Add($row, $def)
    }
    $wrap = New-Object Windows.Forms.Panel; $wrap.Dock = 'Fill'; $wrap.Padding = New-Object Windows.Forms.Padding(12, 4, 12, 4); $wrap.Tag = 'card'
    $wrap.Controls.Add($cl)
    $bar = New-Object Windows.Forms.FlowLayoutPanel
    $bar.Dock = 'Bottom'; $bar.Height = 58; $bar.Padding = New-Object Windows.Forms.Padding(8, 8, 8, 6)
    $mk = { param($text, $w) $b = New-Object Windows.Forms.Button; $b.Text = $text; $b.Size = New-Object Drawing.Size($w, 36); $b.FlatStyle = 'Flat'; $b.Cursor = 'Hand'; $b.Margin = New-Object Windows.Forms.Padding(4); $bar.Controls.Add($b); $b }
    $bAll = & $mk 'סמן הכל' 100
    $bNone = & $mk 'נקה סימון' 100
    $bOk = & $mk 'התקנת המסומנים' 160
    $bOk.Font = New-Object Drawing.Font('Segoe UI Semibold', 10)
    $bCancel = & $mk 'ביטול' 100
    $result = @{ Pkgs = $null }
    $bAll.Add_Click({ for ($i = 0; $i -lt $cl.Items.Count; $i++) { $cl.SetItemChecked($i, $true) } })
    $bNone.Add_Click({ for ($i = 0; $i -lt $cl.Items.Count; $i++) { $cl.SetItemChecked($i, $false) } })
    $bCancel.Add_Click({ $dlg.Close() })
    $bOk.Add_Click({
        $sel = @(); foreach ($x in $cl.CheckedItems) { $sel += $x.Pkg }
        if ($sel.Count -eq 0) { Show-RTLMessageBox 'לא סומנו תוספים.' 'התקנה מתיקייה' | Out-Null; return }
        $result.Pkgs = $sel; $dlg.Close()
    })
    $dlg.Controls.Add($wrap); $dlg.Controls.Add($bar); $dlg.Controls.Add($hint)
    $dlg.Add_Shown({ Apply-FormTheme $dlg })
    $dlg.AutoScaleMode = 'None'
    [void](Set-UiScale $dlg 620 640)
    [void]$dlg.ShowDialog()
    return $result.Pkgs
}

function Show-PluginsWindow {
    $pf = New-Object Windows.Forms.Form
    $pf.Text = 'תוספים לאוצריא'
    $pf.ClientSize = New-Object Drawing.Size(1044, 680); $pf.MinimumSize = New-Object Drawing.Size(880, 560)
    $pf.StartPosition = 'CenterParent'; $pf.RightToLeft = 'Yes'; $pf.RightToLeftLayout = $true
    $pf.Font = New-Object Drawing.Font('Segoe UI', 10)
    $pf.Icon = $form.Icon
    $W = 1044
    # Panel children use left-based coordinates. Controls are placed from the right edge by an
    # explicit layout pass (anchors + scaling misplace them): R = right-aligned, S = stretched, C = centered.
    $sc = 1.0
    $placement = @{}
    function Add-R($parent, $ctl, [int]$right, [int]$y, [int]$w, [int]$h, [string]$mode = 'R') {
        $ctl.Location = New-Object Drawing.Point(($W - $right - $w), $y); $ctl.Size = New-Object Drawing.Size($w, $h); $ctl.Anchor = 'Top, Left'
        $parent.Controls.Add($ctl)
        $placement[$ctl] = @{ Right = $right; Mode = $mode }
    }
    $layoutBars = {
        if ($pf.WindowState -eq 'Minimized') { return }
        foreach ($ctl in @($placement.Keys)) {
            $pl = $placement[$ctl]; $pw = $ctl.Parent.ClientSize.Width; $r = [int]($pl.Right * $sc)
            if ($pw -lt 300) { continue }
            switch ($pl.Mode) {
                'S' { $ctl.Left = $r; $ctl.Width = [Math]::Max(50, $pw - 2 * $r) }
                'C' { $ctl.Left = [int](($pw - $ctl.Width) / 2) }
                default { $ctl.Left = $pw - $r - $ctl.Width }
            }
        }
    }

    # Top bar
    $top = New-Object Windows.Forms.Panel; $top.Size = New-Object Drawing.Size($W, 134); $top.Tag = 'header'
    $lblFolder = New-Object Windows.Forms.Label; $lblFolder.Text = 'תיקיית תוספים (למחשב ללא אינטרנט):'; $lblFolder.TextAlign = 'MiddleLeft'; $lblFolder.Tag = 'strong'
    Add-R $top $lblFolder 12 96 300 30
    $btnDlFolder = New-Object Windows.Forms.Button; $btnDlFolder.Text = 'הורדת המסומנים לתיקייה'
    Add-R $top $btnDlFolder 320 94 215 32
    $btnChkFolder = New-Object Windows.Forms.Button; $btnChkFolder.Text = 'בדיקת עדכונים לתיקייה'
    Add-R $top $btnChkFolder 543 94 205 32
    $btnInstFolder = New-Object Windows.Forms.Button; $btnInstFolder.Text = 'התקנה מתיקייה'
    Add-R $top $btnInstFolder 756 94 150 32
    foreach ($b in @($btnDlFolder, $btnChkFolder, $btnInstFolder)) { $b.FlatStyle = 'Flat'; $b.Cursor = 'Hand' }
    $lblSearch = New-Object Windows.Forms.Label; $lblSearch.Text = 'חיפוש:'; $lblSearch.TextAlign = 'MiddleLeft'
    Add-R $top $lblSearch 12 14 60 26
    $txtSearch = New-Object Windows.Forms.TextBox
    Add-R $top $txtSearch 75 14 250 26
    $lblFilter = New-Object Windows.Forms.Label; $lblFilter.Text = 'הצג:'; $lblFilter.TextAlign = 'MiddleLeft'
    Add-R $top $lblFilter 340 14 45 26
    $cmbFilter = New-Object Windows.Forms.ComboBox; $cmbFilter.DropDownStyle = 'DropDownList'
    Add-R $top $cmbFilter 388 14 170 26
    [void]$cmbFilter.Items.AddRange(@('כל התוספים', 'מותקנים', 'עדכונים זמינים', 'לא מותקנים'))
    $cmbFilter.SelectedIndex = 0
    $chkExperimental = New-Object Windows.Forms.CheckBox; $chkExperimental.Text = 'הצג גם ניסיוניים'; $chkExperimental.Checked = $true
    Add-R $top $chkExperimental 575 14 160 26
    $btnOrder = New-Object Windows.Forms.Button; $btnOrder.Text = 'סידור התוספים באוצריא'; $btnOrder.FlatStyle = 'Flat'; $btnOrder.Cursor = 'Hand'
    $btnOrder.Location = New-Object Drawing.Point(12, 12); $btnOrder.Size = New-Object Drawing.Size(210, 32); $btnOrder.Anchor = 'Top, Left'; $top.Controls.Add($btnOrder)
    $btnOrder.Add_Click({ Show-PluginOrderDialog })
    $lblCounts = New-Object Windows.Forms.Label; $lblCounts.TextAlign = 'MiddleRight'
    $lblCounts.Location = New-Object Drawing.Point(12, 52); $lblCounts.Size = New-Object Drawing.Size(330, 26); $lblCounts.Anchor = 'Top, Left'; $top.Controls.Add($lblCounts)
    $lblPerm = New-Object Windows.Forms.Label; $lblPerm.Text = 'הרשאות:'; $lblPerm.TextAlign = 'MiddleLeft'; $lblPerm.Tag = 'strong'
    Add-R $top $lblPerm 12 52 70 26
    $rbDefault = New-Object Windows.Forms.RadioButton; $rbDefault.Text = 'ברירת מחדל (כמו באוצריא)'; $rbDefault.Checked = $true
    Add-R $top $rbDefault 85 52 230 26
    $rbAll = New-Object Windows.Forms.RadioButton; $rbAll.Text = 'כל ההרשאות שהתוסף מבקש'
    Add-R $top $rbAll 320 52 230 26
    $tt = New-Object Windows.Forms.ToolTip
    $tt.SetToolTip($rbDefault, "כמו בחלון ההרשאות של אוצריא: הכל מאושר חוץ מהפעלה עם פתיחת התוכנה, תרומות בהפעלה, פעילות ברקע וקריאת הלוח.`nבעדכון - ההחלטות הקודמות שלך נשמרות.")
    $tt.SetToolTip($rbAll, 'כל ההרשאות שהתוסף מבקש יאושרו, כולל הפעלה אוטומטית ופעילות ברקע.')
    $top.Dock = 'Top'

    # List
    $lv = New-Object Windows.Forms.ListView
    $lv.View = 'Details'; $lv.CheckBoxes = $true; $lv.FullRowSelect = $true; $lv.GridLines = $false; $lv.HideSelection = $false
    $lv.RightToLeftLayout = $true; $lv.Dock = 'Fill'; $lv.Font = New-Object Drawing.Font('Segoe UI', 10); $lv.BorderStyle = 'None'
    foreach ($col in @(@('שם התוסף', 230), @('מחבר', 130), @('גרסה אחרונה', 95), @('מותקנת', 85), @('מצב', 120), @('יציבות', 75), @('גודל', 75), @('הורדות', 70), @('עודכן', 90))) {
        [void]$lv.Columns.Add($col[0], $col[1])
    }
    # שורת הכותרות של Windows אינה קריאה במצב כהה ונשברת בציור עצמי מימין לשמאל,
    # ולכן היא מוסתרת ובמקומה שורת כותרות משלנו - מיושרת לעמודות, לחיצה ממיינת.
    $lv.HeaderStyle = 'None'
    $hdrPanel = New-Object Windows.Forms.Panel; $hdrPanel.Size = New-Object Drawing.Size($W, 30); $hdrPanel.Tag = 'header'
    $hdrLine = New-Object Windows.Forms.Label; $hdrLine.Dock = 'Bottom'; $hdrLine.Height = 1; $hdrLine.Tag = 'divider'
    $hdrPanel.Controls.Add($hdrLine)
    $hdrLabels = @()
    for ($ci = 0; $ci -lt $lv.Columns.Count; $ci++) {
        $hl = New-Object Windows.Forms.Label
        $hl.Text = $lv.Columns[$ci].Text; $hl.TextAlign = 'MiddleLeft'; $hl.Tag = 'strong'; $hl.Cursor = 'Hand'
        $hl.Font = New-Object Drawing.Font('Segoe UI Semibold', 9.5); $hl.AutoEllipsis = $true
        $hl.Padding = New-Object Windows.Forms.Padding(6, 0, 6, 0)
        $hl.Name = "col$ci"
        $hl.Add_Click({
            param($s, $e)
            $c = [int]($s.Name.Substring(3))
            if ($st.SortCol -eq $c) { $st.SortAsc = -not $st.SortAsc } else { $st.SortCol = $c; $st.SortAsc = $true }
            & $fillList
        })
        $hdrPanel.Controls.Add($hl)
        $hdrLabels += $hl
    }
    # בחלון מימין לשמאל העמודה הראשונה צמודה לקצה הימני של הרשימה
    $layoutHeader = {
        if ($pf.WindowState -eq 'Minimized' -or $hdrPanel.ClientSize.Width -lt 300) { return }
        $x = $lv.ClientSize.Width + ($hdrPanel.ClientSize.Width - $lv.Width)
        $h = $hdrPanel.ClientSize.Height - 1
        for ($ci = 0; $ci -lt $hdrLabels.Count; $ci++) {
            $w = $lv.Columns[$ci].Width
            $x -= $w
            $hdrLabels[$ci].SetBounds($x, 0, $w, $h)
        }
    }
    $hdrPanel.Dock = 'Top'

    # Bottom area
    $bottom = New-Object Windows.Forms.Panel; $bottom.Size = New-Object Drawing.Size($W, 150); $bottom.Tag = 'card'
    $lblDesc = New-Object Windows.Forms.Label; $lblDesc.TextAlign = 'TopLeft'
    Add-R $bottom $lblDesc 12 8 ($W - 24) 46 'S'
    $btnSelAll = New-Object Windows.Forms.Button; $btnSelAll.Text = 'סמן הכל'
    Add-R $bottom $btnSelAll 12 60 110 34
    $btnSelUpdates = New-Object Windows.Forms.Button; $btnSelUpdates.Text = 'סמן עדכונים'
    Add-R $bottom $btnSelUpdates 130 60 130 34
    $btnSelNone = New-Object Windows.Forms.Button; $btnSelNone.Text = 'נקה סימון'
    Add-R $bottom $btnSelNone 268 60 110 34
    $btnInstall = New-Object Windows.Forms.Button; $btnInstall.Text = 'התקנה / עדכון'
    $btnInstall.Font = New-Object Drawing.Font('Segoe UI Semibold', 10.5)
    Add-R $bottom $btnInstall 386 60 240 34
    $btnRemove = New-Object Windows.Forms.Button; $btnRemove.Text = 'הסרה'
    Add-R $bottom $btnRemove 634 60 150 34
    $btnClose = New-Object Windows.Forms.Button; $btnClose.Text = 'סגירה'
    $btnClose.Location = New-Object Drawing.Point(12, 60); $btnClose.Size = New-Object Drawing.Size(112, 34); $btnClose.Anchor = 'Top, Left'; $bottom.Controls.Add($btnClose)
    $btnReload = New-Object Windows.Forms.Button; $btnReload.Text = 'רענון'
    $btnReload.Location = New-Object Drawing.Point(132, 60); $btnReload.Size = New-Object Drawing.Size(110, 34); $btnReload.Anchor = 'Top, Left'; $bottom.Controls.Add($btnReload)
    foreach ($b in @($btnSelAll, $btnSelUpdates, $btnSelNone, $btnInstall, $btnRemove, $btnReload, $btnClose)) { $b.FlatStyle = 'Flat'; $b.Cursor = 'Hand' }
    $tipB = New-Object Windows.Forms.ToolTip
    $tipB.SetToolTip($btnSelAll, 'סימון כל התוספים שמוצגים כעת ברשימה (לפי החיפוש והסינון)')
    $tipB.SetToolTip($btnSelUpdates, 'סימון כל התוספים שיש להם עדכון')
    $tipB.SetToolTip($btnRemove, 'הסרת התוספים המסומנים שמותקנים במחשב')
    $pbar = New-Object ModernProgressBar
    Add-R $bottom $pbar 12 102 ($W - 24) 18 'S'
    $lblStat = New-Object Windows.Forms.Label; $lblStat.TextAlign = 'MiddleCenter'
    Add-R $bottom $lblStat 12 122 ($W - 24) 22 'S'
    $bottom.Dock = 'Bottom'

    $pf.Controls.Add($lv); $pf.Controls.Add($hdrPanel); $pf.Controls.Add($bottom); $pf.Controls.Add($top)

    $st = @{ Catalog = @(); Installed = @{}; Job = $null; Plan = $null; Rows = @(); Ok = 0; Fail = 0; Fails = @(); SortCol = -1; SortAsc = $true; ColNames = @() }
    $st.ColNames = @($lv.Columns | ForEach-Object { $_.Text })
    $appVer = $null
    if ($script:Install) { $appVer = $script:Install.Version }

    $stateOf = {
        param($p)
        $uid = [string]$p.pluginUid
        $loc = $null
        if ($uid) { if ($st.Installed.ContainsKey($uid)) { $loc = $st.Installed[$uid] } }
        else { foreach ($v in $st.Installed.Values) { if ($v.Name -eq [string]$p.name) { $loc = $v; break } } }
        $compat = $true
        if ($appVer) {
            if ($p.compatibleWith -and (Compare-Ver $appVer ([string]$p.compatibleWith)) -lt 0) { $compat = $false }
            if ($p.maxAppVersion -and (Compare-Ver $appVer ([string]$p.maxAppVersion)) -gt 0) { $compat = $false }
        }
        if ($p.LocalOnly) { return @{ Local = $loc; Compat = $true; State = 'current' } }
        $state = if (-not $loc) { 'none' } elseif ((Compare-Ver ([string]$p.version) $loc.Version) -gt 0) { 'update' } else { 'current' }
        return @{ Local = $loc; Compat = $compat; State = $state }
    }

    $fillList = {
        $lv.BeginUpdate()
        $lv.Items.Clear()
        $q = $txtSearch.Text.Trim()
        $nUpd = 0; $nInst = 0
        $t = Get-Theme
        $entries = New-Object System.Collections.Generic.List[object]
        foreach ($p in $st.Catalog) {
            $s = & $stateOf $p
            if ($s.State -eq 'update') { $nUpd++ }; if ($s.Local) { $nInst++ }
            if (-not $chkExperimental.Checked -and $p.status -eq 'experimental' -and -not $s.Local) { continue }
            $fi = $cmbFilter.SelectedIndex
            if ($fi -eq 1 -and -not $s.Local) { continue }
            if ($fi -eq 2 -and $s.State -ne 'update') { continue }
            if ($fi -eq 3 -and $s.Local) { continue }
            if ($q) {
                $hay = "$($p.name) $($p.author) $($p.shortDescription) $($p.pluginUid) $(@($p.tags) -join ' ')"
                if ($hay.IndexOf($q, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
            }
            $stText = if ($p.LocalOnly) { 'מותקן (לא מהחנות)' } elseif (-not $s.Compat) { 'לא תואם לגרסה' } elseif ($s.State -eq 'update') { 'עדכון זמין' } elseif ($s.State -eq 'current') { 'מותקן ✓' } else { 'לא מותקן' }
            $stab = switch ([string]$p.status) { 'stable' { 'יציב' } 'beta' { 'בטא' } 'experimental' { 'ניסיוני' } default { [string]$p.status } }
            $it = New-Object Windows.Forms.ListViewItem([string]$p.name)
            [void]$it.SubItems.Add([string]$p.author)
            [void]$it.SubItems.Add([string]$p.version)
            [void]$it.SubItems.Add($(if ($s.Local) { [string]$s.Local.Version } else { '—' }))
            [void]$it.SubItems.Add($stText)
            [void]$it.SubItems.Add($stab)
            [void]$it.SubItems.Add($(if ($p.LocalOnly) { '' } else { Format-Size ([double]$p.pluginFileSize) }))
            [void]$it.SubItems.Add($(if ($p.LocalOnly) { '' } else { '{0:N0}' -f [double]$p.downloadCount }))
            [void]$it.SubItems.Add([string]$p.updatedAt)
            $it.Tag = $p
            if (-not $s.Compat) { $it.ForeColor = [Drawing.Color]::Gray }
            elseif ($s.State -eq 'update') { $it.ForeColor = $t.Warn }
            elseif ($s.State -eq 'current') { $it.ForeColor = $t.Good }
            else { $it.ForeColor = $t.TextColor }
            if ($st.Checked -and $st.Checked.Contains([string]$p.id)) { $it.Checked = $true }
            # מפתח מיון לפי העמודה שנבחרה (מספרים וגרסאות לפי ערך, לא לפי אותיות)
            $key = switch ($st.SortCol) {
                0 { [string]$p.name }
                1 { [string]$p.author }
                2 { Get-VersionSortKey ([string]$p.version) }
                3 { if ($s.Local) { Get-VersionSortKey ([string]$s.Local.Version) } else { '' } }
                4 { $stText }
                5 { switch ([string]$p.status) { 'stable' { 0 } 'beta' { 1 } 'experimental' { 2 } default { 3 } } }
                6 { [double]$p.pluginFileSize }
                7 { [double]$p.downloadCount }
                8 { [string]$p.updatedAt }
                default { $entries.Count }
            }
            $entries.Add([pscustomobject]@{ Item = $it; Key = $key; Order = $entries.Count })
        }
        $sorted = if ($st.SortCol -ge 0) {
            $entries | Sort-Object -Property @{ Expression = 'Key'; Descending = (-not $st.SortAsc) }, @{ Expression = 'Order'; Descending = $false }
        } else { $entries }
        foreach ($en in $sorted) { [void]$lv.Items.Add($en.Item) }
        for ($c = 0; $c -lt $lv.Columns.Count; $c++) {
            $lv.Columns[$c].Text = $st.ColNames[$c] + $(if ($c -eq $st.SortCol) { if ($st.SortAsc) { ' ▲' } else { ' ▼' } } else { '' })
            $hdrLabels[$c].Text = $lv.Columns[$c].Text
        }
        $lv.EndUpdate()
        $lblCounts.Text = "$($st.Catalog.Count) תוספים  ·  $nInst מותקנים  ·  $nUpd עדכונים"
    }

    # תוספים מותקנים שאינם בחנות - כדי שאפשר יהיה להסיר גם אותם
    $addLocalOnly = {
        $st.Catalog = @($st.Catalog | Where-Object { -not $_.LocalOnly })
        $known = @{}
        foreach ($p in $st.Catalog) { if ($p.pluginUid) { $known[[string]$p.pluginUid] = 1 }; $known['name:' + [string]$p.name] = 1 }
        foreach ($k in @($st.Installed.Keys)) {
            $v = $st.Installed[$k]
            if ($known.ContainsKey($k) -or $known.ContainsKey('name:' + $v.Name)) { continue }
            $st.Catalog += [pscustomobject]@{ id = "local:$k"; pluginUid = $k; name = $v.Name; author = ''; version = $v.Version; status = ''
                shortDescription = 'תוסף מותקן שאינו מופיע בחנות התוספים'; description = ''; tags = @(); pluginFileSize = 0; downloadCount = 0
                updatedAt = ''; compatibleWith = ''; maxAppVersion = $null; requiresNetwork = $false; slug = ''; downloadUrl = $null; sha256 = ''; LocalOnly = $true }
        }
    }

    $updateButtons = {
        $nSel = $st.Checked.Count
        $nInstSel = @($st.Catalog | Where-Object { $st.Checked.Contains([string]$_.id) -and (& $stateOf $_).Local }).Count
        $btnInstall.Text = $(if ($nSel) { "התקנה / עדכון ($nSel)" } else { 'התקנה / עדכון' })
        $btnRemove.Text = $(if ($nInstSel) { "הסרה ($nInstSel)" } else { 'הסרה' })
    }

    $loadData = {
        $pf.Cursor = 'WaitCursor'; $lblStat.Text = 'טוען את רשימת התוספים...'; [Windows.Forms.Application]::DoEvents()
        try {
            $raw = ConvertFrom-Json -InputObject (Get-WebText "$($script:PluginsSite)/api/plugins")
            $st.Catalog = @(@($raw) | Where-Object { $_ -and $_.downloadUrl } | Sort-Object @{ Expression = { -[double]$_.downloadCount } })
            $lblStat.Text = ''
        } catch {
            $st.Catalog = @(); $lblStat.Text = "שגיאה בטעינת רשימת התוספים: $($_.Exception.Message)"
        }
        $st.Installed = Get-InstalledPlugins
        & $addLocalOnly
        & $fillList
        $pf.Cursor = 'Default'
    }

    $st.Checked = New-Object System.Collections.Generic.HashSet[string]
    $lv.Add_ItemCheck({
        param($s, $e)
        if ($st.Job) { $e.NewValue = $e.CurrentValue }   # בזמן התקנה אין לשנות סימונים
    })
    $lv.Add_ItemChecked({
        param($s, $e)
        $p = $e.Item.Tag
        if (-not $p) { return }
        if ($e.Item.Checked) { [void]$st.Checked.Add([string]$p.id) } else { [void]$st.Checked.Remove([string]$p.id) }
        & $updateButtons
    })

    $lv.Add_SelectedIndexChanged({
        if ($lv.SelectedItems.Count -eq 0) { return }
        $p = $lv.SelectedItems[0].Tag
        $s = & $stateOf $p
        $txt = [string]$p.shortDescription
        if (-not $txt) { $txt = [string]$p.description }
        $extra = "מזהה: $($p.pluginUid)   ·   דורש אוצריא $($p.compatibleWith) ומעלה"
        if ($p.requiresNetwork) { $extra += '   ·   דורש אינטרנט' }
        $lblDesc.Text = "$($p.name): $txt`n$extra"
    })

    $lv.Add_DoubleClick({
        if ($lv.SelectedItems.Count -eq 0) { return }
        $p = $lv.SelectedItems[0].Tag
        if ($p.slug) { Start-Process "$($script:PluginsSite)/plugins/$($p.slug)" }
    })

    $lv.Add_ColumnClick({
        param($s, $e)
        if ($st.SortCol -eq $e.Column) { $st.SortAsc = -not $st.SortAsc } else { $st.SortCol = $e.Column; $st.SortAsc = $true }
        & $fillList
    })
    $txtSearch.Add_TextChanged({ & $fillList })
    $cmbFilter.Add_SelectedIndexChanged({ & $fillList })
    $chkExperimental.Add_CheckedChanged({ & $fillList })
    $btnReload.Add_Click({ if (-not $st.Job) { & $loadData } })
    $btnClose.Add_Click({ $pf.Close() })
    $btnSelAll.Add_Click({
        if ($st.Job) { return }
        foreach ($it in $lv.Items) { if ($it.Tag) { [void]$st.Checked.Add([string]$it.Tag.id) } }
        & $fillList
        & $updateButtons
    })
    $btnSelNone.Add_Click({ $st.Checked.Clear(); & $fillList; & $updateButtons })
    $btnSelUpdates.Add_Click({
        foreach ($p in $st.Catalog) { $s = & $stateOf $p; if ($s.State -eq 'update' -and $s.Compat) { [void]$st.Checked.Add([string]$p.id) } }
        & $fillList
        & $updateButtons
    })

    $timerP = New-Object Windows.Forms.Timer; $timerP.Interval = 250
    $timerP.Add_Tick({
        if (-not $st.Job) { return }
        $terminal = $st.Job.State -in @('Completed', 'Failed', 'Stopped')
        if ($st.Mode -eq 'download') {
            foreach ($line in @(Receive-Job -Job $st.Job -ErrorAction SilentlyContinue)) {
                $s = [string]$line
                if ($s -match '^PROGRESS:(\d+):(\d+)$') {
                    $pbar.Value = [int]$matches[2]
                    $lblStat.Text = "הורדו $(Format-Size ([double]$matches[1])) מתוך $(Format-Size $st.DlTotal)"
                }
                elseif ($s.StartsWith('STAGE:')) { $lblStat.Text = $s.Substring(6) }
                elseif ($s -eq 'DONE' -or $s.StartsWith('ERROR:') -or $s -eq 'STOPPED') {
                    Remove-Job -Job $st.Job -Force -ErrorAction SilentlyContinue; $st.Job = $null
                    & $setBusy $false
                    if ($s -eq 'DONE') {
                        $pbar.Value = 100
                        $removed = Remove-OldFolderVersions $st.DlFolder
                        $lblStat.Text = "ההורדה לתיקייה הושלמה ($($st.DlCount) תוספים)"
                        $inFolder = @(Get-FolderPlugins $st.DlFolder).Count
                        $m = "הורדו $($st.DlCount) תוספים לתיקייה:`n$($st.DlFolder)`n`nסך הכל בתיקייה: $inFolder תוספים."
                        if ($removed) { $m += "`nגירסאות ישנות שנמחקו: $removed." }
                        $m += "`n`nלפתוח את התיקייה?"
                        if ((Show-RTLMessageBox $m 'הורדה לתיקייה' 'YesNo' 'Information') -eq 'Yes') { Start-Process explorer.exe -ArgumentList "`"$($st.DlFolder)`"" }
                    } else {
                        $err = if ($s.StartsWith('ERROR:')) { $s.Substring(6) } else { 'ההורדה הופסקה' }
                        $lblStat.Text = 'ההורדה לתיקייה לא הושלמה'
                        Show-RTLMessageBox "שגיאה בהורדה לתיקייה:`n$err`n`nאפשר להוריד שוב - קבצים שכבר ירדו תקינים לא יורדו מחדש." 'הורדה לתיקייה' 'OK' 'Error' | Out-Null
                    }
                    return
                }
            }
            if ($st.Job -and $terminal) {
                Remove-Job -Job $st.Job -Force -ErrorAction SilentlyContinue; $st.Job = $null
                & $setBusy $false
                $lblStat.Text = 'תהליך ההורדה הסתיים באופן לא צפוי'
            }
            return
        }
        foreach ($line in @(Receive-Job -Job $st.Job -ErrorAction SilentlyContinue)) {
            $s = [string]$line
            if ($s -match '^ITEM:(\d+):(.*)$') {
                $k = [int]$matches[1]; $row = $st.Rows[$k]
                $lblStat.Text = "($($k + 1)/$($st.Rows.Count)) $($matches[2]): $($row.name)"
                $pbar.Value = [int]($k * 100 / [Math]::Max(1, $st.Rows.Count))
            } elseif ($s -match '^OK:(\d+):') { $st.Ok++ }
            elseif ($s -match '^FAIL:(\d+):(.*)$') { $st.Fail++; $st.Fails += "• $($st.Rows[[int]$matches[1]].name): $($matches[2])" }
            elseif ($s -eq 'DONE') {
                Remove-Job -Job $st.Job -Force -ErrorAction SilentlyContinue; $st.Job = $null
                $pbar.Value = 100
                & $setBusy $false
                $lblStat.Text = "הסתיים: $($st.Ok) הותקנו/עודכנו" + $(if ($st.Fail) { ", $($st.Fail) נכשלו" } else { '' })
                $st.Checked.Clear()
                $st.Installed = Get-InstalledPlugins; & $addLocalOnly; & $fillList; & $updateButtons
                $msg = "$($st.Ok) תוספים הותקנו או עודכנו."
                if ($st.Fail) { $msg += "`n`n$($st.Fail) נכשלו:`n" + ($st.Fails -join "`n") }
                if ($script:Install -and $st.Ok -gt 0) {
                    $msg += "`n`nלפתוח את אוצריא עכשיו?"
                    if ((Show-RTLMessageBox $msg 'התקנת תוספים' 'YesNo' 'Information') -eq 'Yes') { Start-Process -FilePath $script:Install.Exe }
                } else { Show-RTLMessageBox $msg 'התקנת תוספים' 'OK' $(if ($st.Fail) { 'Warning' } else { 'Information' }) | Out-Null }
                return
            }
        }
        if ($st.Job -and $terminal) {
            Remove-Job -Job $st.Job -Force -ErrorAction SilentlyContinue; $st.Job = $null
            & $setBusy $false
            $lblStat.Text = 'תהליך ההתקנה הסתיים באופן לא צפוי'
        }
    })

    $setBusy = {
        param([bool]$busy)
        foreach ($c in @($btnInstall, $btnRemove, $btnReload, $btnSelAll, $btnSelUpdates, $btnSelNone, $btnOrder, $rbAll, $rbDefault, $btnDlFolder, $btnChkFolder, $btnInstFolder)) { $c.Enabled = -not $busy }
    }

    # התקנה ברקע (מהחנות או מתיקייה) - $rows לתצוגת השמות, $items לתכנית ההתקנה
    $startInstallJob = {
        param([object[]]$rows, [object[]]$items)
        $plan = @{ DataRoot = (Get-DataRoot); AppVersion = $appVer; AllPermissions = $rbAll.Checked; Baseline = $script:PluginBaseline; DefaultOff = $script:PluginDefaultOff; Items = $items }
        $json = ConvertTo-Json -InputObject $plan -Depth 5 -Compress
        $st.Mode = 'install'; $st.Rows = $rows; $st.Ok = 0; $st.Fail = 0; $st.Fails = @()
        & $setBusy $true
        $pbar.Value = 0; $lblStat.Text = 'מתחיל...'
        $st.Job = Start-Job -ScriptBlock $script:PluginWorker -ArgumentList $json, $script:SqliteCs
    }

    $btnInstall.Add_Click({
        if ($st.Job) { return }
        $sel = @($st.Catalog | Where-Object { $st.Checked.Contains([string]$_.id) })
        if ($sel.Count -eq 0) { Show-RTLMessageBox 'לא סומנו תוספים. סמן תוסף אחד או יותר ברשימה.' 'תוספים' | Out-Null; return }
        $rows = @($sel | Where-Object { $_.downloadUrl -and (& $stateOf $_).Compat })
        $skipped = @($sel | Where-Object { -not ($rows -contains $_) })
        if ($skipped.Count -gt 0) {
            $m = "התוספים הבאים לא יותקנו (לא תואמים לגרסת אוצריא $appVer, או שאינם בחנות):`n" + (($skipped | ForEach-Object { '• ' + $_.name }) -join "`n")
            if ($rows.Count -eq 0) { Show-RTLMessageBox $m 'תוספים' 'OK' 'Warning' | Out-Null; return }
            Show-RTLMessageBox $m 'תוספים' 'OK' 'Information' | Out-Null
        }
        $dataRoot = Get-DataRoot
        if (-not (Test-Path -LiteralPath (Join-Path $dataRoot 'databases\plugins_host.db'))) {
            Show-RTLMessageBox "לא נמצא מסד התוספים של אוצריא:`n$dataRoot\databases\plugins_host.db`n`nיש להפעיל את אוצריא לפחות פעם אחת לפני התקנת תוספים." 'תוספים' 'OK' 'Warning' | Out-Null; return
        }
        $nUpd = @($rows | Where-Object { (& $stateOf $_).Local }).Count
        $permText = if ($rbAll.Checked) { 'כל ההרשאות שהתוספים מבקשים יאושרו.' } else { 'ההרשאות יאושרו כברירת המחדל של אוצריא (הפעלה אוטומטית, פעילות ברקע וקריאת לוח - כבויות).' }
        $msg = "יותקנו $($rows.Count) תוספים ($($rows.Count - $nUpd) חדשים, $nUpd עדכונים).`n$permText`n`nלהמשיך?"
        if ((Show-RTLMessageBox $msg 'התקנת תוספים' 'YesNo' 'Question') -ne 'Yes') { return }
        if (-not (Confirm-CloseOtzaria 'יש לסגור אותה כדי לרשום את התוספים.')) { return }
        $items = @()
        foreach ($p in $rows) {
            $url = [string]$p.downloadUrl; if ($url.StartsWith('/')) { $url = $script:PluginsSite + $url }
            $items += @{ Uid = [string]$p.pluginUid; Url = $url; Sha256 = [string]$p.sha256; Size = [long]$p.pluginFileSize }
        }
        & $startInstallJob $rows $items
    })

    # ---------- תיקיית תוספים ----------
    $catalogFor = {
        param($pkg)
        foreach ($p in $st.Catalog) {
            if ($p.LocalOnly) { continue }
            if (([string]$p.pluginUid -and [string]$p.pluginUid -eq $pkg.Id) -or (-not $p.pluginUid -and [string]$p.name -eq $pkg.Name)) { return $p }
        }
        return $null
    }

    # הורדה לתיקייה ברקע: שם הקובץ = שם התוסף בעברית + גירסה; כל קובץ נבדק (SHA256)
    $downloadToFolder = {
        param([object[]]$rows, [string]$folder)
        $items = @(); $total = 0L
        foreach ($p in $rows) {
            $url = [string]$p.downloadUrl; if ($url.StartsWith('/')) { $url = $script:PluginsSite + $url }
            $fname = Get-SafeFileName "$($p.name) $($p.version)"
            if (-not $fname) { $fname = Get-SafeFileName "$($p.pluginUid) $($p.version)" }
            $items += @{ Url = $url; Path = (Join-Path $folder "$fname.otzplugin"); Size = [long]$p.pluginFileSize; Sha256 = [string]$p.sha256 }
            $total += [long]$p.pluginFileSize
        }
        $plan = @{ Items = $items; Assemble = @() }
        $json = ConvertTo-Json -InputObject $plan -Depth 6 -Compress
        $flag = Join-Path $env:TEMP 'otzaria_updater_stop_plugins.flag'
        Remove-Item -LiteralPath $flag -Force -ErrorAction SilentlyContinue
        $st.Mode = 'download'; $st.DlFolder = $folder; $st.DlTotal = $total; $st.DlCount = $rows.Count
        & $setBusy $true
        $pbar.Value = 0; $lblStat.Text = 'מתחיל הורדה לתיקייה...'
        $st.Job = Start-Job -ScriptBlock $script:DownloadWorker -ArgumentList $json, $flag
    }

    $btnDlFolder.Add_Click({
        if ($st.Job) { return }
        $rows = @($st.Catalog | Where-Object { $st.Checked.Contains([string]$_.id) -and $_.downloadUrl })
        if ($rows.Count -eq 0) { Show-RTLMessageBox "לא סומנו תוספים.`nכדי להוריד את כל התוספים לחץ קודם על ""סמן הכל""." 'הורדה לתיקייה' | Out-Null; return }
        $folder = Select-PluginFolder 'בחר תיקייה להורדת התוספים (אפשר ליצור תיקייה חדשה)' $true
        if (-not $folder) { return }
        & $downloadToFolder $rows $folder
    })

    $btnChkFolder.Add_Click({
        if ($st.Job) { return }
        if (@($st.Catalog | Where-Object { -not $_.LocalOnly }).Count -eq 0) { Show-RTLMessageBox 'בדיקת עדכונים דורשת חיבור לאינטרנט (רשימת התוספים לא נטענה).' 'בדיקת עדכונים' 'OK' 'Warning' | Out-Null; return }
        $folder = Select-PluginFolder 'בחר את תיקיית התוספים לבדיקה' $false
        if (-not $folder) { return }
        $pf.Cursor = 'WaitCursor'; $pk = @(Get-FolderPlugins $folder); $pf.Cursor = 'Default'
        if ($pk.Count -eq 0) { Show-RTLMessageBox "לא נמצאו קבצי תוספים (.otzplugin) בתיקייה:`n$folder" 'בדיקת עדכונים' 'OK' 'Warning' | Out-Null; return }
        $upd = @(); $cur = 0; $unknown = @()
        foreach ($k in $pk) {
            $c = & $catalogFor $k
            if (-not $c) { $unknown += $k; continue }
            if ((Compare-Ver ([string]$c.version) $k.Version) -gt 0) { $upd += [pscustomobject]@{ Pkg = $k; Cat = $c } } else { $cur++ }
        }
        $msg = "בתיקייה נמצאו $($pk.Count) תוספים:`n• $cur מעודכנים`n• $($upd.Count) עם גירסה חדשה יותר"
        if ($unknown.Count) { $msg += "`n• $($unknown.Count) שאינם בחנות (יישארו כמו שהם)" }
        if ($upd.Count -eq 0) { Show-RTLMessageBox ($msg + "`n`nהתיקייה מעודכנת.") 'בדיקת עדכונים' | Out-Null; return }
        $list = ($upd | Select-Object -First 20 | ForEach-Object { "• $($_.Pkg.Name):  $($_.Pkg.Version) ← $($_.Cat.version)" }) -join "`n"
        if ($upd.Count -gt 20) { $list += "`n... ועוד $($upd.Count - 20)" }
        $msg += "`n`n$list`n`nלהוריד את הגירסאות החדשות לתיקייה?`n(הגירסאות הישנות יימחקו מהתיקייה)"
        if ((Show-RTLMessageBox $msg 'בדיקת עדכונים' 'YesNo' 'Question') -ne 'Yes') { return }
        & $downloadToFolder @($upd | ForEach-Object { $_.Cat }) $folder
    })

    $btnInstFolder.Add_Click({
        if ($st.Job) { return }
        $dataRoot = Get-DataRoot
        if (-not (Test-Path -LiteralPath (Join-Path $dataRoot 'databases\plugins_host.db'))) {
            Show-RTLMessageBox "לא נמצא מסד התוספים של אוצריא.`nיש להפעיל את אוצריא לפחות פעם אחת לפני התקנת תוספים." 'התקנה מתיקייה' 'OK' 'Warning' | Out-Null; return
        }
        $folder = Select-PluginFolder 'בחר את התיקייה שבה נמצאים קבצי התוספים' $false
        if (-not $folder) { return }
        $pf.Cursor = 'WaitCursor'; $pk = @(Get-FolderPlugins $folder); $pf.Cursor = 'Default'
        if ($pk.Count -eq 0) { Show-RTLMessageBox "לא נמצאו קבצי תוספים (.otzplugin) בתיקייה:`n$folder" 'התקנה מתיקייה' 'OK' 'Warning' | Out-Null; return }
        # אם יש כמה גירסאות של אותו תוסף - החדשה ביותר
        $best = @{}
        foreach ($k in $pk) { if (-not $best.ContainsKey($k.Id) -or (Compare-Ver $k.Version $best[$k.Id].Version) -gt 0) { $best[$k.Id] = $k } }
        $st.Installed = Get-InstalledPlugins
        $sel = Show-FolderInstallDialog @($best.Values) $st.Installed
        if (-not $sel) { return }
        $permText = if ($rbAll.Checked) { 'כל ההרשאות שהתוספים מבקשים יאושרו.' } else { 'ההרשאות יאושרו כברירת המחדל של אוצריא.' }
        if ((Show-RTLMessageBox "יותקנו $(@($sel).Count) תוספים מהתיקייה.`n$permText`n`nלהמשיך?" 'התקנה מתיקייה' 'YesNo' 'Question') -ne 'Yes') { return }
        if (-not (Confirm-CloseOtzaria 'יש לסגור אותה כדי לרשום את התוספים.')) { return }
        $rows = @(); $items = @()
        foreach ($k in $sel) { $rows += [pscustomobject]@{ name = $k.Name }; $items += @{ Uid = $k.Id; LocalPath = $k.Path; Url = ''; Sha256 = '' } }
        & $startInstallJob $rows $items
    })

    $btnRemove.Add_Click({
        if ($st.Job) { return }
        $rows = @($st.Catalog | Where-Object { $st.Checked.Contains([string]$_.id) -and (& $stateOf $_).Local })
        if ($rows.Count -eq 0) { Show-RTLMessageBox 'לא סומנו תוספים מותקנים להסרה.' 'הסרת תוספים' | Out-Null; return }
        $names = ($rows | Select-Object -First 15 | ForEach-Object { '• ' + $_.name }) -join "`n"
        if ($rows.Count -gt 15) { $names += "`n... ועוד $($rows.Count - 15)" }
        $msg = "להסיר $($rows.Count) תוספים?`n`n$names`n`nיימחקו גם הנתונים וההגדרות שהתוספים שמרו. לא ניתן לבטל פעולה זו."
        if ((Show-RTLMessageBox $msg 'הסרת תוספים' 'YesNo' 'Warning') -ne 'Yes') { return }
        if (-not (Confirm-CloseOtzaria 'יש לסגור אותה כדי להסיר תוספים.')) { return }
        $ids = @()
        foreach ($p in $rows) {
            $lid = $(if ($p.pluginUid) { [string]$p.pluginUid } else { @($st.Installed.Keys | Where-Object { $st.Installed[$_].Name -eq [string]$p.name })[0] })
            if ($lid) { $ids += $lid }
        }
        $pf.Cursor = 'WaitCursor'; $lblStat.Text = 'מסיר תוספים...'; [Windows.Forms.Application]::DoEvents()
        try { $r = Remove-OtzariaPlugins $ids } catch { $r = @{ Ok = 0; Fails = @("• $($_.Exception.Message)") } }
        $pf.Cursor = 'Default'
        $st.Checked.Clear()
        $st.Installed = Get-InstalledPlugins; & $addLocalOnly; & $fillList; & $updateButtons
        $lblStat.Text = "הוסרו $($r.Ok) תוספים" + $(if ($r.Fails.Count) { ", $($r.Fails.Count) נכשלו" } else { '' })
        $m = "$($r.Ok) תוספים הוסרו."
        if ($r.Fails.Count) { $m += "`n`nנכשלו:`n" + ($r.Fails -join "`n") }
        Show-RTLMessageBox $m 'הסרת תוספים' 'OK' $(if ($r.Fails.Count) { 'Warning' } else { 'Information' }) | Out-Null
    })

    $pf.Add_FormClosing({
        param($s, $e)
        if ($st.Job) {
            Show-RTLMessageBox 'הפעולה עדיין רצה (התקנה או הורדה). יש להמתין לסיומה.' 'תוספים' | Out-Null
            $e.Cancel = $true
        }
    })
    $pf.Add_Shown({ Apply-FormTheme $pf; $timerP.Start(); & $loadData })
    foreach ($pn in @($top, $bottom, $hdrPanel)) { foreach ($c in $pn.Controls) { if ($c -is [Windows.Forms.Label]) { $c.AutoEllipsis = $true } } }
    $pf.AutoScaleMode = 'None'
    $sc = Set-UiScale $pf 1060 760
    foreach ($col in $lv.Columns) { $col.Width = [int]($col.Width * $sc) }
    $pf.Add_Layout({ & $layoutBars; & $layoutHeader })
    $pf.Add_Resize({ & $layoutBars; & $layoutHeader })
    $top.Add_Resize({ & $layoutBars })
    $bottom.Add_Resize({ & $layoutBars })
    $lv.Add_Resize({ & $layoutHeader })
    & $layoutBars; & $layoutHeader
    [void]$pf.ShowDialog()
    $timerP.Stop(); $timerP.Dispose()
}

$btnPlugins.Add_Click({ Show-PluginsWindow })

# =====================================================================
# Timer, startup, shutdown
# =====================================================================
$timer = New-Object Windows.Forms.Timer
$timer.Interval = 250
$script:LastDarkCheck = [DateTime]::MinValue
$timer.Add_Tick({
    if (-not $script:DarkModeManual -and ([DateTime]::Now - $script:LastDarkCheck).TotalSeconds -ge 3) {
        $script:LastDarkCheck = [DateTime]::Now
        $d = Get-WindowsDarkMode
        if ($d -ne $script:DarkMode) { $script:DarkMode = $d; Apply-FormTheme $form; Update-AllUI }
    }
    Process-ActiveJob
})

$form.Add_Shown({
    Apply-FormTheme $form
    $timer.Start()
    Refresh-Data
})

$form.Add_FormClosing({
    param($s, $e)
    if ($script:Active) {
        $r = Show-RTLMessageBox "הורדה עדיין פעילה. לעצור אותה ולצאת?`n(אפשר יהיה להמשיך את ההורדה בפעם הבאה)" 'יציאה' 'YesNo' 'Question'
        if ($r -ne 'Yes') { $e.Cancel = $true; return }
        Stop-Download $script:Active.Kind
        Start-Sleep -Milliseconds 700
        try { Stop-Job -Job $script:Active.Job -ErrorAction SilentlyContinue; Remove-Job -Job $script:Active.Job -Force -ErrorAction SilentlyContinue } catch { }
    }
    $timer.Stop()
})

$script:UiScale = Set-UiScale $form 720 905 0.9 (New-Object Drawing.Size(705, 868))
if (-not $env:OTZ_UPDATER_NOUI) { [void]$form.ShowDialog() }
