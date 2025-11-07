Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
try { [System.Windows.Forms.Application]::SetHighDpiMode([System.Windows.Forms.HighDpiMode]::PerMonitorV2) | Out-Null } catch {}
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch {}

if (-not ('NewAssetTool.NativeMethods.Dpi' -as [Type])) {
  try {
    Add-Type -Namespace NewAssetTool.NativeMethods -Name Dpi -MemberDefinition @"
using System;
using System.Runtime.InteropServices;

public static class Dpi
{
  [DllImport("user32.dll", SetLastError = true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool SetProcessDpiAwarenessContext(IntPtr value);

  [DllImport("user32.dll")]
  public static extern IntPtr GetThreadDpiAwarenessContext();

  [DllImport("user32.dll")]
  public static extern int GetAwarenessFromDpiAwarenessContext(IntPtr value);
}
"@ -ErrorAction Stop
  } catch {}
}

$script:NewAssetToolPerMonitorDpiContextEnabled = $false
$global:NewAssetToolPerMonitorDpiContextEnabled = $false
try {
  $perMonitorV2Context = [System.IntPtr]-4
  if ('NewAssetTool.NativeMethods.Dpi' -as [Type]) {
    if ([NewAssetTool.NativeMethods.Dpi]::SetProcessDpiAwarenessContext($perMonitorV2Context)) {
      $script:NewAssetToolPerMonitorDpiContextEnabled = $true
      $global:NewAssetToolPerMonitorDpiContextEnabled = $true
      Write-Verbose "[DPI] SetProcessDpiAwarenessContext to PerMonitorV2 succeeded." -Verbose
    } else {
      $lastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
      Write-Verbose "[DPI] SetProcessDpiAwarenessContext returned false (LastError=$lastError). Using default DPI context." -Verbose
    }
  }
} catch {
  Write-Verbose "[DPI] Failed to set PerMonitorV2 awareness: $($_.Exception.Message)" -Verbose
}

function Get-NewAssetToolDpiContextDescription {
  try {
    if (-not ('NewAssetTool.NativeMethods.Dpi' -as [Type])) { return 'unknown' }
    $context = [NewAssetTool.NativeMethods.Dpi]::GetThreadDpiAwarenessContext()
    $awareness = [NewAssetTool.NativeMethods.Dpi]::GetAwarenessFromDpiAwarenessContext($context)
    switch ($awareness) {
      0 { return 'DPI_AWARENESS_UNAWARE' }
      1 { return 'DPI_AWARENESS_SYSTEM_AWARE' }
      2 { return 'DPI_AWARENESS_PER_MONITOR_AWARE' }
      3 { return 'DPI_AWARENESS_PER_MONITOR_AWARE_V2' }
      4 { return 'DPI_AWARENESS_UNAWARE_GDI_SCALED' }
      default { return "DPI_AWARENESS_UNKNOWN($awareness)" }
    }
  } catch {
    return 'unknown'
  }
}

$script:ThemeFontName = 'Segoe UI'
$script:ThemeFontSize = 10
$script:ThemeFont = New-Object System.Drawing.Font($script:ThemeFontName, $script:ThemeFontSize)
$script:ThemeFontSemibold = New-Object System.Drawing.Font('Segoe UI Semibold', $script:ThemeFontSize)
$script:ThemeColors = @{
  Background = [System.Drawing.Color]::FromArgb(248, 249, 251)
  Surface    = [System.Drawing.Color]::FromArgb(255, 255, 255)
  Header     = [System.Drawing.Color]::FromArgb(246, 247, 249)
  Text       = [System.Drawing.Color]::FromArgb(32, 32, 32)
  MutedText  = [System.Drawing.Color]::FromArgb(90, 90, 90)
  Accent     = [System.Drawing.Color]::FromArgb(0, 120, 212)
  AccentHover= [System.Drawing.Color]::FromArgb(0, 102, 189)
  Grid       = [System.Drawing.Color]::FromArgb(230, 234, 238)
  AltRow     = [System.Drawing.Color]::FromArgb(250, 252, 255)
  Selection  = [System.Drawing.Color]::FromArgb(229, 241, 251)
}
$script:NewAssetToolSearchTextBox = $null
function Set-ScanSearchControl {
  param([object]$control)
  $script:NewAssetToolSearchTextBox = $control
}

$script:DeviceTypeSummaryControl = $null
$script:SearchTextButtonStates = @{}
$script:RoundNowDisabledTooltip = 'A detected device type is required to use this feature.'

function Get-CurrentSearchInputText {
  try {
    if ($script:NewAssetToolSearchTextBox) {
      return '' + $script:NewAssetToolSearchTextBox.Text
    }
  } catch {}
  try {
    if ($txtScan) {
      return '' + $txtScan.Text
    }
  } catch {}
  return ''
}

function Get-CurrentDeviceTypeText {
  try {
    if ($script:DeviceTypeSummaryControl) {
      return '' + $script:DeviceTypeSummaryControl.Text
    }
  } catch {}
  try {
    if ($txtType) {
      return '' + $txtType.Text
    }
  } catch {}
  return ''
}

function Update-SearchDependentButtonStates {
  $deviceTypeText = Get-CurrentDeviceTypeText
  $hasDeviceType = -not [string]::IsNullOrWhiteSpace($deviceTypeText)

  if ($btnRoundNow) {
    try {
      $shouldEnableRoundNow = -not $hasDeviceType
      if ($btnRoundNow.Enabled -ne $shouldEnableRoundNow) {
        $btnRoundNow.Enabled = $shouldEnableRoundNow
      }
      if ($tip) {
        $tooltipText = if ($shouldEnableRoundNow) { '' } else { $script:RoundNowDisabledTooltip }
        $tip.SetToolTip($btnRoundNow, $tooltipText)
      }
    } catch {}
  }

  foreach ($entry in $script:SearchTextButtonStates.Values) {
    $button = $entry.Button
    if (-not $button) { continue }
    try {
      $baseEnabled = [bool]$entry.BaseEnabled
      $shouldEnable = $hasDeviceType -and $baseEnabled
      if ($button.Enabled -ne $shouldEnable) {
        $button.Enabled = $shouldEnable
      }
    } catch {}
  }
}

function Update-RoundNowButtonState {
  Update-SearchDependentButtonStates
}

function Set-SearchTextButtonBaseState {
  param(
    [System.Windows.Forms.Control]$Button,
    [bool]$BaseEnabled
  )

  if (-not $Button) { return }
  $key = $null
  try { $key = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($Button) } catch {}
  if ($null -eq $key) {
    try { $key = $Button.GetHashCode() } catch { return }
  }
  $script:SearchTextButtonStates[$key] = [PSCustomObject]@{
    Button = $Button
    BaseEnabled = [bool]$BaseEnabled
  }
  try { Update-RoundNowButtonState } catch {}
}
# ===== Script directory resolver (robust, PS 5.1-safe) =====
function Get-OwnScriptDir {
  try {
    if ($PSScriptRoot -and $PSScriptRoot -ne '') { return $PSScriptRoot }
  } catch {}
  try {
    if ($PSCommandPath -and $PSCommandPath -ne '') { return (Split-Path -Parent $PSCommandPath) }
  } catch {}
  try {
    if ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) {
      return (Split-Path -Parent $MyInvocation.MyCommand.Path)
    }
  } catch {}
  try {
    if ($env:__ScriptDir -and (Test-Path $env:__ScriptDir)) { return $env:__ScriptDir }
  } catch {}
  return (Get-Location).Path
}

# =================== Modern WinForms Theming Kit (PowerShell) ===================

# Ensure core assemblies are available
try {
  Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction SilentlyContinue
} catch {}

# --- High quality rounded button class ---
if (-not ('ModernUI.RoundedButton' -as [Type])) {
  try {
    Add-Type -ReferencedAssemblies System.Windows.Forms,System.Drawing -TypeDefinition @"
namespace ModernUI
{
  using System;
  using System.Drawing;
  using System.Drawing.Drawing2D;
  using System.Windows.Forms;

  public class RoundedButton : Button
  {
    public int CornerRadius { get; set; }

    private Color _baseBackColor;
    private Color _hoverBackColor;
    private bool _isHovering;
    private Color _enabledForeColor;
    private bool _hasEnabledForeColor;
    private bool _suppressForeColorTracking;
    private readonly Color _disabledForeColor = Color.FromArgb(160, 160, 160);

    public RoundedButton()
    {
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        UseVisualStyleBackColor = false;
        CornerRadius = 12;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);

        _baseBackColor = BackColor;
        _hoverBackColor = CalculateHoverColor(_baseBackColor);
        _enabledForeColor = ForeColor;
        _hasEnabledForeColor = true;
        UpdateEnabledStateVisuals();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

        Rectangle bounds = ClientRectangle;
        bounds.Width -= 1;
        bounds.Height -= 1;

        using (GraphicsPath path = CreatePath(bounds, CornerRadius))
        {
            Region oldRegion = Region;
            Region = new Region(path);
            if (oldRegion != null)
            {
                oldRegion.Dispose();
            }

            Color fill = Enabled ? BackColor : ControlPaint.Light(BackColor);
            using (SolidBrush brush = new SolidBrush(fill))
            {
                e.Graphics.FillPath(brush, path);
            }

            if (Focused && ShowFocusCues)
            {
                Rectangle focusBounds = bounds;
                focusBounds.Inflate(-4, -4);
                if (focusBounds.Width > 0 && focusBounds.Height > 0)
                {
                    using (GraphicsPath focusPath = CreatePath(focusBounds, Math.Max(0, CornerRadius - 3)))
                    using (Pen focusPen = new Pen(Color.FromArgb(200, Color.White)))
                    {
                        focusPen.DashStyle = DashStyle.Dot;
                        e.Graphics.DrawPath(focusPen, focusPath);
                    }
                }
            }

            TextFormatFlags format = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis;
            TextRenderer.DrawText(e.Graphics, Text, Font, bounds, ForeColor, format);
        }
    }

    protected override void OnPaintBackground(PaintEventArgs pevent)
    {
        // Prevent default background painting to avoid square artifacts.
    }

    protected override void OnBackColorChanged(EventArgs e)
    {
        if (!_isHovering)
        {
            UpdateBaseBackColor(BackColor);
        }
        else
        {
            _hoverBackColor = CalculateHoverColor(_baseBackColor);
        }

        base.OnBackColorChanged(e);
        Invalidate();
    }

    protected override void OnForeColorChanged(EventArgs e)
    {
        base.OnForeColorChanged(e);
        if (!_suppressForeColorTracking && Enabled)
        {
            _enabledForeColor = ForeColor;
            _hasEnabledForeColor = true;
        }
        Invalidate();
    }

    protected override void OnTextChanged(EventArgs e)
    {
        base.OnTextChanged(e);
        Invalidate();
    }

    protected override void OnEnabledChanged(EventArgs e)
    {
        base.OnEnabledChanged(e);
        UpdateEnabledStateVisuals();
        Invalidate();
    }

    protected override void OnMouseEnter(EventArgs e)
    {
        base.OnMouseEnter(e);
        if (!Enabled)
        {
            return;
        }

        _isHovering = true;
        Cursor = Cursors.Hand;
        BackColor = _hoverBackColor;
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        if (_isHovering)
        {
            _isHovering = false;
            BackColor = _baseBackColor;
        }

        base.OnMouseLeave(e);
    }

    private static Color CalculateHoverColor(Color color)
    {
        return ControlPaint.Light(color);
    }

    private void UpdateBaseBackColor(Color color)
    {
        _baseBackColor = color;
        _hoverBackColor = CalculateHoverColor(color);
    }

    private void UpdateEnabledStateVisuals()
    {
        _isHovering = false;
        BackColor = _baseBackColor;

        if (Enabled)
        {
            Cursor = Cursors.Hand;
            RestoreEnabledForeColor();
        }
        else
        {
            Cursor = Cursors.Default;
            ApplyDisabledForeColor();
        }
    }

    private void ApplyDisabledForeColor()
    {
        _suppressForeColorTracking = true;
        try
        {
            base.ForeColor = _disabledForeColor;
        }
        finally
        {
            _suppressForeColorTracking = false;
        }
    }

    private void RestoreEnabledForeColor()
    {
        if (!_hasEnabledForeColor)
        {
            _enabledForeColor = SystemColors.ControlText;
            _hasEnabledForeColor = true;
        }

        _suppressForeColorTracking = true;
        try
        {
            base.ForeColor = _enabledForeColor;
        }
        finally
        {
            _suppressForeColorTracking = false;
        }
    }

    private static GraphicsPath CreatePath(Rectangle rect, int radius)
    {
        GraphicsPath path = new GraphicsPath();

        if (radius <= 0)
        {
            path.AddRectangle(rect);
            return path;
        }

        int diameter = radius * 2;
        Rectangle arc = new Rectangle(rect.Location, new Size(diameter, diameter));
        path.AddArc(arc, 180, 90);

        arc.X = rect.Right - diameter;
        path.AddArc(arc, 270, 90);

        arc.Y = rect.Bottom - diameter;
        path.AddArc(arc, 0, 90);

        arc.X = rect.Left;
        path.AddArc(arc, 90, 90);

        path.CloseFigure();
        return path;
    }
  }
}
"@
  } catch {}
}

# --- DWM dark title bar + Mica (Win10/11) ---
try {
Add-Type -Namespace Dwm -Name Api -MemberDefinition @"
using System;
using System.Runtime.InteropServices;

public static class Api {
  [DllImport("dwmapi.dll", PreserveSig=true)]
  public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@ -ErrorAction SilentlyContinue
} catch {}

function Enable-ModernWindowEffects {
  param([System.Windows.Forms.Form]$Form, [switch]$Mica)
  try {
    $hwnd = $Form.Handle
    foreach ($attr in 20,19) {
      try { $val = 1; [void][Dwm.Api]::DwmSetWindowAttribute($hwnd, $attr, [ref]$val, 4); break } catch {}
    }
    if ($Mica) {
      try { $micaAttr = 38; $micaVal = 2; [void][Dwm.Api]::DwmSetWindowAttribute($hwnd, $micaAttr, [ref]$micaVal, 4) } catch {}
    }
  } catch {}
}

# --- Rounded corners for any control (e.g., Button) ---
function Set-RoundedCorners {
  param([System.Windows.Forms.Control]$Control, [int]$Radius = 8)

  $applyRegion = ({
    param($sender, $eventArgs)
    try {
      if (-not $sender) { return }
      $r = New-Object System.Drawing.Rectangle(0, 0, $sender.Width, $sender.Height)
      $d = $Radius * 2
      $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
      $gp.AddArc($r.X, $r.Y, $d, $d, 180, 90)
      $gp.AddArc($r.Right - $d, $r.Y, $d, $d, 270, 90)
      $gp.AddArc($r.Right - $d, $r.Bottom - $d, $d, $d, 0, 90)
      $gp.AddArc($r.X, $r.Bottom - $d, $d, $d, 90, 90)
      $gp.CloseAllFigures()
      if ($sender.Region) { $sender.Region.Dispose() }
      $sender.Region = New-Object System.Drawing.Region($gp)
    } catch {}
  }).GetNewClosure()

  $Control.Add_Resize($applyRegion)
  $applyRegion.Invoke($Control, [System.EventArgs]::Empty)
}

# --- DataGridView modern style (dark, compact, anti-flicker) ---
function Style-DataGridView {
  param([System.Windows.Forms.DataGridView]$Dgv)
  try {
    $pi = $Dgv.GetType().GetProperty('DoubleBuffered', 'NonPublic,Instance')
    if ($pi) { $pi.SetValue($Dgv, $true, $null) }
  } catch {}
  $Dgv.BorderStyle = 'None'
  $Dgv.BackgroundColor = $script:ThemeColors.Surface
  $Dgv.EnableHeadersVisualStyles = $false
  $Dgv.GridColor = $script:ThemeColors.Grid
  $Dgv.RowHeadersVisible = $false
  $Dgv.AutoSizeColumnsMode = 'Fill'
  $Dgv.SelectionMode = 'FullRowSelect'
  $Dgv.MultiSelect = $false
  $Dgv.AllowUserToResizeRows = $false
  $Dgv.RowTemplate.Height = 28

  $fg  = $script:ThemeColors.Text
  $bg  = $script:ThemeColors.Surface
  $bg2 = $script:ThemeColors.AltRow
  $sel = $script:ThemeColors.Selection
  $header = $script:ThemeColors.Header

  $Dgv.DefaultCellStyle.BackColor   = $bg
  $Dgv.DefaultCellStyle.ForeColor   = $fg
  $Dgv.DefaultCellStyle.Font = $script:ThemeFont
  $Dgv.DefaultCellStyle.SelectionBackColor = $sel
  $Dgv.DefaultCellStyle.SelectionForeColor = $fg
  $Dgv.AlternatingRowsDefaultCellStyle.BackColor = $bg2

  $Dgv.ColumnHeadersDefaultCellStyle.BackColor = $header
  $Dgv.ColumnHeadersDefaultCellStyle.ForeColor = $fg
  $Dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = $header
  $Dgv.ColumnHeadersDefaultCellStyle.SelectionForeColor = $fg
  $Dgv.ColumnHeadersDefaultCellStyle.Font = $script:ThemeFontSemibold
  $Dgv.RowHeadersDefaultCellStyle.BackColor = $header
  $Dgv.RowHeadersDefaultCellStyle.ForeColor = $fg
  $Dgv.RowHeadersDefaultCellStyle.SelectionBackColor = $header
  $Dgv.RowHeadersDefaultCellStyle.SelectionForeColor = $fg
  foreach ($c in $Dgv.Columns) { $c.MinimumWidth = 60 }
}

# --- Recursively theme common controls (dark palette, flat buttons, fonts) ---
function Set-ModernTheme {
  param([System.Windows.Forms.Control]$Root)

  $font = $Root.Font
  if (-not $font) {
    $Root.Font = $script:ThemeFont
  } elseif ($font.Name -eq $script:ThemeFontName) {
    if ([Math]::Abs($font.Size - $script:ThemeFontSize) -gt 0.01) { $Root.Font = $script:ThemeFont }
  } elseif ($font.Name -eq 'Segoe UI Semibold') {
    if ([Math]::Abs($font.Size - $script:ThemeFontSize) -gt 0.01) { $Root.Font = $script:ThemeFontSemibold }
  } elseif ($font.Name -like 'Segoe UI*') {
    if ([Math]::Abs($font.Size - $script:ThemeFontSize) -gt 0.01) {
      try { $Root.Font = New-Object System.Drawing.Font($font.Name, $script:ThemeFontSize, $font.Style) } catch { $Root.Font = $script:ThemeFont }
    }
  } elseif ($font.Name -notlike 'Segoe MDL2 Assets') {
    $Root.Font = $script:ThemeFont
  }

  $bgForm    = $script:ThemeColors.Background
  $bgPane    = $script:ThemeColors.Surface
  $bgHeader  = $script:ThemeColors.Header
  $fgText    = $script:ThemeColors.Text
  $accent    = $script:ThemeColors.Accent
  $accentHv  = $script:ThemeColors.AccentHover

  if ($Root -is [System.Windows.Forms.Form]) { $Root.BackColor = $bgForm; $Root.ForeColor = $fgText }

  foreach ($ctl in $Root.Controls) {
    switch -Regex ($ctl.GetType().FullName) {
      'System\.Windows\.Forms\.StatusStrip|System\.Windows\.Forms\.ToolStrip|System\.Windows\.Forms\.MenuStrip|System\.Windows\.Forms\.ContextMenuStrip' {
        $ctl.BackColor = $bgHeader; $ctl.ForeColor = $fgText
      }
      'System\.Windows\.Forms\.(Panel|GroupBox|TableLayoutPanel|FlowLayoutPanel)' {
        $ctl.BackColor = $bgPane; $ctl.ForeColor = $fgText
      }
      'System\.Windows\.Forms\.TabControl' {
        $ctl.BackColor = $bgPane; $ctl.ForeColor = $fgText
      }
      'System\.Windows\.Forms\.TabPage' {
        $ctl.BackColor = $bgPane; $ctl.ForeColor = $fgText
      }
      'System\.Windows\.Forms\.Label' {
        try {
          if ($ctl.ForeColor.IsEmpty -or $ctl.ForeColor.ToArgb() -eq [System.Drawing.SystemColors]::ControlText.ToArgb()) {
            $ctl.ForeColor = $fgText
          }
        } catch { $ctl.ForeColor = $fgText }
      }
      'System\.Windows\.Forms\.Button|ModernUI\.RoundedButton' {
        try { $ctl.FlatStyle = 'Standard' } catch {}
        try { $ctl.UseVisualStyleBackColor = $true } catch {}
        $ctl.BackColor = [System.Drawing.SystemColors]::Control
        if ($ctl.Enabled) {
          $ctl.ForeColor = [System.Drawing.SystemColors]::ControlText
        } else {
          $ctl.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
        }
        if ($ctl -is [ModernUI.RoundedButton]) {
          $ctl.CornerRadius = 4
        }
      }
      'System\.Windows\.Forms\.TextBox' {
        $ctl.BorderStyle = 'FixedSingle'
        $ctl.BackColor = [System.Drawing.Color]::White
        $ctl.ForeColor = $fgText
      }
      'System\.Windows\.Forms\.ComboBox' {
        $ctl.FlatStyle = 'Standard'
        $ctl.BackColor = [System.Drawing.Color]::White
        $ctl.ForeColor = $fgText
        $tagValues = @()
        try {
          $tag = $ctl.Tag
          if ($null -ne $tag) {
            if ($tag -is [string]) {
              $tagValues = @($tag)
            } elseif ($tag -is [System.Collections.IEnumerable]) {
              foreach ($tagItem in $tag) { if ($null -ne $tagItem) { $tagValues += $tagItem } }
            } else {
              $tagValues = @($tag)
            }
          }
        } catch {}

        $allowsEditableInput = $false
        foreach ($tagValue in $tagValues) {
          if ($null -eq $tagValue) { continue }
          $tagText = $tagValue.ToString()
          if ($tagText -match '(?i)(AllowEditableInput|AllowDropDown|AllowTyping|AllowTextInput)') { $allowsEditableInput = $true; break }
        }

        $styleName = ''
        try { $styleName = $ctl.DropDownStyle.ToString() } catch {}
        $isEditableStyle = $styleName -in @('DropDown','Simple')

        if (-not $allowsEditableInput -and -not $isEditableStyle) {
          if (-not $ctl.DropDownStyle -or $styleName -ne 'DropDownList') { $ctl.DropDownStyle = 'DropDownList' }
        }
      }
      'System\.Windows\.Forms\.CheckBox|System\.Windows\.Forms\.RadioButton' {
        $ctl.ForeColor = $fgText; $ctl.BackColor = [System.Drawing.Color]::Transparent
      }
      'System\.Windows\.Forms\.DataGridView' { Style-DataGridView $ctl }
    }
    if ($ctl.HasChildren) { Set-ModernTheme $ctl }
  }
}

# --- Lightweight icons via MDL2 glyphs ---
function Set-IconText {
  param([System.Windows.Forms.Control]$Control, [int]$Codepoint)
  try {
    $Control.Font = New-Object System.Drawing.Font('Segoe MDL2 Assets', 12)
    $Control.Text = [char]$Codepoint
  } catch {}
}

# ---------- Integration wrapper ----------
function Apply-ModernThemeToForm {
  param([Parameter(Mandatory)][System.Windows.Forms.Form]$Form)

  function Get-AllControls { param([System.Windows.Forms.Control]$Root)
    $list = New-Object System.Collections.Generic.List[System.Windows.Forms.Control]
    $q = New-Object System.Collections.Queue
    $q.Enqueue($Root)
    while ($q.Count) {
      $p = $q.Dequeue()
      foreach ($c in $p.Controls) { [void]$list.Add($c); $q.Enqueue($c) }
    }
    return $list
  }

  function Enable-DoubleBuffer { param([System.Windows.Forms.Control]$Ctrl)
    try {
      $flags = [System.Windows.Forms.ControlStyles]::OptimizedDoubleBuffer -bor
               [System.Windows.Forms.ControlStyles]::AllPaintingInWmPaint -bor
               [System.Windows.Forms.ControlStyles]::UserPaint
      $Ctrl.GetType().InvokeMember('SetStyle','InvokeMethod,NonPublic,Instance',$null,$Ctrl,@($flags,$true)) | Out-Null
      $Ctrl.Update()
    } catch {}
  }

  Set-ModernTheme $Form
  Enable-DoubleBuffer $Form

  $all = Get-AllControls $Form
  foreach ($container in $all | Where-Object {
      $_ -is [System.Windows.Forms.TableLayoutPanel] -or
      $_ -is [System.Windows.Forms.FlowLayoutPanel]  -or
      $_ -is [System.Windows.Forms.Panel]
    }) { Enable-DoubleBuffer $container }

  foreach ($dgv in $all | Where-Object { $_ -is [System.Windows.Forms.DataGridView] }) { Style-DataGridView $dgv }

  $Form.Add_Shown({ Enable-ModernWindowEffects $Form -Mica })
}

# ================= End Modern WinForms Theming Kit =================
# Force Data/Output to be script-relative and exist (Output).
$__ownDir = Get-OwnScriptDir
$script:DataFolder   = Join-Path $__ownDir 'Data'
$script:OutputFolder = Join-Path $__ownDir 'Output'
if (-not (Test-Path $script:OutputFolder)) { New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null }
# ===== end resolver =====

# ------------------ Globals ------------------
$script:DataFolder   = $null
$script:OutputFolder = $null
$script:Computers = @()
$script:Monitors  = @()
$script:Mics      = @()
$script:Scanners  = @()

$script:Carts    = @()
$script:IndexByAsset = @{} 
$script:IndexBySerial = @{} 
$script:IndexByName = @{}
$script:ComputerByAsset = @{}
$script:ComputerByName = @{} 
$script:ChildrenByParent = @{}
$script:LocationRows = @()             # dynamic dropdown data (Computers + user adds)
$script:LocationMasterRows = @()       # authoritative master data for validation
$script:UserAddedLocationRows = @()
$script:RoundingByAssetTag = @{}
$script:CurrentDisplay = $null
$script:CurrentParent  = $null
$script:editing = $false
$script:NewAssetToolMainForm = $null
# Canonical column order for rounding event exports (includes Comments column)
$script:RoundingEventColumns = @(
  'Timestamp','AssetTag','Name','Serial','City','Location','Building','Floor','Room',
  'CheckStatus','RoundingMinutes','CableMgmtOK','CablingNeeded','LabelOK','CartOK','PeripheralsOK',
  'MaintenanceType','Department','RoundingUrl','Comments','Rounded'
)
# Tolerant header map + fast caches for Room validation
$script:LocCols = @{}
$script:RoomsNorm  = @()  # all normalized Room strings from LocationMaster*.csv
$script:RoomCodes  = @()  # extracted room codes (e.g., 4003, 2S)
# ------------------ Helpers ------------------
function Canonical-Asset([string]$raw){
  if([string]::IsNullOrWhiteSpace($raw)){ return $null }
  $s = $raw.Trim().ToUpper() -replace '\s',''
  if($s -match '^HSS-?(\d+)$'){ return ('HSS-{0}' -f $matches[1]) }
  if($s -match '^C-?0*(\d+)$'){ return ('C{0}' -f $matches[1]) }
  if($s -match '^CRT-?(.+)$'){  return ('CRT-{0}' -f $matches[1]) }
  return $s
}
function HostnameKeyVariants([string]$raw){
  $out = New-Object System.Collections.ArrayList
  if([string]::IsNullOrWhiteSpace($raw)){ return $out }
  $u = $raw.Trim().ToUpper()
  [void]$out.Add($u)
  $compact = ($u -replace '[-\s]','')
  if($compact -and -not $out.Contains($compact)){ [void]$out.Add($compact) }
  if($u -match '^(AO)-?(.+)$'){
    [void]$out.Add(('{0}-{1}' -f $matches[1],$matches[2]))
    [void]$out.Add(('{0}{1}'  -f $matches[1],$matches[2]))
  }
  return $out
}
function Get-AssetKeyVariants([string]$raw){
  $variants = New-Object System.Collections.ArrayList
  if([string]::IsNullOrWhiteSpace($raw)){ return $variants }
  $key = $raw.Trim().ToUpper()
  if([string]::IsNullOrWhiteSpace($key)){ return $variants }
  if(-not $variants.Contains($key)){ [void]$variants.Add($key) }
  $canon = Canonical-Asset $key
  if($canon -and -not $variants.Contains($canon)){ [void]$variants.Add($canon) }
  $compact = ($key -replace '[-\s]','')
  if($compact -and -not $variants.Contains($compact)){ [void]$variants.Add($compact) }
  return $variants
}
function Normalize-Scan([string]$raw){
  if([string]::IsNullOrWhiteSpace($raw)){return $null}
  $s=$raw.Trim().ToUpper()
  $s=$s -replace '^(HOST\s*NAME|HOSTNAME)\s*[:#]?\s*',''
  $s=$s -replace '^(SN#?|S/N|SERIAL)\s*[:#]?\s*',''
  $s=$s -replace '^(ASSET\s*#?|ASSET#)\s*[:#]?\s*',''
  $s=$s -replace '\s',''
  if($s -match '^HSS[- ]?(\d+)$'){return @{Value=("HSS-{0}" -f $matches[1]);Kind='AssetTag'}}
  if($s -match '^C[- ]?0*(\d+)$'){return @{Value=("C{0}" -f $matches[1]);Kind='AssetTag'}}
  if($s -match '^(CRT[- ]?.+)$'){ return @{Value=($s -replace '^CRT[- ]?','CRT-');Kind='AssetTag'}}
  if($s -match '^(PC\d+)$'){     return @{Value=$matches[1];Kind='Hostname'}}
  if($s -match '^(LD-?\d+)$'){   return @{Value=($s -replace '^LD','LD-');Kind='Hostname'}}   # fixed
  if($s -match '^(TD-?\d+)$'){   return @{Value=($s -replace '^TD','TD-');Kind='Hostname'}}   # fixed
  if($s -match '^(AO[-]?\w+)$'){ return @{Value=($s -replace '^AO','AO-');Kind='Hostname'}}
  if($s -match '^[A-Z0-9\-]{5,}$'){return @{Value=$s;Kind='Serial'}}
  return @{Value=$s;Kind='Unknown'}
}
function Extract-RITM([string]$po){
  if([string]::IsNullOrWhiteSpace($po)){ return "" }
  $trimmed = $po.Trim()
  $m = [regex]::Match($trimmed, '(RITM\d+)')
  if($m.Success){ return $m.Groups[1].Value }
  $trpMatch = [regex]::Match($trimmed, 'TRP(?<date>\d{8})')
  if($trpMatch.Success){
    $dateDigits = $trpMatch.Groups['date'].Value
    try {
      $dt = [datetime]::ParseExact($dateDigits, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture)
      return "TRP - " + $dt.ToString('dd MMM yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {}
  }
  $trpMatchSix = [regex]::Match($trimmed, 'TRP(?<date>\d{6})(?!\d)')
  if($trpMatchSix.Success){
    $dateDigits = $trpMatchSix.Groups['date'].Value
    try {
      $month = [int]$dateDigits.Substring(0,2)
      $day = [int]$dateDigits.Substring(2,2)
      $year = 2000 + [int]$dateDigits.Substring(4,2)
      $dt = New-Object System.DateTime($year, $month, $day)
      return "TRP - " + $dt.ToString('dd MMM yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {}
  }
  return $trimmed
}
function Parse-DateLoose([string]$s){
  if([string]::IsNullOrWhiteSpace($s)){ return $null }
  $fmts = @('yyyy-MM-dd','yyyy/MM/dd','MM/dd/yyyy','MM-dd-yyyy','dd/MM/yyyy','dd-MM-yyyy','d/M/yyyy','M/d/yyyy')
  foreach($f in $fmts){ try{ return [datetime]::ParseExact($s,$f,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::AssumeLocal) } catch {} }
  try { return (Get-Date -Date $s) } catch { return $null }
}
function Fmt-DateLong($dt){ if(-not $dt){ return '' } try { return ([datetime]$dt).ToString('dd MMMM yyyy') } catch { return '' } }
function Get-RoundingStatus([Nullable[DateTime]]$dt){
  if(-not $dt){return 'Red'}
  $today=(Get-Date).Date; $dow=[int](Get-Date $today -UFormat %u)
  $monday=$today.AddDays(-($dow-1))
  if($dt -ge $monday){'Green'} elseif($dt -ge $today.AddDays(-35)){'Yellow'} else{'Red'}
}
# ---- Safe floor sorting ----
function Sort-Floors {
  param([object[]]$Values)
  if(-not $Values){ return @() }
  $pairs = @()
  foreach($v in $Values){
    $s = [string]$v
    $n = $s.Trim().ToUpper()
    $group = 2; $rank = 0
    if($n -in @('G','GRD','GROUND')){ $group = 0; $rank = 0 }
    elseif($n -match '^-?\d+$'){ $group = 1; $rank = [int]$n }
    elseif($n -match '^(B\d+)$'){ $group = 0; try{$rank = -1 * [int]($n.Substring(1))}catch{$rank = -1} }
    else { $group = 2; $rank = 0 }
    $pairs += [pscustomobject]@{ Orig=$s; G=$group; R=$rank; S=$n }
  }
  return ($pairs | Sort-Object G,R,S | Select-Object -ExpandProperty Orig)
}
function Register-SerialIndex([string]$serial,$record){
  if([string]::IsNullOrWhiteSpace($serial) -or -not $record){ return }
  $serialKey = $serial.Trim().ToUpper()
  if([string]::IsNullOrWhiteSpace($serialKey)){ return }
  $script:IndexBySerial[$serialKey] = $record
  $compactSerial = ($serialKey -replace '[-\s]','')
  if($compactSerial -and ($compactSerial -ne $serialKey)){
    $script:IndexBySerial[$compactSerial] = $record
  }
}
function Build-Indices {
  $script:IndexByAsset.Clear(); $script:IndexBySerial.Clear(); $script:IndexByName.Clear()
  $script:ComputerByAsset.Clear(); $script:ComputerByName.Clear()
  foreach($rec in $script:Computers){
    if($rec.asset_tag){
      foreach($variant in (Get-AssetKeyVariants $rec.asset_tag)){
        $script:IndexByAsset[$variant] = $rec
        $script:ComputerByAsset[$variant] = $rec
      }
    }
    if($rec.serial_number){ Register-SerialIndex $rec.serial_number $rec }
    if($rec.name){
      foreach($k in (HostnameKeyVariants $rec.name)){
        $script:IndexByName[$k] = $rec
        $script:ComputerByName[$k] = $rec
      }
    }
  }
  foreach($tbl in @('Monitors','Mics','Scanners','Carts')){
    $collection = (Get-Variable -Scope Script -Name $tbl -ErrorAction SilentlyContinue).Value
    if(-not $collection){ continue }
    foreach($rec in $collection){
      if($rec.asset_tag){
        foreach($variant in (Get-AssetKeyVariants $rec.asset_tag)){
          $script:IndexByAsset[$variant] = $rec
        }
      }
      if($rec.serial_number){ Register-SerialIndex $rec.serial_number $rec }
      if($rec.name){
        foreach($k in (HostnameKeyVariants $rec.name)){ $script:IndexByName[$k] = $rec }
      }
    }
  }
  $script:ChildrenByParent.Clear()
  foreach($tbl in @('Monitors','Mics','Scanners','Carts')){
    $collection = (Get-Variable -Scope Script -Name $tbl -ErrorAction SilentlyContinue).Value
    if(-not $collection){ continue }
    foreach($rec in $collection){
      $par = $rec.u_parent_asset
      if(-not [string]::IsNullOrWhiteSpace($par)){
        $key = (Canonical-Asset $par); if(-not $key){ $key = $par }
        if(-not $script:ChildrenByParent.ContainsKey($key)){
          $script:ChildrenByParent[$key] = New-Object System.Collections.ArrayList
        }
        [void]$script:ChildrenByParent[$key].Add($rec)
      }
    }
  }
}
function Find-RecordRaw([string]$q){
  $n=Normalize-Scan $q; if(-not $n){ return $null }
  $key=$n.Value.ToUpper()
  if($script:IndexByAsset.ContainsKey($key)){ return $script:IndexByAsset[$key] }
  elseif($script:IndexBySerial.ContainsKey($key)){ return $script:IndexBySerial[$key] }
  elseif($script:IndexByName.ContainsKey($key)){ return $script:IndexByName[$key] }
  return $null
}
function Resolve-ParentComputer($rec){
  if(-not $rec){ return $null }
  if($rec.Type -eq 'Computer'){ return $rec }
  if($rec.PSObject.Properties['u_parent_asset'] -and $rec.u_parent_asset){
    $upa = $rec.u_parent_asset.Trim().ToUpper()
    $cat = Canonical-Asset $upa
    if($cat -and $script:ComputerByAsset.ContainsKey($cat.ToUpper())){ return $script:ComputerByAsset[$cat.ToUpper()] }
    foreach($k in (HostnameKeyVariants $upa)){ if($script:ComputerByName.ContainsKey($k)){ return $script:ComputerByName[$k] } }
    if($script:IndexBySerial.ContainsKey($upa)){
      $cand=$script:IndexBySerial[$upa]; if($cand -and $cand.Type -eq 'Computer'){ return $cand }
    }
  }
  if($rec.name){
    $nmU = $rec.name.ToUpper()
    $base = ($nmU -replace '-MIC$','' -replace '-SCN$','')
    foreach($k in (HostnameKeyVariants $nmU)){ if($script:ComputerByName.ContainsKey($k)){ return $script:ComputerByName[$k] } }
    foreach($k in (HostnameKeyVariants $base)){ if($script:ComputerByName.ContainsKey($k)){ return $script:ComputerByName[$k] } }
  }
  if($rec.PSObject.Properties['RITM'] -and $rec.RITM){
    $cands = $script:Computers | Where-Object { $_.RITM -eq $rec.RITM }
    if($cands.Count -eq 1){ return $cands[0] }
  }
  return $null
}
function Get-ChildrenForParent($parentRec){
  $relationships = New-Object System.Collections.ArrayList
  if(-not $parentRec){ return $relationships }
  $parATKey = (Canonical-Asset $parentRec.asset_tag)
  if([string]::IsNullOrWhiteSpace($parATKey)){ return $relationships }

  $direct = New-Object System.Collections.ArrayList
  $addDirect = {
    param($item)
    if($item -and -not $direct.Contains($item)){ [void]$direct.Add($item) }
  }

  if($script:ChildrenByParent.ContainsKey($parATKey)){
    foreach($ch in $script:ChildrenByParent[$parATKey]){ & $addDirect $ch }
  }

  foreach($tbl in @('Monitors','Mics','Scanners','Carts')){
    $collection = (Get-Variable -Scope Script -Name $tbl -ErrorAction SilentlyContinue).Value
    if(-not $collection){ continue }
    foreach($rec in $collection){
      if(-not $rec.u_parent_asset){ continue }
      $upa = $rec.u_parent_asset.Trim()
      $matchHost = $false
      foreach($k in (HostnameKeyVariants $upa)){
        if($parentRec.name -and ($parentRec.name.Trim().ToUpper() -eq $k)){ $matchHost = $true; break }
      }
      if(-not $matchHost -and $parentRec.serial_number){
        if(($parentRec.serial_number.Trim().ToUpper()) -eq ($upa.Trim().ToUpper())){ $matchHost = $true }
      }
      if($matchHost){ & $addDirect $rec }
    }
  }

  $getObjKey = {
    param($obj)
    if(-not $obj){ return $null }
    try { return [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($obj) } catch {}
    try { return $obj.GetHashCode() } catch { return $null }
  }

  $peripheralAssignments = @{}
  $peripheralSet = New-Object System.Collections.ArrayList
  $addPeripheral = {
    param($peripheral,$cart)
    if(-not $peripheral){ return }
    if(($peripheral.Type -ne 'Mic') -and ($peripheral.Type -ne 'Scanner')){ return }
    if(-not $peripheralSet.Contains($peripheral)){ [void]$peripheralSet.Add($peripheral) }
    if($cart){
      $key = & $getObjKey $peripheral
      if($key -ne $null){ $peripheralAssignments[$key] = $cart }
    }
  }

  $carts = @($direct | Where-Object { $_.Type -eq 'Cart' })
  foreach($cart in $carts){
    $cartKey = Canonical-Asset $cart.asset_tag
    if($cartKey -and $script:ChildrenByParent.ContainsKey($cartKey)){
      foreach($gch in $script:ChildrenByParent[$cartKey]){
        if($gch -and ($gch.Type -eq 'Mic' -or $gch.Type -eq 'Scanner')){ & $addPeripheral $gch $cart }
      }
    }
    foreach($tbl in @('Mics','Scanners')){
      $collection = (Get-Variable -Scope Script -Name $tbl -ErrorAction SilentlyContinue).Value
      if(-not $collection){ continue }
      foreach($rec in $collection){
        if(-not $rec.u_parent_asset){ continue }
        $upa = $rec.u_parent_asset.Trim().ToUpper()
        $cartNameU = if($cart.name){ $cart.name.Trim().ToUpper() } else { '' }
        if($cartNameU -and $upa -eq $cartNameU){ & $addPeripheral $rec $cart }
      }
    }
  }

  foreach($rec in $direct){
    if(($rec.Type -eq 'Mic') -or ($rec.Type -eq 'Scanner')){
      $key = & $getObjKey $rec
      if($key -eq $null -or -not $peripheralAssignments.ContainsKey($key)){
        $matchedCart = $null
        if($rec.u_parent_asset){
          $token = $rec.u_parent_asset.Trim()
          $canon = Canonical-Asset $token
          foreach($cart in $carts){
            if($canon){
              $cartCanon = Canonical-Asset $cart.asset_tag
              if($cartCanon -and ($cartCanon.Trim().ToUpper() -eq $canon.Trim().ToUpper())){ $matchedCart = $cart; break }
            }
            if(-not $matchedCart -and $cart.name){
              foreach($variant in (HostnameKeyVariants $cart.name)){
                if($variant -and ($variant.Trim().ToUpper() -eq $token.Trim().ToUpper())){ $matchedCart = $cart; break }
              }
              if($matchedCart){ break }
            }
          }
        }
        if($matchedCart){ & $addPeripheral $rec $matchedCart } else { & $addPeripheral $rec $null }
      }
    }
  }

  $isTangentParent = ((Get-DetectedType $parentRec) -eq 'Tangent')

  $addRelation = {
    param($child,$immediateParent,$role)
    if(-not $child -or -not $immediateParent){ return }
    $rel = [pscustomobject]@{
      Record = $child
      Parent = $immediateParent
      Role   = $role
    }
    [void]$relationships.Add($rel)
  }

  if(-not $isTangentParent){
    foreach($rec in $direct){ & $addRelation $rec $parentRec 'Child' }
    foreach($peripheral in $peripheralSet){
      if(-not $direct.Contains($peripheral)){ & $addRelation $peripheral $parentRec 'Child' }
    }
  } else {
    foreach($cart in $carts){ & $addRelation $cart $parentRec 'Child' }
    foreach($rec in $direct){
      if($rec.Type -eq 'Cart'){ continue }
      $key = & $getObjKey $rec
      if($key -ne $null -and $peripheralAssignments.ContainsKey($key)){ continue }
      & $addRelation $rec $parentRec 'Child'
    }
    foreach($peripheral in $peripheralSet){
      $key = & $getObjKey $peripheral
      $cartParent = $null
      if($key -ne $null -and $peripheralAssignments.ContainsKey($key)){ $cartParent = $peripheralAssignments[$key] }
      if($cartParent){
        & $addRelation $peripheral $cartParent 'Grandchild'
      } else {
        & $addRelation $peripheral $parentRec 'Child'
      }
    }
  }

  return $relationships
}
function Compute-ProposedName($rec,$parent){
  if(-not $rec -or -not $parent){ return $null }
  if($rec.Type -eq 'Monitor'){ return $parent.name }
  elseif($rec.Type -eq 'Mic'){ return ($parent.name + "-Mic") }
  elseif($rec.Type -eq 'Scanner'){ return ($parent.name + "-SCN") }
  elseif($rec.Type -eq 'Cart'){ return ($parent.name + "-CRT") } else { return $null }
}
function Get-DetectedType($rec){
  if(-not $rec){ return '' }
  if($rec.Type -eq 'Monitor' -or $rec.Kind -eq 'Monitor'){ return 'Monitor' }
  if($rec.Type -eq 'Mic' -or $rec.Kind -eq 'Mic'){ return 'Microphone' }
  if($rec.Type -eq 'Scanner' -or $rec.Kind -eq 'Scanner'){ return 'Scanner' }
  if($rec.Type -eq 'Computer' -or $rec.Kind -eq 'Computer'){
    if($rec.name -match '^(?i)WT'){ return 'Thin Client' }
    if($rec.name -match '^(?i)PC'){ return 'Desktop' }
    if($rec.name -match '^(?i)LD'){ return 'Laptop' }
    if($rec.name -match '^(?i)AO'){ return 'Tangent' }
    return 'Computer'
  }
  if($rec.Type -eq 'Cart' -or $rec.Kind -eq 'Cart'){ return 'Cart' }
  return $rec.Type
}
function Color-RoundCell([string]$s){
  if($s -eq 'Green'){ $txtRound.BackColor=[System.Drawing.Color]::PaleGreen; return }
  if($s -eq 'Yellow'){ $txtRound.BackColor=[System.Drawing.Color]::LightYellow; return }
  $txtRound.BackColor=[System.Drawing.Color]::MistyRose
}
function Show-RoundingStatus($parentPC){
  $dt = $null
  if ($parentPC) { $dt = $parentPC.LastRounded }
  if ($dt) {
    $dateText = Fmt-DateLong $dt
    $today = (Get-Date).Date
    $d = [int](($today - $dt.Date).TotalDays)
    if ($d -le 0) {
      $txtRound.Text = ($dateText + " - Today")
    } else {
      $plural = if ($d -eq 1) { '' } else { 's' }
      $txtRound.Text = ("{0} - {1} day{2} ago" -f $dateText, $d, $plural)
    }
  } else {
    $txtRound.Text = ''
  }
  Color-RoundCell (Get-RoundingStatus $dt)
}
function Match-ParentToken([string]$token,$pc){
  if([string]::IsNullOrWhiteSpace($token) -or -not $pc){ return $false }
  $t = $token.Trim().ToUpper()
  $cat = Canonical-Asset $t
  if($cat){ if($pc.asset_tag -and ($cat.ToUpper() -eq $pc.asset_tag.Trim().ToUpper())){ return $true } }
  foreach($k in (HostnameKeyVariants $t)){ if($pc.name -and ($pc.name.Trim().ToUpper() -eq $k)){ return $true } }
  if($pc.serial_number -and ($pc.serial_number.Trim().ToUpper() -eq $t)){ return $true }
  return $false
}
# NEW: helper used elsewhere in the script
function Match-Token-To-Record([string]$token,$rec){
  if([string]::IsNullOrWhiteSpace($token) -or -not $rec){ return $false }
  $t = $token.Trim().ToUpper()
  $cat = Canonical-Asset $t
  if($cat){ if($rec.asset_tag -and ($rec.asset_tag.Trim().ToUpper() -eq $cat.ToUpper())){ return $true } }
  foreach($k in (HostnameKeyVariants $t)){ if($rec.name -and ($rec.name.Trim().ToUpper() -eq $k)){ return $true } }
  if($rec.serial_number -and ($rec.serial_number.Trim().ToUpper() -eq $t)){ return $true }
  return $false
}
function Validate-ParentAndName($displayRec,$parentRec){
  if($displayRec -and $displayRec.Type -ne 'Computer'){
       $raw = $null
    if($displayRec.PSObject.Properties['u_parent_asset']){ $raw = $displayRec.u_parent_asset }
    if([string]::IsNullOrWhiteSpace($raw)){
            $txtParent.Text='(blank)'
      $txtParent.BackColor=[System.Drawing.Color]::MistyRose
      $tip.SetToolTip($txtParent,"u_parent_asset is blank.")
    } else {
      $txtParent.Text=$raw
      $ok = $false
      $msg = ""
 if($parentRec -and (Match-ParentToken $raw $parentRec)){
        $ok = $true
        $msg = "u_parent_asset matches the resolved parent."
      } else {
        if($parentRec -and ($displayRec.Type -eq 'Mic' -or $displayRec.Type -eq 'Scanner') -and ($parentRec.name -match '^(?i)AO')){
          $carts = Find-CartsForComputer $parentRec
          if($carts.Count -gt 0){
            foreach($ct in $carts){
              if(Match-Token-To-Record $raw $ct){ $ok = $true; $msg = "u_parent_asset matches the resolved cart '"+$ct.name+"'."; break }
            }
            if(-not $ok){ $msg = "u_parent_asset does not match resolved cart for this Tangent." }
          } else {
            $msg = "No Cart found for this Tangent; expected u_parent_asset to match the Tangent or its Cart."
          }
        } else {
          $msg = if($parentRec){ "u_parent_asset does not match resolved parent '" + $parentRec.name + "'." } else { "u_parent_asset could not be resolved to a known computer." }
        }
      }
      if($ok){ $txtParent.BackColor=[System.Drawing.Color]::PaleGreen } else { $txtParent.BackColor=[System.Drawing.Color]::MistyRose }
      $tip.SetToolTip($txtParent,$msg)
    }
  } else {
    $txtParent.Text='(n/a)'
    $txtParent.BackColor=[System.Drawing.Color]::White
    $tip.SetToolTip($txtParent,"")
  }
  if($displayRec -and $parentRec -and $displayRec.Type -ne 'Computer'){
    $expected = Compute-ProposedName $displayRec $parentRec
    if($expected -and $displayRec.name -and ($displayRec.name.Trim().ToUpper() -ne $expected.Trim().ToUpper())){
      $txtHost.BackColor=[System.Drawing.Color]::MistyRose
      $tip.SetToolTip($txtHost, "Expected name: " + $expected)
    } else {
      $txtHost.BackColor=[System.Drawing.Color]::White
      $tip.SetToolTip($txtHost,"")
    }
  } else {
    $txtHost.BackColor=[System.Drawing.Color]::White
    $tip.SetToolTip($txtHost,"")
  }
}
function Find-CartsForComputer($pc){
  $res = New-Object System.Collections.ArrayList
  if(-not $pc){ return $res }
  foreach($ct in $script:Carts){
    if([string]::IsNullOrWhiteSpace($ct.u_parent_asset)){ continue }
    if(Match-Token-To-Record $ct.u_parent_asset $pc){ [void]$res.Add($ct) }
  }
  return $res
}
# ---- Value & Header Normalization ----
function Normalize-Field([string]$s){
  if([string]::IsNullOrWhiteSpace($s)){ return '' }
  $t = [string]$s
  $t = $t -replace '[\u200B\u200C\u200D\uFEFF]', ''              # zero-width & BOM
  $t = $t -replace '[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]', ' '  # space variants
  $t = $t -replace '[\u2010-\u2015\u2212\uFE58\uFE63\uFF0D]', '-'        # dashes
  $t = $t -replace '[\u2018\u2019\uFF07]', "'"                            # apostrophes
  $t = $t -replace '[\u201C\u201D\uFF02]', '"'                            # quotes
  $t = $t.Trim()
  $t = $t -replace '\s+', ' '
  $t = $t -replace '\s*\(\s*',' ('
  $t = $t -replace '\s*\)\s*',')'
  return $t.ToUpperInvariant()
}
function Normalize-Header([string]$s){
  if([string]::IsNullOrWhiteSpace($s)){ return '' }
  $t = $s -replace '[\uFEFF]', ''
  $t = $t.Trim()
  $t = $t -replace '\s+', ' '
  return $t.ToUpperInvariant()
}

if(-not (Get-Variable -Scope Script -Name LastLocationSelections -ErrorAction SilentlyContinue)){
  $script:LastLocationSelections = @{
    City      = ''
    Location  = ''
    Building  = ''
    Floor     = ''
    Room      = ''
  }
}
function Normalize-LocationComparisonValue([string]$value){
  if([string]::IsNullOrWhiteSpace($value)){ return '' }
  return Normalize-Field $value
}
function Update-LastLocationSelections {
  param(
    [string]$City,
    [string]$Location,
    [string]$Building,
    [string]$Floor,
    [string]$Room
  )
  $script:LastLocationSelections.City     = Normalize-LocationComparisonValue $City
  $script:LastLocationSelections.Location = Normalize-LocationComparisonValue $Location
  $script:LastLocationSelections.Building = Normalize-LocationComparisonValue $Building
  $script:LastLocationSelections.Floor    = Normalize-LocationComparisonValue $Floor
  $script:LastLocationSelections.Room     = Normalize-LocationComparisonValue $Room
}
function Get-LocColName([string]$wanted){
  if(-not $script:LocCols){ $script:LocCols = @{} }
  if($script:LocCols.ContainsKey($wanted)){ return $script:LocCols[$wanted] }
  $wantKey = Normalize-Header $wanted
  $actual = $null
  foreach($row in $script:LocationRows){
    foreach($p in $row.PSObject.Properties){
      if( (Normalize-Header $p.Name) -eq $wantKey ){ $actual = $p.Name; break }
    }
    if($actual){ break }
  }
  if(-not $actual){ $actual = $wanted }
  $script:LocCols[$wanted] = $actual
  return $actual
}
function Get-LocVal($row, [string]$wanted){
  if(-not $row){ return $null }
  $col = Get-LocColName $wanted
  $prop = $row.PSObject.Properties[$col]
  if($prop){ return $prop.Value }
  $bom = ([char]0xFEFF) + $wanted
  $prop2 = $row.PSObject.Properties[$bom]
  if($prop2){ return $prop2.Value }
  return $null
}
# ---- Room caches ----
function Extract-RoomCode([string]$s){
  if([string]::IsNullOrWhiteSpace($s)){ return '' }
  $t = Normalize-Field $s
  $m = [regex]::Match($t, '^[A-Z0-9]+')
  if($m.Success){ return $m.Value } else { return '' }
}
# ---- Department master/user-adds ----
function Get-DepartmentFile([string]$fileName){
  $path = $null
  try { if($script:DataFolder){ $path = Join-Path $script:DataFolder $fileName } } catch {}
  if(-not $path -or -not (Test-Path $path)){
    try { $path = Join-Path $PSScriptRoot (Join-Path 'Data' $fileName) } catch {}
  }
  if(-not $path -or -not (Test-Path $path)){
    try { $path = $fileName } catch {}
  }
  if($path -and (Test-Path $path)){ return $path }
  return $null
}
function Import-DepartmentRows([string]$path){
  $rows = @()
  if(-not $path){ return $rows }
  try {
    $lines = Get-Content -Path $path -Encoding UTF8
  } catch { return $rows }
  if(-not $lines){ return $rows }
  if(-not ($lines -is [System.Collections.IEnumerable])){ $lines = @($lines) }
  $cleanLines = @()
  foreach($line in $lines){
    if($null -eq $line){ continue }
    $txt = [string]$line
    if($txt){
      $trimmed = $txt.Trim()
      if($trimmed.Length -gt 0){ $cleanLines += $txt }
    }
  }
  if($cleanLines.Count -eq 0){ return $rows }
  $first = [string]$cleanLines[0]
  $firstNoQuotes = ($first -replace '^[\uFEFF\"\s]+','') -replace '[\"\s]+$',''
  $hasHeader = ((Normalize-Field $firstNoQuotes) -eq (Normalize-Field 'Department'))
  if(-not $hasHeader){ $cleanLines = @('Department') + $cleanLines }
  try {
    $rows = $cleanLines | ConvertFrom-Csv
  } catch {
    $fallback = New-Object System.Collections.Generic.List[object]
    foreach($line in $cleanLines){
      $txt = [string]$line
      if(-not $txt){ continue }
      $value = $txt.Trim()
      if((Normalize-Field $value) -eq 'DEPARTMENT'){ continue }
      if($value.StartsWith('"') -and $value.EndsWith('"')){ $value = $value.Substring(1, $value.Length-2) }
      $value = $value -replace '""','"'
      $fallback.Add([pscustomobject]@{ Department = $value }) | Out-Null
    }
    $rows = $fallback.ToArray()
  }
  return $rows
}
function Add-DepartmentValue([string]$dept){
  if([string]::IsNullOrWhiteSpace($dept)){ return }
  $display = ($dept -replace '[\uFEFF]', '').Trim()
  if([string]::IsNullOrWhiteSpace($display)){ return }
  $norm = Normalize-Field $display
  if(-not $norm){ return }
  if(-not $script:DepartmentListNorm.Contains($norm)){
    [void]$script:DepartmentList.Add($display)
    [void]$script:DepartmentListNorm.Add($norm)
  }
}
function Load-DepartmentMaster(){
  try{
    $script:DepartmentMaster = @()
    $script:DepartmentUserAdds = @()
    $script:DepartmentList = New-Object System.Collections.Generic.List[string]
    $script:DepartmentListNorm = New-Object System.Collections.Generic.HashSet[string]
    $fileD = Get-DepartmentFile 'DepartmentMaster.csv'
    if($fileD){
      $script:DepartmentMaster = Import-DepartmentRows $fileD
      foreach($row in $script:DepartmentMaster){
        if($row -and $row.PSObject.Properties['Department']){ Add-DepartmentValue $row.Department }
      }
    }
    $fileDU = Get-DepartmentFile 'DepartmentMaster-UserAdds.csv'
    if($fileDU){
      $script:DepartmentUserAdds = Import-DepartmentRows $fileDU
      foreach($row in $script:DepartmentUserAdds){
        if($row -and $row.PSObject.Properties['Department']){ Add-DepartmentValue $row.Department }
      }
    }
    $sorted = @($script:DepartmentList | Sort-Object -Unique)
    $script:DepartmentList = New-Object System.Collections.Generic.List[string]
    foreach($item in $sorted){ [void]$script:DepartmentList.Add($item) }
  } catch {}
}
function Save-DepartmentUserAdd([string]$dept){
  try{
    if([string]::IsNullOrWhiteSpace($dept)){ return }
    $n = Normalize-Field $dept
    if($script:DepartmentListNorm.Contains($n)){ return }
    $fileDU = $null
    try {
      if($script:DataFolder){ $fileDU = Join-Path $script:DataFolder 'DepartmentMaster-UserAdds.csv' }
    } catch {}
    if(-not $fileDU){
      try {
        $dataFolder = Join-Path $PSScriptRoot 'Data'
        if(-not (Test-Path $dataFolder)){
          try { New-Item -Path $dataFolder -ItemType Directory -Force | Out-Null } catch {}
        }
        $fileDU = Join-Path $dataFolder 'DepartmentMaster-UserAdds.csv'
      } catch {}
    }
    if(-not $fileDU){
      try { $fileDU = 'DepartmentMaster-UserAdds.csv' } catch {}
    }
    if($fileDU){
      try {
        $dir = Split-Path -Path $fileDU -Parent
        if($dir -and -not (Test-Path $dir)){
          try { New-Item -Path $dir -ItemType Directory -Force | Out-Null } catch {}
        }
      } catch {}
    }
    $exists = Test-Path $fileDU
    $row = [pscustomobject]@{ Department = $dept }
    if(-not $exists){ $row | Export-Csv -Path $fileDU -NoTypeInformation -Encoding UTF8 }
    else { $row | Export-Csv -Path $fileDU -NoTypeInformation -Append -Encoding UTF8 }
    Load-DepartmentMaster
  } catch { }
}
function Populate-Department-Combo([string]$current){
  try{
    if(-not $script:DepartmentList -or $script:DepartmentList.Count -eq 0){
      try { Load-DepartmentMaster } catch {}
    }
    $items = @()
    if($script:DepartmentList){
      $items = @($script:DepartmentList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    foreach($combo in @($cmbDept,$cmbDepartment,$ddlDept,$ddlDepartment)){
      if(-not $combo){ continue }
      try {
        $combo.BeginUpdate()
        $combo.Items.Clear()
        if($items.Count -gt 0){ $combo.Items.AddRange(@($items)) }
        if($null -ne $current){ $combo.Text = $current }
        $combo.AutoCompleteMode  = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
        $combo.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
      } finally {
        try { $combo.EndUpdate() } catch {}
      }
    }
    foreach($txt in @($txtDept,$txtDepartment)){
      if($txt -and $null -ne $current){ $txt.Text = $current }
    }
  } catch {}
}
function Rebuild-RoomCaches(){
  $rooms = New-Object System.Collections.Generic.List[string]
  $codes = New-Object System.Collections.Generic.List[string]
  foreach($row in $script:LocationMasterRows){
    if(-not $row){ continue }
    $raw = $null
    try {
      $raw = $row.Room
      if(-not $raw -and $row.PSObject -and $row.PSObject.Properties['Room']){
        $raw = $row.PSObject.Properties['Room'].Value
      }
    } catch {}
    if([string]::IsNullOrWhiteSpace($raw)){ continue }
    $n = Normalize-Field $raw
    if($n){ [void]$rooms.Add($n) }
    $c = Extract-RoomCode $raw
    if($c){ [void]$codes.Add($c) }
  }
  $script:RoomsNorm = ($rooms | Select-Object -Unique)
  $script:RoomCodes = ($codes | Select-Object -Unique)
}
# ------------------ Load & Save ------------------
function Load-LocationMaster($folder){
  $script:LocationMasterRows = @()
  $script:UserAddedLocationRows = @()
  $lm = Join-Path $folder 'LocationMaster.csv'
  if(Test-Path $lm){ $script:LocationMasterRows += Import-Csv $lm }
  $lm2 = Join-Path $folder 'LocationMaster-UserAdds.csv'
  if(Test-Path $lm2){ $script:UserAddedLocationRows += Import-Csv $lm2 }
  $script:LocCols = @{}
  Rebuild-RoomCaches
  Rebuild-LocationDropdownRows
}
function Save-LocationUserAdd([string]$city,[string]$loc,[string]$b,[string]$f,[string]$r){
  try{
    $file = $null
    try {
      if($script:DataFolder){ $file = Join-Path $script:DataFolder 'LocationMaster-UserAdds.csv' }
    } catch {}
    if(-not $file){
      try {
        $dataFolder = Join-Path $PSScriptRoot 'Data'
        if(-not (Test-Path $dataFolder)){
          try { New-Item -Path $dataFolder -ItemType Directory -Force | Out-Null } catch {}
        }
        $file = Join-Path $dataFolder 'LocationMaster-UserAdds.csv'
      } catch {}
    }
    if(-not $file){
      try { $file = 'LocationMaster-UserAdds.csv' } catch {}
    }
    if($file){
      try {
        $dir = Split-Path -Path $file -Parent
        if($dir -and -not (Test-Path $dir)){
          try { New-Item -Path $dir -ItemType Directory -Force | Out-Null } catch {}
        }
      } catch {}
      if(-not (Test-Path $file)){
        'City,Location,Building,Floor,Room' | Out-File -FilePath $file -Encoding UTF8
      }
      ('"{0}","{1}","{2}","{3}","{4}"' -f $city,$loc,$b,$f,$r) | Add-Content -Path $file -Encoding UTF8
    }
  } catch {}
  try {
    $newRow = [pscustomobject]@{ City=$city; Location=$loc; Building=$b; Floor=$f; Room=$r }
    $script:UserAddedLocationRows += $newRow
  } catch {}
  Rebuild-LocationDropdownRows
}

function Rebuild-LocationDropdownRows(){
  $unique = @{}
  $rows = New-Object 'System.Collections.Generic.List[object]'
  $addRow = {
    param($city,$location,$building,$floor,$room)
    $locVal = if($location){ ([string]$location).Trim() } else { '' }
    if([string]::IsNullOrWhiteSpace($locVal)){ return }
    $cityVal = if($city){ ([string]$city).Trim() } else { '' }
    $bldVal = if($building){ ([string]$building).Trim() } else { '' }
    $floorVal = if($floor){ ([string]$floor).Trim() } else { '' }
    $roomVal = if($room){ ([string]$room).Trim() } else { '' }
    $key = '{0}|{1}|{2}|{3}|{4}' -f `
      (Normalize-LocationComparisonValue $cityVal),
      (Normalize-LocationComparisonValue $locVal),
      (Normalize-LocationComparisonValue $bldVal),
      (Normalize-LocationComparisonValue $floorVal),
      (Normalize-LocationComparisonValue $roomVal)
    if(-not $unique.ContainsKey($key)){
      $unique[$key] = $true
      [void]$rows.Add([pscustomobject]@{
        City = $cityVal
        Location = $locVal
        Building = $bldVal
        Floor = $floorVal
        Room = $roomVal
      })
    }
  }

  foreach($row in $script:UserAddedLocationRows){
    if(-not $row){ continue }
    $cityVal = $null
    $locVal = $null
    $bldVal = $null
    $floorVal = $null
    $roomVal = $null
    try { $cityVal = $row.City } catch {}
    try { $locVal = $row.Location } catch {}
    try { $bldVal = $row.Building } catch {}
    try { $floorVal = $row.Floor } catch {}
    try { $roomVal = $row.Room } catch {}
    & $addRow $cityVal $locVal $bldVal $floorVal $roomVal
  }

  foreach($pc in $script:Computers){
    if(-not $pc){ continue }
    $cityVal = $null
    try {
      if($pc.PSObject.Properties['City']){ $cityVal = $pc.City }
      elseif($pc.PSObject.Properties['location.city']){ $cityVal = $pc.PSObject.Properties['location.city'].Value }
    } catch {}
    $locVal = $null
    $bldVal = $null
    $floorVal = $null
    $roomVal = $null
    try { $locVal = $pc.location } catch {}
    try { $bldVal = $pc.u_building } catch {}
    try { $floorVal = $pc.u_floor } catch {}
    try { $roomVal = $pc.u_room } catch {}
    & $addRow $cityVal $locVal $bldVal $floorVal $roomVal
  }

  try {
    $script:LocationRows = $rows.ToArray()
  } catch {
    $script:LocationRows = @()
  }
  $script:LocCols = @{}
}
function Load-RoundingMapping([string]$folder){
  $script:RoundingByAssetTag.Clear()
  $path = Join-Path $folder 'Rounding.csv'
  if(Test-Path $path){
    try{
      $rows = Import-Csv $path
      foreach($r in $rows){
        $at = $r.'Asset Tag'; $id = $r.SlNo
        if(-not [string]::IsNullOrWhiteSpace($at) -and -not [string]::IsNullOrWhiteSpace($id)){
          $script:RoundingByAssetTag[$at.Trim().ToUpper()] = $id.Trim()
        }
      }
    } catch {}
  }
}
function Load-DataFolder([string]$folder){
  $script:DataFolder = $folder
  if(-not $script:OutputFolder){ $script:OutputFolder = $folder }
  Load-LocationMaster $folder
  Load-RoundingMapping $folder
  try { Load-DepartmentMaster } catch {}
  $cfile   = Join-Path $folder 'Computers.csv'
  $mfile   = Join-Path $folder 'Monitors.csv'
  $micfile = Join-Path $folder 'Mics.csv'
  $sfile   = Join-Path $folder 'Scanners.csv'
  $script:Computers = @(); $script:Monitors = @(); $script:Mics = @(); $script:Scanners = @()
  
  $cartfile = Join-Path $folder 'Carts.csv'
  $script:Carts = @()
if(Test-Path $cfile){
    $raw = Import-Csv $cfile
    foreach($r in $raw){
      $obj = [pscustomobject]@{
      u_device_rounding = $r.u_device_rounding
      u_department_location = $r.u_department_location
        Kind='Computer'; Type='Computer'
        name=$r.name; asset_tag=$r.asset_tag; serial_number=$r.serial_number
        location=$r.location; u_building=$r.u_building; u_room=$r.u_room; u_floor=$r.u_floor
        po_number=$r.po_number; u_scheduled_retirement=$r.u_scheduled_retirement
        u_last_rounded_date=$r.u_last_rounded_date
        City=$(if($r.'location.city'){ $r.'location.city'.Trim() } else { '' })
      }
      $obj | Add-Member -NotePropertyName RITM -NotePropertyValue (Extract-RITM $obj.po_number) -Force
      $obj | Add-Member -NotePropertyName Retire -NotePropertyValue (Parse-DateLoose $obj.u_scheduled_retirement) -Force
      $obj | Add-Member -NotePropertyName LastRounded -NotePropertyValue (Parse-DateLoose $obj.u_last_rounded_date) -Force
      $script:Computers += $obj
    }
  }
  if(Test-Path $mfile){
    $raw = Import-Csv $mfile
    foreach($r in $raw){
      $obj = [pscustomobject]@{
      u_device_rounding = $r.u_device_rounding
      u_department_location = $r.u_department_location
        Kind='Peripheral'; Type='Monitor'
        u_parent_asset=$r.u_parent_asset
        name=$r.name; asset_tag=$r.asset_tag; serial_number=$r.serial_number
        location=$r.location; u_building=$r.u_building; u_room=$r.u_room; u_floor=$r.u_floor
        po_number=$r.po_number; u_scheduled_retirement=$r.u_scheduled_retirement
      }
      $obj | Add-Member -NotePropertyName RITM -NotePropertyValue (Extract-RITM $obj.po_number) -Force
      $obj | Add-Member -NotePropertyName Retire -NotePropertyValue (Parse-DateLoose $obj.u_scheduled_retirement) -Force
      $script:Monitors += $obj
    }
  }
  if(Test-Path $micfile){
    $raw = Import-Csv $micfile
    foreach($r in $raw){
      $obj = [pscustomobject]@{
      u_device_rounding = $r.u_device_rounding
      u_department_location = $r.u_department_location
        Kind='Peripheral'; Type='Mic'
        u_parent_asset=$r.u_parent_asset
        name=$r.name; asset_tag=$r.asset_tag; serial_number=$r.serial_number
        location=$r.location; u_building=$r.u_building; u_room=$r.u_room; u_floor=$r.u_floor
        po_number=$r.po_number; u_scheduled_retirement=$r.u_scheduled_retirement
      }
      $obj | Add-Member -NotePropertyName RITM -NotePropertyValue (Extract-RITM $obj.po_number) -Force
      $script:Mics += $obj
    }
  }
  if(Test-Path $sfile){
    $raw = Import-Csv $sfile
    foreach($r in $raw){
      $obj = [pscustomobject]@{
      u_device_rounding = $r.u_device_rounding
      u_department_location = $r.u_department_location
        Kind='Peripheral'; Type='Scanner'
        u_parent_asset=$r.u_parent_asset
        name=$r.name; asset_tag=$r.asset_tag; serial_number=$r.serial_number
        location=$r.location; u_building=$r.u_building; u_room=$r.u_room; u_floor=$r.u_floor
        po_number=$r.po_number; u_scheduled_retirement=$r.u_scheduled_retirement
      }
      $script:Scanners += $obj
    }
  }
  if(Test-Path $cartfile){
    $raw = Import-Csv $cartfile
    foreach($r in $raw){
      $asset = $null
      if($r.PSObject.Properties['asset'] -and $r.asset){ $asset = $r.asset }
      elseif($r.PSObject.Properties['asset_tag'] -and $r.asset_tag){ $asset = $r.asset_tag }
      $obj = [pscustomobject]@{
      u_device_rounding = $r.u_device_rounding
      u_department_location = $r.u_department_location
        Kind='Peripheral'; Type='Cart'
        u_parent_asset=$r.u_parent_asset
        name=$r.name
        asset_tag=$asset
        serial_number=$r.serial_number
        location=$r.location
        u_building=$null; u_room=$null; u_floor=$null
        po_number=$null; u_scheduled_retirement=$null
      }
      $script:Carts += $obj
    }
  }
  Rebuild-LocationDropdownRows
  Build-Indices
}
function Save-AllCSVs {
  $out = if($script:OutputFolder){$script:OutputFolder}else{$script:DataFolder}
  if(-not $out){
    [System.Windows.Forms.MessageBox]::Show("No output folder available.","Save") | Out-Null; return
  }
  if($script:Computers.Count -gt 0){
    $path = Join-Path $out 'Computers.csv'
    $script:Computers | Select-Object name,asset_tag,serial_number,location,u_building,u_room,u_floor,po_number,u_scheduled_retirement,u_last_rounded_date |
      Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
  }
  foreach($pair in @(@('Monitors',$script:Monitors), @('Mics',$script:Mics), @('Scanners',$script:Scanners), @('Carts',$script:Carts))){
    $name = $pair[0]; $rows=$pair[1]
    if($rows.Count -gt 0){
      $path = Join-Path $out ($name + '.csv')
      $rows | Select-Object u_parent_asset,name,asset_tag,serial_number,location,u_building,u_room,u_floor,po_number,u_scheduled_retirement |
        Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    }
  }
  [System.Windows.Forms.MessageBox]::Show("Saved CSVs to
$($out)","Save") | Out-Null
}
# ------------------ UI (TableLayout; DPI aware) ------------------
$LEFT_COL_PERCENT   = 46
$RIGHT_COL_PERCENT  = 54
$GAP                = 6
$form = New-Object System.Windows.Forms.Form
$script:NewAssetToolManualScaleFactor = 1.0
$desiredChromeScale = 0.8
$applyWinFormsManualScale = {
  param([string]$Source = 'unspecified')

  if (-not $global:NewAssetToolPerMonitorDpiContextEnabled) { return }

  $current = $script:NewAssetToolManualScaleFactor
  if ([Math]::Abs([double]$current) -lt [double]::Epsilon) { $current = 1.0 }
  $ratio = $desiredChromeScale / [double]$current

  $contextDescription = 'unknown'
  try {
    if (Get-Command Get-NewAssetToolDpiContextDescription -ErrorAction SilentlyContinue) {
      $contextDescription = Get-NewAssetToolDpiContextDescription
    }
  } catch {}

  if ([Math]::Abs($ratio - 1.0) -gt 0.0001) {
    Write-Verbose (
      "[DPI][WinForms] Scaling form by ratio {0:n3} from {1:n3} to {2:n3} ({3}) context={4}" -f 
      $ratio, $current, $desiredChromeScale, $Source, $contextDescription
    ) -Verbose
    try { $form.Scale([float]$ratio) } catch {}
  } else {
    Write-Verbose (
      "[DPI][WinForms] Scale already at target {0:n3} ({1}) context={2}" -f 
      $desiredChromeScale, $Source, $contextDescription
    ) -Verbose
  }

  $script:NewAssetToolManualScaleFactor = $desiredChromeScale
}
$form.Text = "Inventory Assoc Finder - OMI"
$statusPathLabelDefault = "Data: (not set)    |    Output: (not set)"
$form.StartPosition="CenterScreen"
$form.WindowState='Maximized'
$form.BackColor=[System.Drawing.Color]::White
$form.KeyPreview = $true
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
# ToolTip
$tip = New-Object System.Windows.Forms.ToolTip
$tip.AutoPopDelay = 8000
$tip.InitialDelay = 400
$tip.ReshowDelay  = 200
$tip.ShowAlways   = $true
# ---------- HEADER ----------
$panelTop = New-Object System.Windows.Forms.Panel
$panelTop.Dock = 'Top'
$panelTop.AutoSize = $true
$panelTop.AutoSizeMode = 'GrowAndShrink'
$panelTop.Padding = New-Object System.Windows.Forms.Padding($GAP, $GAP, $GAP, 0)
$panelTop.BackColor = $script:ThemeColors.Header
# Row 1: Paths + counters
$flpTop = New-Object System.Windows.Forms.FlowLayoutPanel
$flpTop.Dock = 'Top'
$flpTop.AutoSize = $true
$flpTop.AutoSizeMode = 'GrowAndShrink'
$flpTop.WrapContents = $false
$flpTop.FlowDirection = 'LeftToRight'
$flpTop.Margin = '0,0,0,0'
$flpTop.Padding = '0,0,0,0'
$lblDataPath = New-Object System.Windows.Forms.Label
$lblDataPath.Text = "Data: (not set)"
$lblDataPath.AutoSize = $true
$lblDataPath.Margin = '0,6,12,0'
$lblOutputPath = New-Object System.Windows.Forms.Label
$lblOutputPath.Text = "Output: (not set)"
$lblOutputPath.AutoSize = $true
$lblOutputPath.Margin = '0,6,12,0'
$lblDataStatus = New-Object System.Windows.Forms.Label
$lblDataStatus.Text = "Computers: 0 | Monitors: 0 | Mics: 0 | Scanners: 0 | Carts: 0 | Locations: 0"
$lblDataStatus.AutoSize = $true
$lblDataStatus.Margin = '0,6,0,0'
$flpTop.Visible = $false  # moved to status bar
# Row 2: Scan box logic (UI hosted in WPF shell)
$txtScan = New-Object System.Windows.Forms.TextBox
$txtScan.Visible = $false
# Maintain layout order even though scan UI is hosted externally
$panelTop.Controls.Add($flpTop)
# ---------- END HEADER ----------
# Main 2-col table
$LEFT_COL_WIDTH  = 520
$PANEL2_MIN_WIDTH = 400
$splitter = New-Object System.Windows.Forms.SplitContainer
$splitter.Dock = 'Fill'
$splitter.Orientation = [System.Windows.Forms.Orientation]::Vertical
$splitter.FixedPanel = [System.Windows.Forms.FixedPanel]::Panel1

function Set-SplitterMinimums {
  param(
    [System.Windows.Forms.SplitContainer]$target,
    [int]$panel1Desired,
    [int]$panel2Desired
  )

  $available = [Math]::Max(0, $target.Width - $target.Padding.Horizontal - $target.SplitterWidth)
  $panel1Min = [Math]::Min($panel1Desired, $available)
  $panel2Available = [Math]::Max(0, $available - $panel1Min)
  $panel2Min = [Math]::Min($panel2Desired, $panel2Available)

  $target.Panel1MinSize = $panel1Min
  $target.Panel2MinSize = $panel2Min
}

Set-SplitterMinimums -target $splitter -panel1Desired $LEFT_COL_WIDTH -panel2Desired $PANEL2_MIN_WIDTH
$splitter.Add_SizeChanged({ Set-SplitterMinimums -target $splitter -panel1Desired $LEFT_COL_WIDTH -panel2Desired $PANEL2_MIN_WIDTH })
$form.Add_Shown({ Set-SplitterMinimums -target $splitter -panel1Desired $LEFT_COL_WIDTH -panel2Desired $PANEL2_MIN_WIDTH })

if ($global:NewAssetToolPerMonitorDpiContextEnabled) {
  $form.Add_Shown({ & $applyWinFormsManualScale 'Form.Shown' })
  try {
    $form.Add_DpiChanged({
      param($sender, $eventArgs)
      $script:NewAssetToolManualScaleFactor = 1.0
      & $applyWinFormsManualScale 'Form.DpiChanged'
    })
  } catch {
    Write-Verbose "[DPI][WinForms] Failed to attach DpiChanged handler: $($_.Exception.Message)" -Verbose
  }
}

$splitter.SplitterDistance = $LEFT_COL_WIDTH
$splitter.SplitterWidth = 6
$splitter.IsSplitterFixed = $false
$splitter.Padding = New-Object System.Windows.Forms.Padding(16)
$splitter.BackColor = [System.Drawing.Color]::White
Set-SplitterMinimums -target $splitter -panel1Desired $LEFT_COL_WIDTH -panel2Desired $PANEL2_MIN_WIDTH
function New-L($t,$x,$y){$l=New-Object System.Windows.Forms.Label;$l.Text=$t;$l.AutoSize=$true;$l.Location=New-Object System.Drawing.Point($x,$y);$l}
function New-RO($x,$y,$w){$t=New-Object System.Windows.Forms.TextBox;$t.Location="$x,$y";$t.Size="$w,24";$t.ReadOnly=$true;$t.BackColor='White';$t}
# Left column stack
$tlpLeft = New-Object System.Windows.Forms.TableLayoutPanel
$tlpLeft.Dock = 'Fill'
$tlpLeft.ColumnCount = 1
$tlpLeft.RowCount = 2
$tlpLeft.Margin = New-Object System.Windows.Forms.Padding($GAP, $GAP, 3, $GAP)
$tlpLeft.Padding = New-Object System.Windows.Forms.Padding($GAP)
$tlpLeft.RowStyles.Clear()
$tlpLeft.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 100)) )
$tlpLeft.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 100)) )
# Device Summary

$grpSummary = New-Object System.Windows.Forms.GroupBox; $grpSummary.Text="Device Summary"
$grpSummary.AutoSize = $true
$grpSummary.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$grpSummary.Dock = 'Top'
$grpSummary.Margin = New-Object System.Windows.Forms.Padding($GAP)
$grpSummary.Padding = New-Object System.Windows.Forms.Padding($GAP)

$tlpSummary = New-Object System.Windows.Forms.TableLayoutPanel
$tlpSummary.AutoSize = $true
$tlpSummary.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$tlpSummary.Dock = 'Top'
$tlpSummary.Margin = New-Object System.Windows.Forms.Padding(0)
$tlpSummary.Padding = New-Object System.Windows.Forms.Padding(0)

$tlpSummary.ColumnCount = 2
$tlpSummary.RowCount = 0
$tlpSummary.ColumnStyles.Clear()
$tlpSummary.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpSummary.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tlpSummary.RowStyles.Clear()

function New-SummaryLabel {
  param(
    [string]$text,
    [bool]$isFirst = $false,
    [int]$topMargin = 10
  )
  $label = New-Object System.Windows.Forms.Label
  $label.Text = $text
  $label.AutoSize = $true
  $label.Margin = if($isFirst){ New-Object System.Windows.Forms.Padding(0,0,8,0) } else { New-Object System.Windows.Forms.Padding(0,$topMargin,8,0) }
  $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  return $label
}

function New-SummaryTextBox {
  param(
    [bool]$IsFirst = $false,
    [bool]$IsLast = $false,
    [bool]$ReadOnly = $true
  )
  $box = New-Object System.Windows.Forms.TextBox
  $box.Anchor = 'Top,Left,Right'
  $topMargin = if($IsFirst){ 0 } else { 8 }
  $bottomMargin = if($IsLast){ $GAP } else { 0 }
  $box.Margin = New-Object System.Windows.Forms.Padding(12,$topMargin,0,$bottomMargin)
  $box.MinimumSize = New-Object System.Drawing.Size(0,24)
  $box.Height = 24
  if($ReadOnly){
    $box.ReadOnly = $true
    $box.BackColor = [System.Drawing.Color]::White
  }
  return $box
}

function Add-SummaryRow {
  param(
    [string]$LabelText,
    [System.Windows.Forms.Control]$Control,
    [bool]$IsFirst = $false,
    [int]$LabelTopMargin = 8
  )
  $row = $tlpSummary.RowCount
  $tlpSummary.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  $tlpSummary.Controls.Add((New-SummaryLabel $LabelText $IsFirst $LabelTopMargin), 0, $row)
  $tlpSummary.Controls.Add($Control, 1, $row)
  $tlpSummary.RowCount++
}

$txtType = New-SummaryTextBox -IsFirst $true
Add-SummaryRow -LabelText 'Detected Type:' -Control $txtType -IsFirst $true -LabelTopMargin 0

$script:DeviceTypeSummaryControl = $txtType
try {
  $txtType.Add_TextChanged({
    try { Update-SearchDependentButtonStates } catch {}
  })
} catch {}
try { Update-SearchDependentButtonStates } catch {}

$nameRow = New-Object System.Windows.Forms.TableLayoutPanel
$nameRow.AutoSize = $true
$nameRow.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$nameRow.ColumnCount = 2
$nameRow.RowCount = 1
$nameRow.Margin = New-Object System.Windows.Forms.Padding(0,8,0,0)
$nameRow.ColumnStyles.Clear()
$nameRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$nameRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$nameRow.RowStyles.Clear()
$nameRow.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$nameRow.Dock = 'Fill'

$btnCopyHost = New-Object System.Windows.Forms.Button
$btnCopyHost.Text = "📋"
$btnCopyHost.Size = '28,24'
$btnCopyHost.Margin = New-Object System.Windows.Forms.Padding(0,0,2,0)
$btnCopyHost.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
$btnCopyHost.UseVisualStyleBackColor = $true
$btnCopyHost.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnCopyHost.Add_EnabledChanged({
  param($sender,$eventArgs)
  if ($sender.Enabled) {
    $sender.Cursor = [System.Windows.Forms.Cursors]::Hand
    $sender.ForeColor = [System.Drawing.SystemColors]::ControlText
  } else {
    $sender.Cursor = [System.Windows.Forms.Cursors]::Default
    $sender.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
  }
})
if ($btnCopyHost.Enabled) {
  $btnCopyHost.ForeColor = [System.Drawing.SystemColors]::ControlText
} else {
  $btnCopyHost.Cursor = [System.Windows.Forms.Cursors]::Default
  $btnCopyHost.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
}
$nameRow.Controls.Add($btnCopyHost,0,0)

$nameValueRow = New-Object System.Windows.Forms.TableLayoutPanel
$nameValueRow.AutoSize = $true
$nameValueRow.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$nameValueRow.ColumnCount = 2
$nameValueRow.RowCount = 1
$nameValueRow.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)
$nameValueRow.ColumnStyles.Clear()
$nameValueRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$nameValueRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$nameValueRow.RowStyles.Clear()
$nameValueRow.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$nameValueRow.Dock = 'Fill'

$txtHost = New-SummaryTextBox
$txtHost.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)
$txtHost.Dock = 'Fill'
$nameValueRow.Controls.Add($txtHost,0,0)

$btnFixName = New-Object ModernUI.RoundedButton
$btnFixName.Text = "Fix"
$btnFixName.Size = '60,24'
$btnFixName.Anchor = 'Top,Right'
$btnFixName.Margin = New-Object System.Windows.Forms.Padding(6,0,0,0)
$btnFixName.Enabled = $false
$nameValueRow.Controls.Add($btnFixName,1,0)

$nameRow.Controls.Add($nameValueRow,1,0)
$tip.SetToolTip($btnCopyHost, 'Copy host name to clipboard')
$btnCopyHost.Add_Click({
  $textToCopy = $txtHost.Text
  if(-not [string]::IsNullOrWhiteSpace($textToCopy)){
    [System.Windows.Forms.Clipboard]::SetText($textToCopy)
  }
})

$tlpSummary.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpSummary.Controls.Add((New-SummaryLabel 'Name:'), 0, $tlpSummary.RowCount)
$tlpSummary.Controls.Add($nameRow, 1, $tlpSummary.RowCount)
$tlpSummary.RowCount++

$txtAT = New-SummaryTextBox
$txtAT.Margin = New-Object System.Windows.Forms.Padding(12,8,0,0)
Add-SummaryRow -LabelText 'Asset Tag:' -Control $txtAT

$txtSN = New-SummaryTextBox
Add-SummaryRow -LabelText 'Serial:' -Control $txtSN

$txtParent = New-SummaryTextBox
Add-SummaryRow -LabelText 'Parent:' -Control $txtParent

$txtRITM = New-SummaryTextBox
Add-SummaryRow -LabelText 'PO RITM:' -Control $txtRITM

$txtRetire = New-SummaryTextBox
Add-SummaryRow -LabelText 'Retire Date:' -Control $txtRetire

$txtRound = New-SummaryTextBox -IsLast $true
Add-SummaryRow -LabelText 'Last Rounded:' -Control $txtRound

$grpSummary.Controls.Add($tlpSummary)
# Device Location (with City)
$grpLoc = New-Object System.Windows.Forms.GroupBox; $grpLoc.Text="Device Location"; $grpLoc.Dock='Fill'
$grpLoc.Margin = New-Object System.Windows.Forms.Padding($GAP)
$grpLoc.Padding = New-Object System.Windows.Forms.Padding($GAP)

$tlpLoc = New-Object System.Windows.Forms.TableLayoutPanel
$tlpLoc.Dock = 'Fill'
$tlpLoc.ColumnCount = 2
$tlpLoc.RowCount = 0
$tlpLoc.Margin = New-Object System.Windows.Forms.Padding(0)
$tlpLoc.Padding = New-Object System.Windows.Forms.Padding(0)
$tlpLoc.ColumnStyles.Clear()
$tlpLoc.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpLoc.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tlpLoc.RowStyles.Clear()
$grpLoc.Controls.Add($tlpLoc)

function Add-LocLabel([string]$text, [bool]$isFirst){
  $label = New-Object System.Windows.Forms.Label
  $label.Text = $text
  $label.AutoSize = $true
  $label.Margin = if($isFirst){ New-Object System.Windows.Forms.Padding(0,0,0,0) } else { New-Object System.Windows.Forms.Padding(0,8,0,0) }
  $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  return $label
}

function New-LocTextBox {
  param([bool]$isFirst = $false)
  $box = New-Object System.Windows.Forms.TextBox
  $box.ReadOnly = $true
  $box.BackColor = [System.Drawing.Color]::White
  $box.Dock = 'Fill'
  $box.Margin = if($isFirst){ New-Object System.Windows.Forms.Padding(12,0,0,0) } else { New-Object System.Windows.Forms.Padding(12,8,0,0) }
  $box.MinimumSize = New-Object System.Drawing.Size(0,24)
  $box.Height = 24
  return $box
}

function New-LocCombo {
  param([bool]$isFirst = $false)
  $combo = New-Object System.Windows.Forms.ComboBox
  $combo.Dock = 'Fill'
  $combo.Margin = if($isFirst){ New-Object System.Windows.Forms.Padding(12,0,0,0) } else { New-Object System.Windows.Forms.Padding(12,8,0,0) }
  $combo.Visible = $false
  $combo.DropDownStyle = 'DropDown'
  $combo.Tag = 'AllowDropDown'
  return $combo
}

# City
$tlpLoc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$cityRowIndex = $tlpLoc.RowCount
$tlpLoc.Controls.Add((Add-LocLabel 'City:' $true), 0, $cityRowIndex)
$txtCity = New-LocTextBox -isFirst $true
$tlpLoc.Controls.Add($txtCity, 1, $cityRowIndex)
$cmbCity = New-LocCombo -isFirst $true
$tlpLoc.Controls.Add($cmbCity, 1, $cityRowIndex)
$tlpLoc.RowCount++

# Location
$tlpLoc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$locationRowIndex = $tlpLoc.RowCount
$tlpLoc.Controls.Add((Add-LocLabel 'Location:' $false), 0, $locationRowIndex)
$txtLocation = New-LocTextBox
$tlpLoc.Controls.Add($txtLocation, 1, $locationRowIndex)
$cmbLocation = New-LocCombo
$tlpLoc.Controls.Add($cmbLocation, 1, $locationRowIndex)
$tlpLoc.RowCount++

# Building
$tlpLoc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$buildingRowIndex = $tlpLoc.RowCount
$tlpLoc.Controls.Add((Add-LocLabel 'Building:' $false), 0, $buildingRowIndex)
$txtBldg = New-LocTextBox
$tlpLoc.Controls.Add($txtBldg, 1, $buildingRowIndex)
$cmbBuilding = New-LocCombo
$tlpLoc.Controls.Add($cmbBuilding, 1, $buildingRowIndex)
$tlpLoc.RowCount++

# Floor
$tlpLoc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$floorRowIndex = $tlpLoc.RowCount
$tlpLoc.Controls.Add((Add-LocLabel 'Floor:' $false), 0, $floorRowIndex)
$txtFloor = New-LocTextBox
$tlpLoc.Controls.Add($txtFloor, 1, $floorRowIndex)
$cmbFloor = New-LocCombo
$tlpLoc.Controls.Add($cmbFloor, 1, $floorRowIndex)
$tlpLoc.RowCount++

# Room
$tlpLoc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$roomRowIndex = $tlpLoc.RowCount
$tlpLoc.Controls.Add((Add-LocLabel 'Room:' $false), 0, $roomRowIndex)
$txtRoom = New-LocTextBox
$tlpLoc.Controls.Add($txtRoom, 1, $roomRowIndex)
$cmbRoom = New-LocCombo
$tlpLoc.Controls.Add($cmbRoom, 1, $roomRowIndex)
$tlpLoc.RowCount++

# Department
$tlpLoc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$deptRowIndex = $tlpLoc.RowCount
$tlpLoc.Controls.Add((Add-LocLabel 'Department:' $false), 0, $deptRowIndex)
$txtDept = New-LocTextBox
$tlpLoc.Controls.Add($txtDept, 1, $deptRowIndex)
$cmbDept = New-LocCombo
$tlpLoc.Controls.Add($cmbDept, 1, $deptRowIndex)
$tlpLoc.RowCount++

if(-not $btnEditLoc){
  $btnEditLoc = New-Object ModernUI.RoundedButton
  $btnEditLoc.Text = 'Edit Location'
  $btnEditLoc.Size = '120,26'
}
$btnEditLoc.Anchor = 'Top,Right'
$btnEditLoc.Margin = New-Object System.Windows.Forms.Padding(0,8,0,0)
$tlpLoc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpLoc.Controls.Add($btnEditLoc, 0, $tlpLoc.RowCount)
$tlpLoc.SetColumnSpan($btnEditLoc, 2)
$tlpLoc.RowCount++

# Editable combos
$cmbCity.Visible=$false; $cmbLocation.Visible=$false; $cmbBuilding.Visible=$false; $cmbFloor.Visible=$false; $cmbRoom.Visible=$false; $cmbDept.Visible=$false
Populate-Department-Combo ''
# Left/Right compose
$tlpLeft.Controls.Add($grpSummary,0,0)
$tlpLeft.Controls.Add($grpLoc,0,1)
# Right column stack
$tlpRight = New-Object System.Windows.Forms.TableLayoutPanel
$tlpRight.Dock = 'Fill'
$tlpRight.ColumnCount = 1
$tlpRight.RowCount = 2
$tlpRight.Margin = New-Object System.Windows.Forms.Padding(3, $GAP, $GAP, $GAP)
$tlpRight.Padding = New-Object System.Windows.Forms.Padding($GAP)
$tlpRight.RowStyles.Clear()
$tlpRight.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 100)) )
$tlpRight.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 100)) )
# Associated devices (right column, top)
$grpAssoc = New-Object System.Windows.Forms.GroupBox; $grpAssoc.Text="Associated Devices"; $grpAssoc.Dock='Fill'
$grpAssoc.Margin = New-Object System.Windows.Forms.Padding($GAP)
$grpAssoc.Padding = New-Object System.Windows.Forms.Padding($GAP)
$tlpAssoc = New-Object System.Windows.Forms.TableLayoutPanel
$tlpAssoc.Dock = 'Fill'
$tlpAssoc.ColumnCount = 1
$tlpAssoc.RowCount = 2
$tlpAssoc.Margin = '0,0,0,0'
$tlpAssoc.Padding = '0,0,0,0'
$tlpAssoc.ColumnStyles.Clear()
$tlpAssoc.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)) )
$tlpAssoc.RowStyles.Clear()
$tlpAssoc.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
$tlpAssoc.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 200)) )
$assocToolbarPanel = New-Object System.Windows.Forms.Panel
$assocToolbarPanel.Dock = 'Fill'
$assocToolbarPanel.AutoSize = $false
$assocToolbarPanel.Margin = '0,0,0,6'
$assocToolbarPanel.Padding = '0,0,0,0'
$assocToolbarPanel.Height = 40
try { $assocToolbarPanel.MinimumSize = New-Object System.Drawing.Size(0, 40) } catch {}
$btnAddPeripheral = New-Object ModernUI.RoundedButton
$btnAddPeripheral.Text   = 'Add Peripheral'
$btnAddPeripheral.Size   = '140,32'
$btnAddPeripheral.Margin = '0,0,8,0'
$btnAddPeripheral.Anchor = 'Left'
$btnAddPeripheral.BackColor = [System.Drawing.SystemColors]::Control
$btnAddPeripheral.ForeColor = [System.Drawing.SystemColors]::ControlText
$btnRemove = New-Object ModernUI.RoundedButton
$btnRemove.Text   = 'Remove Peripheral'
$btnRemove.Size   = '160,32'
$btnRemove.Margin = '0,0,0,0'
$btnRemove.Anchor = 'Left'
$btnRemove.BackColor = [System.Drawing.SystemColors]::Control
$btnRemove.ForeColor = [System.Drawing.SystemColors]::ControlText
$assocButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$assocButtonsPanel.Dock = 'Left'
$assocButtonsPanel.AutoSize = $true
$assocButtonsPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$assocButtonsPanel.WrapContents = $false
$assocButtonsPanel.FlowDirection = 'LeftToRight'
$assocButtonsPanel.Margin = '0,0,0,0'
$assocButtonsPanel.Padding = '0,4,0,0'
$assocButtonsPanel.Controls.Add($btnAddPeripheral)
$assocButtonsPanel.Controls.Add($btnRemove)
$assocToolbarPanel.Controls.Add($assocButtonsPanel)

Set-SearchTextButtonBaseState -Button $btnAddPeripheral -BaseEnabled $false
Set-SearchTextButtonBaseState -Button $btnRemove -BaseEnabled $true
$assocGridPanel = New-Object System.Windows.Forms.Panel
$assocGridPanel.Dock = 'Fill'
$assocGridPanel.Margin = '0,0,0,0'
$assocGridPanel.Padding = '0,0,0,0'
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Dock='Fill'; $dgv.AutoGenerateColumns=$false; $dgv.AllowUserToAddRows=$false; $dgv.ReadOnly=$true
$dgv.SelectionMode=[System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgv.MultiSelect=$false; $dgv.RowHeadersVisible=$false; $dgv.BackgroundColor=[System.Drawing.Color]::White; $dgv.BorderStyle='None'
$dgv.AutoSizeColumnsMode='Fill'
$dgv.AutoSizeRowsMode=[System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
$dgv.ScrollBars=[System.Windows.Forms.ScrollBars]::Vertical
$dgv.EnableHeadersVisualStyles = $false
$dgv.GridColor = [System.Drawing.Color]::FromArgb(230,234,238)
$dgv.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$dgv.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
$dgv.RowTemplate.Height = 28
$dgv.AllowUserToResizeRows = $false
$dgv.Margin = '0,0,0,0'
try {
  $headerStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
  $headerStyle.BackColor = [System.Drawing.Color]::FromArgb(246,247,249)
  $headerStyle.ForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
  $headerStyle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
  $headerStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(229,241,251)
  $headerStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
  $headerStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
  $dgv.ColumnHeadersDefaultCellStyle = $headerStyle
} catch {}
try {
  $cellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
  $cellStyle.BackColor = [System.Drawing.Color]::White
  $cellStyle.ForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
  $cellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(229,241,251)
  $cellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
  $cellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9)
  $dgv.DefaultCellStyle = $cellStyle
} catch {}
try {
  $altStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
  $altStyle.BackColor = [System.Drawing.Color]::FromArgb(250,252,255)
  $altStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(229,241,251)
  $altStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
  $dgv.AlternatingRowsDefaultCellStyle = $altStyle
} catch {}
# Enable double buffering to reduce flicker
try { $dgv.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags] 'NonPublic,Instance').SetValue($dgv, $true, $null) } catch {}
function New-TextCol([string]$name,[string]$header,[int]$width,[bool]$ro=$true){
  $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $col.Name=$name; $col.HeaderText=$header; $col.Width=[math]::Max($width,60); $col.MinimumWidth=60; $col.ReadOnly=$ro
  return $col
}
$dgv.Columns.Add((New-TextCol 'Role' 'Role' 70))       | Out-Null
$dgv.Columns.Add((New-TextCol 'Type' 'Type' 90))       | Out-Null
$dgv.Columns.Add((New-TextCol 'Name' 'Name' 140))      | Out-Null
$dgv.Columns.Add((New-TextCol 'AssetTag' 'Asset Tag' 120)) | Out-Null
$dgv.Columns.Add((New-TextCol 'Serial' 'Serial' 120))  | Out-Null
$dgv.Columns.Add((New-TextCol 'RITM' 'RITM' 100))      | Out-Null
$dgv.Columns.Add((New-TextCol 'Retire' 'Retire' 120)) | Out-Null
try{
  $dgv.Columns['Role'].FillWeight   = 60
  $dgv.Columns['Type'].FillWeight   = 90
  $dgv.Columns['Name'].FillWeight   = 90
  $dgv.Columns['AssetTag'].FillWeight = 120
  $dgv.Columns['Serial'].FillWeight = 120
  $dgv.Columns['RITM'].FillWeight   = 170
  $dgv.Columns['Retire'].FillWeight = 110
} catch {}
$assocGridPanel.Controls.Add($dgv)
$cards = New-Object System.Windows.Forms.FlowLayoutPanel
$cards.AutoScroll=$true; $cards.WrapContents=$true; $cards.FlowDirection='LeftToRight'
$cards.Visible = $false
$tlpAssoc.Controls.Add($assocToolbarPanel,0,0)
$tlpAssoc.Controls.Add($assocGridPanel,0,1)
$grpAssoc.Controls.Add($tlpAssoc)
# Rounding group
$grpMaint = New-Object System.Windows.Forms.GroupBox; $grpMaint.Text="Device Rounding"; $grpMaint.Dock='Fill'
$grpMaint.Margin = New-Object System.Windows.Forms.Padding($GAP)
$grpMaint.Padding = New-Object System.Windows.Forms.Padding(12)

$lblMaintType=New-Object System.Windows.Forms.Label; $lblMaintType.Text='Maintenance Type'; $lblMaintType.AutoSize=$true
$cmbMaintType=New-Object System.Windows.Forms.ComboBox; $cmbMaintType.DropDownStyle='DropDownList'; $cmbMaintType.Dock='None'; $cmbMaintType.Anchor='Left'
$cmbMaintType.Items.AddRange(@('Excluded','General Rounding','Mobile Cart','Critical Clinical'))
$cmbMaintType.TabIndex = 0
$cmbMaintType.MinimumSize = New-Object System.Drawing.Size(0,0)

$lblChkStatus=New-Object System.Windows.Forms.Label; $lblChkStatus.Text="Check Status"; $lblChkStatus.AutoSize=$true
$cmbChkStatus=New-Object System.Windows.Forms.ComboBox; $cmbChkStatus.DropDownStyle='DropDownList'; $cmbChkStatus.Dock='None'; $cmbChkStatus.Anchor='Left'
$cmbChkStatus.Items.AddRange(@(
  "Complete",
  "Inaccessible - Asset not found",
  "Inaccessible - In storage",
  "Inaccessible - In use by Customer",
  "Inaccessible - Laptop is not onsite",
  "Inaccessible - Other",
  "Inaccessible - Restricted area",
  "Inaccessible - Room locked - Card Swipe",
 "Inaccessible - Room locked - Key Lock",
  "Inaccessible - Under renovation",
  "Inaccessible - User working at home",
  "Pending Repair"
)); $cmbChkStatus.SelectedIndex=0
$cmbChkStatus.TabIndex = 1
$cmbChkStatus.MinimumSize = New-Object System.Drawing.Size(0,0)

function Set-ComboPreferredWidth {
  param(
    [System.Windows.Forms.ComboBox]$Combo,
    [int]$Padding = 15,
    [int]$AdditionalWidth = 0
  )

  $maxWidth = 0
  foreach($item in $Combo.Items){
    $size = [System.Windows.Forms.TextRenderer]::MeasureText($item.ToString(), $Combo.Font)
    if($size.Width -gt $maxWidth){ $maxWidth = $size.Width }
  }

  $desiredWidth = $maxWidth + $Padding + $AdditionalWidth
  $Combo.Width = $desiredWidth
  $Combo.MinimumSize = New-Object System.Drawing.Size($desiredWidth, 0)
  $Combo.DropDownWidth = [Math]::Max($Combo.DropDownWidth, $desiredWidth)
}

Set-ComboPreferredWidth -Combo $cmbMaintType -AdditionalWidth 30
Set-ComboPreferredWidth -Combo $cmbChkStatus -AdditionalWidth 100

$lblTime=New-Object System.Windows.Forms.Label; $lblTime.Text="Rounding Time (min)"; $lblTime.AutoSize=$true
$numTime=New-Object System.Windows.Forms.NumericUpDown; $numTime.Minimum=1; $numTime.Maximum=120; $numTime.Value=3; $numTime.Width=120
$numTime.TabIndex = 2

$script:RoundingBaseMinutes = 3
$script:RoundingStartTime = $null
$script:RoundingTimer = New-Object System.Windows.Forms.Timer
$script:RoundingTimer.Interval = 1000
$script:IsUpdatingRoundingTime = $false

function Stop-RoundingTimer {
  if ($script:RoundingTimer) { $script:RoundingTimer.Stop() }
  $script:RoundingStartTime = $null
}

function Set-RoundingTimerBase {
  if (-not $numTime) { return }
  try {
    $base = [decimal]$script:RoundingBaseMinutes
    if ($base -lt $numTime.Minimum) { $base = $numTime.Minimum }
    if ($base -gt $numTime.Maximum) { $base = $numTime.Maximum }
    if ($numTime.Value -ne $base) {
      $script:IsUpdatingRoundingTime = $true
      try { $numTime.Value = $base }
      finally { $script:IsUpdatingRoundingTime = $false }
    }
  } catch {}
}

function Reset-RoundingTimer {
  Stop-RoundingTimer
  Set-RoundingTimerBase
}

function Start-RoundingTimer {
  Stop-RoundingTimer
  Set-RoundingTimerBase
  $script:RoundingStartTime = [DateTime]::UtcNow
  if ($script:RoundingTimer) { $script:RoundingTimer.Start() }
}

$script:RoundingTimer.Add_Tick({
  $start = $script:RoundingStartTime
  if (-not $start) { return }
  $elapsed = [DateTime]::UtcNow - $start
  $elapsedMinutes = [math]::Floor($elapsed.TotalMinutes)
  $target = [decimal]$script:RoundingBaseMinutes
  if ($elapsedMinutes -ge $script:RoundingBaseMinutes) {
    $max = [double]$numTime.Maximum
    $target = [decimal]([math]::Min($elapsedMinutes, $max))
  }
  $current = [decimal]$numTime.Value
  if ($target -gt $current) {
    $script:IsUpdatingRoundingTime = $true
    try { $numTime.Value = $target }
    finally { $script:IsUpdatingRoundingTime = $false }
  }
})

$numTime.Add_ValueChanged({
  if ($script:IsUpdatingRoundingTime) { return }
  try {
    $script:RoundingBaseMinutes = [decimal]$numTime.Value
  } catch {}
  Stop-RoundingTimer
})

$chkCable=New-Object System.Windows.Forms.CheckBox; $chkCable.Text="Validate Cable Management"; $chkCable.AutoSize=$true; $chkCable.TabIndex = 3
$chkCableNeeded=New-Object System.Windows.Forms.CheckBox; $chkCableNeeded.Text="Cabling Needed"; $chkCableNeeded.AutoSize=$true; $chkCableNeeded.TabIndex = 4
$chkCart=New-Object System.Windows.Forms.CheckBox; $chkCart.Text="Check Physical Cart Is Working"; $chkCart.AutoSize=$true; $chkCart.TabIndex = 5
$chkLabels=New-Object System.Windows.Forms.CheckBox; $chkLabels.Text="Ensure monitor appropriately labelled"; $chkLabels.AutoSize=$true; $chkLabels.TabIndex = 6
$chkPeriph=New-Object System.Windows.Forms.CheckBox; $chkPeriph.Text="Validate peripherals are connected and working"; $chkPeriph.AutoSize=$true; $chkPeriph.TabIndex = 7

$btnCheckComplete=New-Object ModernUI.RoundedButton; $btnCheckComplete.Text="Check Complete"; $btnCheckComplete.Size='150,36'; $btnCheckComplete.TabIndex = 8
$btnSave=New-Object ModernUI.RoundedButton; $btnSave.Text="Save Event"; $btnSave.Size='132,36'; $btnSave.TabIndex = 9
$btnManualRound=New-Object ModernUI.RoundedButton; $btnManualRound.Text="Manual Round"; $btnManualRound.Size='140,36'; $btnManualRound.Enabled=$false; $btnManualRound.TabIndex = 10
$btnRoundNow=New-Object ModernUI.RoundedButton; $btnRoundNow.Text="Round Now"; $btnRoundNow.Size='140,36'; $btnRoundNow.Enabled=$false; $btnRoundNow.TabIndex = 11

Set-SearchTextButtonBaseState -Button $btnSave -BaseEnabled $true

$btnCheckComplete.BackColor = [System.Drawing.SystemColors]::Control
$btnCheckComplete.ForeColor = [System.Drawing.SystemColors]::ControlText
$btnSave.BackColor = [System.Drawing.SystemColors]::Control
$btnSave.ForeColor = [System.Drawing.SystemColors]::ControlText
$btnManualRound.BackColor = [System.Drawing.SystemColors]::Control
$btnManualRound.ForeColor = [System.Drawing.SystemColors]::ControlText
$btnRoundNow.BackColor = [System.Drawing.SystemColors]::Control
$btnRoundNow.ForeColor = [System.Drawing.SystemColors]::ControlText

$lblComments=New-Object System.Windows.Forms.Label; $lblComments.Text='Comments'; $lblComments.AutoSize=$true; $lblComments.TabIndex = 12
$txtComments = New-Object System.Windows.Forms.TextBox; $txtComments.Multiline=$true; $txtComments.AcceptsReturn=$true; $txtComments.ScrollBars='Vertical'; $txtComments.Dock='Fill'; $txtComments.TabIndex=13; $txtComments.WordWrap = $true

$layoutMaint = New-Object System.Windows.Forms.TableLayoutPanel
$layoutMaint.Dock = 'Fill'
$layoutMaint.ColumnCount = 1
$layoutMaint.RowCount = 6
$layoutMaint.ColumnStyles.Clear()
$layoutMaint.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$layoutMaint.RowStyles.Clear()
$layoutMaint.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layoutMaint.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layoutMaint.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layoutMaint.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layoutMaint.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layoutMaint.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$rowCombos = New-Object System.Windows.Forms.TableLayoutPanel
$rowCombos.Dock = 'Fill'
$rowCombos.AutoSize = $true
$rowCombos.AutoSizeMode = 'GrowAndShrink'
$rowCombos.ColumnCount = 4
$rowCombos.RowCount = 1
$rowCombos.ColumnStyles.Clear()
$rowCombos.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rowCombos.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rowCombos.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rowCombos.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rowCombos.RowStyles.Clear()
$rowCombos.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rowCombos.Controls.Add($lblMaintType,0,0)
$rowCombos.Controls.Add($cmbMaintType,1,0)
$rowCombos.Controls.Add($lblChkStatus,2,0)
$rowCombos.Controls.Add($cmbChkStatus,3,0)
$lblMaintType.Margin = New-Object System.Windows.Forms.Padding(0,0,12,0)
$cmbMaintType.Margin = New-Object System.Windows.Forms.Padding(0,0,24,0)
$lblChkStatus.Margin = New-Object System.Windows.Forms.Padding(0,0,12,0)
$cmbChkStatus.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)

$rowTime = New-Object System.Windows.Forms.TableLayoutPanel
$rowTime.Dock = 'Fill'
$rowTime.AutoSize = $true
$rowTime.AutoSizeMode = 'GrowAndShrink'
$rowTime.ColumnCount = 2
$rowTime.ColumnStyles.Clear()
$rowTime.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rowTime.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rowTime.Controls.Add($lblTime,0,0)
$rowTime.Controls.Add($numTime,1,0)
$rowTime.Margin = New-Object System.Windows.Forms.Padding(0,12,0,0)
$lblTime.Margin = New-Object System.Windows.Forms.Padding(0,0,12,0)
$numTime.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)

$rowChecks = New-Object System.Windows.Forms.TableLayoutPanel
$rowChecks.Dock = 'Fill'
$rowChecks.AutoSize = $true
$rowChecks.AutoSizeMode = 'GrowAndShrink'
$rowChecks.ColumnCount = 2
$rowChecks.RowCount = 3
$rowChecks.ColumnStyles.Clear()
$rowChecks.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$rowChecks.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$rowChecks.RowStyles.Clear()
$rowChecks.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rowChecks.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rowChecks.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rowChecks.Controls.Add($chkCable,0,0)
$rowChecks.Controls.Add($chkCableNeeded,1,0)
$rowChecks.Controls.Add($chkLabels,0,1)
$rowChecks.Controls.Add($chkCart,1,1)
$rowChecks.Controls.Add($chkPeriph,0,2)
$rowChecks.SetColumnSpan($chkPeriph,2)
$rowChecks.Margin = New-Object System.Windows.Forms.Padding(0,12,0,0)
$chkCable.Margin = New-Object System.Windows.Forms.Padding(0,0,12,0)
$chkCableNeeded.Margin = New-Object System.Windows.Forms.Padding(12,0,0,0)
$chkLabels.Margin = New-Object System.Windows.Forms.Padding(0,6,12,0)
$chkCart.Margin = New-Object System.Windows.Forms.Padding(12,6,0,0)
$chkPeriph.Margin = New-Object System.Windows.Forms.Padding(0,6,0,0)

$actionsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$actionsPanel.Dock = 'Fill'
$actionsPanel.AutoSize = $true
$actionsPanel.AutoSizeMode = 'GrowAndShrink'
$actionsPanel.WrapContents = $false
$actionsPanel.FlowDirection = 'LeftToRight'
$actionsPanel.Margin = New-Object System.Windows.Forms.Padding(0,12,0,0)
$btnCheckComplete.Margin = New-Object System.Windows.Forms.Padding(0,0,12,0)
$btnSave.Margin = New-Object System.Windows.Forms.Padding(0,0,12,0)
$btnManualRound.Margin = New-Object System.Windows.Forms.Padding(0,0,12,0)
$btnRoundNow.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)
$actionsPanel.Controls.Add($btnCheckComplete)
$actionsPanel.Controls.Add($btnSave)
$actionsPanel.Controls.Add($btnManualRound)
$actionsPanel.Controls.Add($btnRoundNow)

Update-RoundNowButtonState

$lblComments.Margin = New-Object System.Windows.Forms.Padding(0,12,0,0)
$txtComments.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)

$layoutMaint.Controls.Add($rowCombos,0,0)
$layoutMaint.Controls.Add($rowTime,0,1)
$layoutMaint.Controls.Add($rowChecks,0,2)
$layoutMaint.Controls.Add($actionsPanel,0,3)
$layoutMaint.Controls.Add($lblComments,0,4)
$layoutMaint.Controls.Add($txtComments,0,5)

$grpMaint.Controls.Add($layoutMaint)
$txtComments.Add_TextChanged({
  if(Get-Command Update-RoundingCommentsLayout -ErrorAction SilentlyContinue){
    Update-RoundingCommentsLayout
  }
})
$txtComments.Add_SizeChanged({
  if(Get-Command Update-RoundingCommentsLayout -ErrorAction SilentlyContinue){
    Update-RoundingCommentsLayout
  }
})
# Compose columns
$tlpRight.Controls.Add($grpAssoc,0,0)
$tlpRight.Controls.Add($grpMaint,0,1)
$splitter.Panel1.Controls.Add($tlpLeft)
$splitter.Panel2.Controls.Add($tlpRight)
# StatusStrip
$status = New-Object System.Windows.Forms.StatusStrip
$statusPathLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusPathLabel.Spring = $true
$statusPathLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusPathLabel.Margin = New-Object System.Windows.Forms.Padding(6,3,6,3)
$statusPathLabel.Text = $statusPathLabelDefault
$status.Items.Add($statusPathLabel) | Out-Null
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
$statusLabel.Margin = New-Object System.Windows.Forms.Padding(6,3,6,3)
$statusLabel.Text = "Ready"
$status.Items.Add($statusLabel) | Out-Null
# Add to form
$form.SuspendLayout()
$form.Controls.Add($splitter)   # Fill
$form.Add_Shown({
  foreach($c in @($cmbDept,$cmbDepartment,$ddlDept,$ddlDepartment)){ if ($c) { $c.Visible = $false } }
  foreach($t in @($txtDept,$txtDepartment)){ if ($t) { $t.Visible = $true } }
  if ($cmbDept) { $cmbDept.Visible = $false }
  if ($txtDept) { $txtDept.Visible = $true }
  if ($statusPathLabel) {
    $statusPathLabel.Text = "Data: " + $DataFolder + "    |    Output: " + $OutputFolder
  }
})
$form.Controls.Add($panelTop)  # Top
$form.Controls.Add($status)    # Bottom
$form.ResumeLayout($true)
$form.PerformLayout()
# -------- Responsive row sizing (DPI aware) --------
function Apply-ResponsiveHeights {
  try {
    # Let the left column size itself to its content while keeping the right column responsive
    $tlpLeft.RowStyles[0].SizeType = [System.Windows.Forms.SizeType]::AutoSize
    $tlpLeft.RowStyles[0].Height   = 0
    $tlpLeft.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::AutoSize
    $tlpLeft.RowStyles[1].Height   = 0
    $rowsShown = [Math]::Max($dgv.Rows.Count, 1)
    $assocInfo = Size-AssocForRows $rowsShown
    $assocTarget = 0
    if($assocInfo -and $assocInfo.Target){
      $assocTarget = [Math]::Max([int]$assocInfo.Target, 0)
    }
    $stripHeight = $assocToolbarPanel.PreferredSize.Height
    if($stripHeight -le 0){
      $stripHeight = [Math]::Max($assocButtonsPanel.PreferredSize.Height, [Math]::Max($btnAddPeripheral.PreferredSize.Height, $btnRemove.PreferredSize.Height))
    }
    if($stripHeight -le 0){
      $stripHeight = [Math]::Max($assocToolbarPanel.Height, [Math]::Max($assocButtonsPanel.Height, [Math]::Max($btnAddPeripheral.Height, $btnRemove.Height)))
    }
    if($stripHeight -le 0){ $stripHeight = 36 }
    $assocPadding = $grpAssoc.Padding.Vertical + $grpAssoc.Margin.Vertical + $tlpAssoc.Margin.Vertical + $tlpAssoc.Padding.Vertical + $assocToolbarPanel.Margin.Vertical + $assocToolbarPanel.Padding.Vertical + $assocGridPanel.Margin.Vertical + $assocGridPanel.Padding.Vertical
    $minAssoc   = [Math]::Max($assocTarget + $stripHeight + $assocPadding, 220)
    $minRound   = [Math]::Max($grpMaint.PreferredSize.Height, 220)
    $tlpRight.RowStyles[0].SizeType = [System.Windows.Forms.SizeType]::Absolute
    $tlpRight.RowStyles[0].Height   = $minAssoc
    $tlpRight.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::Absolute
    $tlpRight.RowStyles[1].Height   = $minRound
  } catch { }
}
function Update-RoundingCommentsLayout {
  try {
    if(-not $txtComments){ return }
    if($txtComments.Dock -eq [System.Windows.Forms.DockStyle]::Fill){
      Apply-ResponsiveHeights
      return
    }
    $width = [int][Math]::Max($txtComments.ClientSize.Width, 100)
    if($width -le 0){ $width = [Math]::Max($txtComments.Width - 8, 100) }
    $measureSize = New-Object System.Drawing.Size($width, 0)
    $flags = [System.Windows.Forms.TextFormatFlags]::WordBreak -bor [System.Windows.Forms.TextFormatFlags]::TextBoxControl
    $text = $txtComments.Text
    if([string]::IsNullOrEmpty($text)){ $text = ' ' }
    $measured = [System.Windows.Forms.TextRenderer]::MeasureText($text + ' ', $txtComments.Font, $measureSize, $flags)
    $padding = $txtComments.Height - $txtComments.ClientSize.Height
    if($padding -le 0){ $padding = 8 }
    $minHeight = 72
    $desired = [Math]::Max($measured.Height + $padding, $minHeight)
    if([Math]::Abs($desired - $txtComments.Height) -gt 2){
      $txtComments.Height = $desired
    }
    Apply-ResponsiveHeights
  } catch { }
}
function Get-AssocSizing([int]$rows){
  $result = [pscustomobject]@{ Target = 0; Grid = 0; Rows = 0 }
  try{
    if($rows -lt 1){ $rows = 1 }
    $result.Rows = $rows
    $visibleHeight = 0
    $visibleCount  = 0
    foreach($row in $dgv.Rows){
      if($null -eq $row){ continue }
      if($row.IsNewRow){ continue }
      if(-not $row.Visible){ continue }
      $visibleHeight += [int][Math]::Max($row.Height, 0)
      $visibleCount++
      if($visibleCount -ge $rows){ break }
    }
    if($visibleCount -lt $rows){
      $rowH = [Math]::Max($dgv.RowTemplate.Height, 22)
      $visibleHeight += ($rows - $visibleCount) * $rowH
    }
    $hdrH = 0
    if($dgv.ColumnHeadersVisible){
      $hdrH = [Math]::Max($dgv.ColumnHeadersHeight, 24)
    }
    $gridH = $visibleHeight + $hdrH
    $gridH += [Math]::Max([System.Windows.Forms.SystemInformation]::BorderSize.Height * 2, 2)
    $totalColW = 0
    foreach($c in $dgv.Columns){ if($c.Visible){ $totalColW += [int]$c.Width } }
    $clientW = [Math]::Max($dgv.ClientSize.Width, $dgv.DisplayRectangle.Width)
    if($clientW -le 0){ $clientW = $dgv.Width }
    if($totalColW -gt $clientW){
      $gridH += [System.Windows.Forms.SystemInformation]::HorizontalScrollBarHeight
    }
    $target = $gridH + $dgv.Margin.Vertical + $tlpAssoc.Padding.Vertical + $assocGridPanel.Padding.Vertical
    $result.Target = [Math]::Max([int]$target, 0)
    $result.Grid   = [Math]::Max([int]$gridH, 0)
  } catch { }
  return $result
}
function Size-AssocForRows([int]$rows){
  $info = Get-AssocSizing $rows
  try{
    if($info.Target -gt 0){
      $tlpAssoc.RowStyles[0].SizeType = [System.Windows.Forms.SizeType]::AutoSize
      $tlpAssoc.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::Absolute
      $tlpAssoc.RowStyles[1].Height   = $info.Grid
      try { $assocGridPanel.MinimumSize = New-Object System.Drawing.Size(0, [Math]::Max([int]$info.Grid, 0)) } catch {}
    }
  } catch { }
  return $info
}
$form.Add_Shown({
  Apply-ResponsiveHeights
  Size-AssocForRows([Math]::Max($dgv.Rows.Count,1)) | Out-Null
})
$form.Add_Shown({ Update-RoundingCommentsLayout })
# -------- UI logic ---------
function Update-Counters(){ $locCount = $script:LocationRows.Count; $lblDataStatus.Text = ("Computers: {0} | Monitors: {1} | Mics: {2} | Scanners: {3} | Carts: {4} | Locations: {5}" -f `
    $script:Computers.Count,$script:Monitors.Count,$script:Mics.Count,$script:Scanners.Count,$script:Carts.Count,$locCount) }
function Update-CartCheckbox-State([object]$parentRec){
  $chkCart.Checked = $false; $chkCart.Enabled = $false
  if($parentRec -and $parentRec.name -match '^(?i)AO'){ $chkCart.Enabled = $true }
}
function Resolve-ComputerRecord([object]$rec){
  if(-not $rec){ return $null }
  try {
    if($rec.PSObject.Properties['u_device_rounding']){ return $rec }
  } catch {}
  try {
    if($rec.asset_tag){
      $key = $rec.asset_tag.Trim().ToUpper()
      if($script:ComputerByAsset.ContainsKey($key)){ return $script:ComputerByAsset[$key] }
    }
  } catch {}
  try {
    if($rec.name){
      foreach($k in (HostnameKeyVariants $rec.name)){
        if($script:ComputerByName.ContainsKey($k)){ return $script:ComputerByName[$k] }
      }
    }
  } catch {}
  return $null
}
function Set-ComboSelectionCaseInsensitive([System.Windows.Forms.ComboBox]$combo,[string]$value){
  if(-not $combo){ return }
  if([string]::IsNullOrWhiteSpace($value)){
    try { $combo.SelectedIndex = -1 } catch {}
    return
  }
  $target = $value.Trim()
  $selectedIndex = -1
  for($i = 0; $i -lt $combo.Items.Count; $i++){
    $item = $combo.Items[$i]
    if(-not $item){ continue }
    $text = $item.ToString()
    if($text.Trim().ToUpper() -eq $target.ToUpper()){
      $selectedIndex = $i
      $target = $text
      break
    }
  }
  if($selectedIndex -ge 0){
    try { $combo.SelectedIndex = $selectedIndex } catch { $combo.SelectedItem = $target }
  } else {
    [void]$combo.Items.Add($value)
    try { $combo.SelectedIndex = ($combo.Items.Count - 1) } catch {}
  }
}
function Get-MaintenanceTypeOrDefault([string]$maintenanceType,[string]$deviceName){
  $mt = ''
  if(-not [string]::IsNullOrWhiteSpace($maintenanceType)){
    $mt = $maintenanceType.Trim()
  }
  if(-not [string]::IsNullOrWhiteSpace($mt)){
    return $mt
  }
  if(-not [string]::IsNullOrWhiteSpace($deviceName)){
    try {
      if($deviceName.Trim() -match '^(?i)AO'){
        return 'Mobile Cart'
      }
    } catch {}
  }
  return 'General Rounding'
}
function Update-MaintenanceTypeSelection([object]$displayRec,[object]$parentRec){
  $displayName = ''
  try {
    if($displayRec -and $displayRec.name){ $displayName = [string]$displayRec.name }
  } catch {}
  
  $displayIsComputerOrTangent = $false
  try {
    if($displayRec){
      $type = ''
      $kind = ''
      try { if($displayRec.PSObject.Properties['Type']){ $type = '' + $displayRec.Type } } catch {}
      try { if($displayRec.PSObject.Properties['Kind']){ $kind = '' + $displayRec.Kind } } catch {}
      if($type -eq 'Computer' -or $kind -eq 'Computer'){
        $displayIsComputerOrTangent = $true
      }
      elseif(-not [string]::IsNullOrWhiteSpace($displayName) -and ($displayName.Trim() -match '^(?i)AO')){
        $displayIsComputerOrTangent = $true
      }
    }
    } catch {}

  $source = $null
  if($displayIsComputerOrTangent){
    $source = Resolve-ComputerRecord $displayRec
  } else {
    $source = Resolve-ComputerRecord $parentRec
  }
  if(-not $source){
    # Fall back to whichever record we did not already try.
    if($displayIsComputerOrTangent){
      $source = Resolve-ComputerRecord $parentRec
    } else {
      $source = Resolve-ComputerRecord $displayRec
    }
      }
  $mt = ''
  try {
    if($source -and $source.PSObject.Properties['u_device_rounding']){
      $mt = ('' + $source.u_device_rounding).Trim()
    }
	} catch {}

  $selection = Get-MaintenanceTypeOrDefault $mt $displayName
  Set-ComboSelectionCaseInsensitive $cmbMaintType $selection
}
function Get-RoundingUrlForParent($pc){
  if(-not $pc -or -not $pc.asset_tag){ return $null }
  $k = $pc.asset_tag.Trim().ToUpper()
  if($script:RoundingByAssetTag.ContainsKey($k)){
    $id = $script:RoundingByAssetTag[$k]
    return "https://devicerounding.nttdatanucleus.com/DeviceMaintenance/Index?DeviceId=$id"
  }
  return $null
}
function Update-ManualRoundButton($parentRec){
  if($parentRec){
    $url = Get-RoundingUrlForParent $parentRec
    if($url){ $btnManualRound.Enabled = $true; $btnManualRound.Tag = $url; $tip.SetToolTip($btnManualRound,$url); return }
  }
  $btnManualRound.Enabled = $false; $btnManualRound.Tag = $null; $tip.SetToolTip($btnManualRound,"")
}
# ---- Location Validation ----
function Get-City-ForLocation([string]$loc){
  if([string]::IsNullOrWhiteSpace($loc)){ return '' }
  $nLoc = Normalize-Field $loc
  $matches = @($script:LocationRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Location')) -eq $nLoc })
  if($matches.Count -gt 0){
    $row = $matches[$matches.Count - 1]
    return ([string](Get-LocVal $row 'City'))
  }
  if($script:LocationMasterRows.Count -gt 0){
    $masterRow = $script:LocationMasterRows | Where-Object { (Normalize-Field ([string]$_.'Location')) -eq $nLoc } | Select-Object -First 1
    if($masterRow){
      try { return '' + $masterRow.City } catch {}
    }
  }
  return ''
}
function Validate-Location($rec){
  # In edit mode, do not actively validate or repaint to avoid flicker; just reflect current text.
  if($script:editing){ return }
  # Show raw values in UI
  $cityText = ''
  try {
    if($rec -and $rec.PSObject.Properties['City']){ $cityText = '' + $rec.City }
    elseif($rec -and $rec.PSObject.Properties['location.city']){ $cityText = '' + $rec.PSObject.Properties['location.city'].Value }
  } catch {}
  if([string]::IsNullOrWhiteSpace($cityText)){
    $cityText = Get-City-ForLocation $rec.location
  }
  $txtCity.Text = $cityText
  $deptVal = ''
  try{
    if($rec){
      if($rec.PSObject.Properties['u_department_location']){ $deptVal = $rec.u_department_location }
      elseif($rec.PSObject.Properties['Department']){ $deptVal = $rec.Department }
    }
  } catch {}
  if($txtDept){ $txtDept.Text = $deptVal }
  $txtLocation.Text = $rec.location
  $txtBldg.Text     = $rec.u_building
  $txtFloor.Text    = $rec.u_floor
  $txtRoom.Text     = $rec.u_room
  try{ $cmbDept.Text = $deptVal } catch {}

  $tip.SetToolTip($txtRoom, "")
  $okC=$false; $okL=$false; $okB=$false; $okF=$false; $okR=$false
  $nLoc  = Normalize-Field $rec.location
  $nBld  = Normalize-Field $rec.u_building
  $nFlr  = Normalize-Field ([string]$rec.u_floor)
  $nRoom = Normalize-Field $rec.u_room
  if($script:LocationMasterRows.Count -gt 0){
    $rowsL = $script:LocationMasterRows | Where-Object { (Normalize-Field ([string]$_.'Location')) -eq $nLoc }
    $okL   = ($rowsL.Count -gt 0)
    if($okL){
      $okC = $true
      $masterCity = ''
      try {
        $masterCity = '' + ($rowsL | Select-Object -First 1).City
      } catch {}
      if(-not [string]::IsNullOrWhiteSpace($masterCity) -and [string]::IsNullOrWhiteSpace($txtCity.Text)){
        $txtCity.Text = $masterCity
      }
    }
    $rowsB = @()
    if($okL -and $nBld){
      $rowsB = $rowsL | Where-Object { (Normalize-Field ([string]$_.'Building')) -eq $nBld }
      $okB   = ($rowsB.Count -gt 0)
    }
    $rowsF = @()
    if($okB -and $nFlr){
      $rowsF = $rowsB | Where-Object { (Normalize-Field ([string]$_.'Floor')) -eq $nFlr }
      $okF   = ($rowsF.Count -gt 0)
    }
    if($nRoom){
      $okR = ($script:RoomsNorm -contains $nRoom)
      if(-not $okR){
        $code = Extract-RoomCode $rec.u_room
        if($code -and ($script:RoomCodes -contains $code)){
          $okR = $true
          $tip.SetToolTip($txtRoom, "Matched by room code " + $code + " (exact text differs in LocationMaster).")
        } else {
          $tip.SetToolTip($txtRoom, "Room not found in LocationMaster Room column.")
        }
      }
    } else { $okR = $false }
  }
  $txtCity.BackColor     = if($okC){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
  $txtLocation.BackColor = if($okL){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
  $txtBldg.BackColor     = if($okB){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
  $txtFloor.BackColor    = if($okF){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
  $txtRoom.BackColor     = if($okR){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
  try{
    $d = $deptVal
    $okD = $false
    if(-not $script:DepartmentListNorm -or $script:DepartmentListNorm.Count -eq 0){
      try { Load-DepartmentMaster } catch {}
    }
    if($d -and $script:DepartmentListNorm){ $okD = $script:DepartmentListNorm.Contains((Normalize-Field $d)) }
    $colorD = if($okD){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
    if($txtDept){ $txtDept.BackColor = $colorD }
    if($cmbDept){ $cmbDept.BackColor = $colorD }
  } catch {}
  Update-LastLocationSelections $txtCity.Text $txtLocation.Text $txtBldg.Text $txtFloor.Text $txtRoom.Text
}
function Refresh-AssocGrid($parentRec){
  $dgv.Rows.Clear(); if(-not $parentRec){ Size-AssocForRows(1) | Out-Null; return }
  $prow = $dgv.Rows.Add()
  $dgv.Rows[$prow].Cells['Role'].Value='Parent'
  $dgv.Rows[$prow].Cells['Type'].Value=(Get-DetectedType $parentRec)
  $dgv.Rows[$prow].Cells['Name'].Value=$parentRec.name
  $dgv.Rows[$prow].Cells['AssetTag'].Value=$parentRec.asset_tag
  $dgv.Rows[$prow].Cells['Serial'].Value=$parentRec.serial_number
  $dgv.Rows[$prow].Cells['RITM'].Value=$parentRec.RITM
  $dgv.Rows[$prow].Cells['Retire'].Value= (Fmt-DateLong $parentRec.Retire)
  $dgv.Rows[$prow].DefaultCellStyle.BackColor = $script:ThemeColors.Header
  $kids = Get-ChildrenForParent $parentRec
  foreach($childInfo in $kids){
    $ch = $null
    try { $ch = $childInfo.Record } catch { $ch = $null }
    if(-not $ch){ continue }
    $immediateParent = $parentRec
    try {
      if($childInfo.Parent){ $immediateParent = $childInfo.Parent }
    } catch {}
    $roleLabel = 'Child'
    try {
      if($childInfo.Role){ $roleLabel = $childInfo.Role }
    } catch {}
    $rowIdx = $dgv.Rows.Add()
    $r = $dgv.Rows[$rowIdx]
    $r.Cells['Role'].Value=$roleLabel
    $r.Cells['Type'].Value=$ch.Type
    if($ch.name){ $r.Cells['Name'].Value = $ch.name } else { $r.Cells['Name'].Value = '' }
    try {
      $expectedName = Compute-ProposedName $ch $immediateParent
      $defaultNameColor = if($dgv.DefaultCellStyle -and $dgv.DefaultCellStyle.ForeColor){
        $dgv.DefaultCellStyle.ForeColor
      } else {
        [System.Drawing.Color]::FromArgb(32,32,32)
      }
      if(-not [string]::IsNullOrWhiteSpace($expectedName)){
        $actualName = if($ch.name){ $ch.name.Trim() } else { '' }
        if([string]::IsNullOrWhiteSpace($actualName) -or ($actualName.Trim().ToUpper() -ne $expectedName.Trim().ToUpper())){
          $r.Cells['Name'].Style.ForeColor = [System.Drawing.Color]::IndianRed
          $r.Cells['Name'].ToolTipText = "Expected name: " + $expectedName
        } else {
          $r.Cells['Name'].Style.ForeColor = $defaultNameColor
          $r.Cells['Name'].ToolTipText = ''
        }
      } else {
        $r.Cells['Name'].Style.ForeColor = $defaultNameColor
        $r.Cells['Name'].ToolTipText = ''
      }
    } catch {}
    $r.Cells['AssetTag'].Value=$ch.asset_tag
    $r.Cells['Serial'].Value=$ch.serial_number
    if(($ch.Type -eq 'Mic') -or ($ch.Type -eq 'Scanner')){
      $r.Cells['RITM'].Value=''; $r.Cells['Retire'].Value=''
    } else {
      $ritm=$ch.RITM
      $r.Cells['RITM'].Value=$ritm; try{ if($ritm -and $ritm.Length -gt 12){ $r.Cells['RITM'].ToolTipText = $ritm } } catch {}
      $r.Cells['Retire'].Value=(Fmt-DateLong $ch.Retire)
      if([string]::IsNullOrWhiteSpace($ritm)){
        $r.Cells['RITM'].Style.ForeColor=[System.Drawing.Color]::Black
      } elseif($parentRec.RITM -and $ritm -eq $parentRec.RITM){
        $r.Cells['RITM'].Style.ForeColor=[System.Drawing.Color]::ForestGreen
      } else { $r.Cells['RITM'].Style.ForeColor=[System.Drawing.Color]::IndianRed }
    }
  }
  Size-AssocForRows([Math]::Max($dgv.Rows.Count,1)) | Out-Null
}
function Make-Card($title,$kvPairs,[System.Drawing.Color]$ritmColor,[bool]$showRITM,[bool]$showRetire,$tagPayload){
  $p = New-Object System.Windows.Forms.Panel
  $p.Width = 280; $p.Height = 160; $p.Margin = '6,6,6,6'
  $p.BackColor = $script:ThemeColors.Surface; $p.BorderStyle='FixedSingle'
  $p.Tag = $tagPayload
  $p.Add_DoubleClick({
    $ids = $_.Sender.Tag
    if($ids){
      $rec = $null
      if($ids.asset){ $key = ($ids.asset).Trim().ToUpper(); if($script:IndexByAsset.ContainsKey($key)){ $rec = $script:IndexByAsset[$key] } }
      if(-not $rec -and $ids.serial){ $key = ($ids.serial).Trim().ToUpper(); if($script:IndexBySerial.ContainsKey($key)){ $rec = $script:IndexBySerial[$key] } }
      if(-not $rec -and $ids.name){
        foreach($k in (HostnameKeyVariants $ids.name)){ if($script:IndexByName.ContainsKey($k)){ $rec = $script:IndexByName[$k]; break } }
      }
      if($rec){
        $par = Resolve-ParentComputer $rec
        Populate-UI $rec $par
        Start-RoundingTimer
      }
    }
  })
  $lblTitle = New-Object System.Windows.Forms.Label
  $lblTitle.Text = $title; $lblTitle.Font = $script:ThemeFontSemibold
  $lblTitle.AutoSize=$true; $lblTitle.Location='8,6'
  $p.Controls.Add($lblTitle)
  $y = 28
  foreach($kv in $kvPairs){
    if(($kv.Key -eq 'RITM' -and -not $showRITM) -or ($kv.Key -eq 'Retire' -and -not $showRetire)){ continue }
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize=$true; $lbl.Location = "8,$y"
    $lbl.Text = ("{0}: {1}" -f $kv.Key, $kv.Value)
    if($kv.Key -eq 'RITM' -and $kv.Value){ $lbl.ForeColor = $ritmColor }
    $p.Controls.Add($lbl); $y += 18
  }
  return $p
}
function Refresh-AssocCards($parentRec){
  if(-not $cards){ return }
  $cards.SuspendLayout(); $cards.Controls.Clear()
  if($parentRec){
    $cards.Controls.Add( (Make-Card ("Parent - " + (Get-DetectedType $parentRec)) @(
      @{Key='Name';Value=$parentRec.name},
      @{Key='Asset';Value=$parentRec.asset_tag},
      @{Key='Serial';Value=$parentRec.serial_number},
      @{Key='RITM';Value=$parentRec.RITM},
      @{Key='Retire';Value=(Fmt-DateLong $parentRec.Retire)}
    ) ([System.Drawing.Color]::Black) $true $true @{asset=$parentRec.asset_tag;serial=$parentRec.serial_number;name=$parentRec.name}) )
    $pRITM = $parentRec.RITM
    $kids = Get-ChildrenForParent $parentRec
    foreach($childInfo in $kids){
      $ch = $null
      try { $ch = $childInfo.Record } catch { $ch = $null }
      if(-not $ch){ continue }
      $roleLabel = 'Child'
      try {
        if($childInfo.Role){ $roleLabel = $childInfo.Role }
      } catch {}
      $ritm = $ch.RITM
      $col = [System.Drawing.Color]::Black
      if(-not [string]::IsNullOrWhiteSpace($ritm)){
        if($pRITM -and $ritm -eq $pRITM){ $col=[System.Drawing.Color]::ForestGreen } else { $col=[System.Drawing.Color]::IndianRed }
      }
      $showR = ($ch.Type -eq 'Monitor')
      $cards.Controls.Add( (Make-Card (($roleLabel + " - " + (Get-DetectedType $ch))) @(
        @{Key='Name';Value=$ch.name},
        @{Key='Asset';Value=$ch.asset_tag},
        @{Key='Serial';Value=$ch.serial_number},
        @{Key='RITM';Value=$ritm},
        @{Key='Retire';Value=(Fmt-DateLong $ch.Retire)}
      ) $col $showR $showR @{asset=$ch.asset_tag;serial=$ch.serial_number;name=$ch.name}) )
    }
  }
  $cards.ResumeLayout()
}
function Refresh-AssocViews($parentRec){ Refresh-AssocGrid $parentRec; Refresh-AssocCards $parentRec }
# ----- Peripheral Preview/Link/Remove + logging -----
function Normalize-UniversalSearch([string]$raw){
  if([string]::IsNullOrWhiteSpace($raw)){ return $null }
  $s = $raw.Trim().ToUpper()
  $s = $s -replace '\s',''
  $s = $s -replace '-',''
  if($s.Length -ge 2 -and $s.StartsWith('C')){
    $s = $s.Substring(0, $s.Length - 1)
  }
  return $s
}
function Resolve-PeripheralLookup([string]$query){
  $normalizedInput = Normalize-UniversalSearch $query
  $rawUpper = $null
  if(-not [string]::IsNullOrWhiteSpace($query)){
    $rawUpper = $query.Trim().ToUpper()
  }
  if([string]::IsNullOrWhiteSpace($normalizedInput) -and [string]::IsNullOrWhiteSpace($rawUpper)){
    return [pscustomobject]@{ NormalizedInput=$null; LinkValue=$null; Candidate=$null; Kind=$null; LookupKey=$null }
  }

  $scan = $null
  try { $scan = Normalize-Scan $query } catch {}
  $value = $null; $kind = $null
  $keys = New-Object System.Collections.ArrayList
  $addKey = {
    param($candidate)
    if([string]::IsNullOrWhiteSpace($candidate)){ return }
    $upper = $candidate.Trim().ToUpper()
    if([string]::IsNullOrWhiteSpace($upper)){ return }
    if(-not $keys.Contains($upper)){ [void]$keys.Add($upper) }
  }

  if($scan){
    $value = $scan.Value
    $kind  = $scan.Kind
    if(-not [string]::IsNullOrWhiteSpace($value)){
      if($scan.Kind -eq 'AssetTag'){
        foreach($variant in (Get-AssetKeyVariants $value)){ & $addKey $variant }
      } elseif($scan.Kind -eq 'Hostname'){
        foreach($variant in (HostnameKeyVariants $value)){ & $addKey $variant }
      } else {
        & $addKey $value
      }
    }
  }

  if([string]::IsNullOrWhiteSpace($value)){
    if(-not [string]::IsNullOrWhiteSpace($normalizedInput)){ $value = $normalizedInput }
    elseif(-not [string]::IsNullOrWhiteSpace($rawUpper)){ $value = $rawUpper }
  }

  if(-not [string]::IsNullOrWhiteSpace($normalizedInput)){ & $addKey $normalizedInput }
  if(-not [string]::IsNullOrWhiteSpace($rawUpper)){
    & $addKey $rawUpper
    $compactRaw = ($rawUpper -replace '[-\s]','')
    if($compactRaw -and ($compactRaw -ne $rawUpper)){ & $addKey $compactRaw }
  }

  $cand = $null
  $matchedKey = $null
  foreach($key in $keys){
    if($script:IndexByAsset.ContainsKey($key))      { $cand = $script:IndexByAsset[$key];      $matchedKey = $key; break }
    if($script:IndexBySerial.ContainsKey($key))     { $cand = $script:IndexBySerial[$key];     $matchedKey = $key; break }
    if($script:IndexByName.ContainsKey($key))       { $cand = $script:IndexByName[$key];       $matchedKey = $key; break }
  }

  if(-not $kind -and $cand){
    if(-not [string]::IsNullOrWhiteSpace($cand.name)){ $kind = 'Name' }
    elseif(-not [string]::IsNullOrWhiteSpace($cand.asset_tag)){ $kind = 'AssetTag' }
    elseif(-not [string]::IsNullOrWhiteSpace($cand.serial_number)){ $kind = 'Serial' }
  }

  return [pscustomobject]@{
    NormalizedInput = if($normalizedInput){ $normalizedInput } else { $rawUpper }
    LinkValue       = $value
    Candidate       = $cand
    Kind            = $kind
    LookupKey       = $matchedKey
  }
}
function Get-ProposedParentToken($cand,$parentRec){
  if(-not $cand){ return $null }
  if(-not $parentRec){ return $cand.u_parent_asset }
  if($cand.Type -eq 'Cart'){ return $parentRec.asset_tag }
  if(($cand.Type -eq 'Mic') -or ($cand.Type -eq 'Scanner')){
    $targetParent = $parentRec.asset_tag
    if($parentRec.name -match '^(?i)AO'){
      $carts = Find-CartsForComputer $parentRec
      if($carts.Count -gt 0){
        $cart = $carts[0]
        if($cart.asset_tag){ $targetParent = $cart.asset_tag }
        elseif($cart.name){ $targetParent = $cart.name }
      }
    }
    return $targetParent
  }
  return $parentRec.asset_tag
}
function Get-ParentDisplayName([string]$token,$fallbackRec=$null){
  if([string]::IsNullOrWhiteSpace($token)){ return '(none)' }
  $trimmed = $token.Trim()
  if($fallbackRec -and $fallbackRec.asset_tag){
    $fallbackKey = $fallbackRec.asset_tag.Trim().ToUpper()
    if($trimmed.Trim().ToUpper() -eq $fallbackKey -and -not [string]::IsNullOrWhiteSpace($fallbackRec.name)){
      return $fallbackRec.name
    }
  }
  $key = $trimmed.ToUpper()
  if($script:IndexByAsset.ContainsKey($key)){
    $parentRec = $script:IndexByAsset[$key]
    if($parentRec -and -not [string]::IsNullOrWhiteSpace($parentRec.name)){
      return $parentRec.name
    }
  }
  return $trimmed
}
function Get-PeripheralLinkPreview($cand,$parentRec){
  if(-not $cand){
    return [pscustomobject]@{
      ProposedName          = ''
      ProposedParentToken   = ''
      ProposedParentDisplay = ''
    }
  }
  $previewParentToken = Get-ProposedParentToken $cand $parentRec
  $proposedName = Compute-ProposedName $cand $parentRec
  $parentDisplay = Get-ParentDisplayName $previewParentToken $parentRec
  return [pscustomobject]@{
    ProposedName          = $proposedName
    ProposedParentToken   = $previewParentToken
    ProposedParentDisplay = if([string]::IsNullOrWhiteSpace($parentDisplay)){ '' } else { $parentDisplay }
  }
}
function Resolve-ParentAssetTag([string]$token){
  if([string]::IsNullOrWhiteSpace($token)){ return '' }
  $trimmed = $token.Trim()

  $assetKeys = New-Object System.Collections.ArrayList
  foreach($variant in (Get-AssetKeyVariants $trimmed)){
    if(-not $assetKeys.Contains($variant)){ [void]$assetKeys.Add($variant) }
  }
  $upper = $trimmed.ToUpper()
  if(-not $assetKeys.Contains($upper)){ [void]$assetKeys.Add($upper) }

  foreach($key in $assetKeys){
    if($script:IndexByAsset.ContainsKey($key)){
      $rec = $script:IndexByAsset[$key]
      if($rec -and -not [string]::IsNullOrWhiteSpace($rec.asset_tag)){
        return $rec.asset_tag
      }
    }
  }

  $nameKeys = New-Object System.Collections.ArrayList
  foreach($variant in (HostnameKeyVariants $trimmed)){
    if(-not $nameKeys.Contains($variant)){ [void]$nameKeys.Add($variant) }
  }
  if(-not $nameKeys.Contains($upper)){ [void]$nameKeys.Add($upper) }

  foreach($nameKey in $nameKeys){
    if($script:IndexByName.ContainsKey($nameKey)){
      $rec = $script:IndexByName[$nameKey]
      if($rec -and -not [string]::IsNullOrWhiteSpace($rec.asset_tag)){
        return $rec.asset_tag
      }
    }
  }

  return $trimmed
}

function Log-AssocChange([string]$action,[string]$deviceType,[string]$childAT,[string]$oldParent,[string]$newParent,[string]$oldName,[string]$newName){
  $out = if($script:OutputFolder){$script:OutputFolder}else{$script:DataFolder}
  if(-not $out){ return }
  $file = Join-Path $out 'CMDBUpdates.csv'
  $oldParentTag = Resolve-ParentAssetTag $oldParent
  $newParentTag = Resolve-ParentAssetTag $newParent
  $row = [pscustomobject]@{
    Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Action    = $action
    DeviceType= $deviceType
    AssetTag  = $childAT
    OldParent = $oldParentTag
    NewParent = $newParentTag
    OldName   = $oldName
    NewName   = $newName
  }
  if(Test-Path $file){ $row | Export-Csv -Path $file -NoTypeInformation -Append -Encoding UTF8 }
  else { $row | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8 }
}
function Link-Peripheral([string]$query,$parentRec,$lookupResult=$null){
  if(-not $parentRec){ return $false }
  if(-not $lookupResult){ $lookupResult = Resolve-PeripheralLookup $query }
  $cand = $null
  if($lookupResult){ $cand = $lookupResult.Candidate }
  if(-not $cand){ [System.Windows.Forms.MessageBox]::Show("Peripheral not found.","Link") | Out-Null; return $false }
  if($cand.Kind -ne 'Peripheral'){ [System.Windows.Forms.MessageBox]::Show("Selected item is not a peripheral.","Link") | Out-Null; return $false }
  $oldParent = $cand.u_parent_asset
  $oldName = $cand.name
  $targetParent = Get-ProposedParentToken $cand $parentRec
  $cand.u_parent_asset = $targetParent
  $newName = Compute-ProposedName $cand $parentRec
  if(-not [string]::IsNullOrWhiteSpace($newName)){ $cand.name = $newName }
  Build-Indices
  Log-AssocChange 'Link' (Get-DetectedType $cand) $cand.asset_tag $oldParent $cand.u_parent_asset $oldName $cand.name
  Refresh-AssocViews $parentRec
  Update-CartCheckbox-State $parentRec
  Update-ManualRoundButton   $parentRec
  Validate-ParentAndName $script:CurrentDisplay $script:CurrentParent
  Update-FixNameButton $script:CurrentDisplay $script:CurrentParent
  return $true
}
function Show-AddPeripheralDialog($parentRec){
  if(-not $parentRec){ return }
  $dialog = New-Object System.Windows.Forms.Form
  $dialog.Text = 'Add Peripheral (Name/Asset/Serial)'
  $dialog.StartPosition = 'CenterParent'
  $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $dialog.MaximizeBox = $false
  $dialog.MinimizeBox = $false
  $dialog.ShowIcon = $false
  $dialog.AutoSize = $true
  $dialog.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $dialog.Padding = New-Object System.Windows.Forms.Padding(12)

  $layout = New-Object System.Windows.Forms.TableLayoutPanel
  $layout.Dock = 'Fill'
  $layout.AutoSize = $true
  $layout.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $layout.ColumnCount = 1
  $layout.RowCount = 3
  $layout.RowStyles.Clear()
  $layout.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
  $layout.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
  $layout.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )

  # Input area
  $inputPanel = New-Object System.Windows.Forms.TableLayoutPanel
  $inputPanel.AutoSize = $true
  $inputPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $inputPanel.ColumnCount = 1
  $inputPanel.RowCount = 3
  $inputPanel.Dock = 'Top'
  $inputPanel.RowStyles.Clear()
  for($i=0;$i -lt 3;$i++){ [void]$inputPanel.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) }
  $lblSearch = New-Object System.Windows.Forms.Label
  $lblSearch.Text = 'Universal Search (name / asset / serial):'
  $lblSearch.AutoSize = $true
  $lblSearch.Margin = '0,0,0,2'
  $txtSearch = New-Object System.Windows.Forms.TextBox
  $txtSearch.Width = 260
  $txtSearch.Margin = '0,0,0,4'
  $lblHint = New-Object System.Windows.Forms.Label
  $lblHint.Text = 'Press Enter to search by cart name, asset tag, or serial number.'
  $lblHint.AutoSize = $true
  $lblHint.ForeColor = $script:ThemeColors.MutedText
  $lblHint.Margin = '0,0,0,0'
  $inputPanel.Controls.Add($lblSearch,0,0)
  $inputPanel.Controls.Add($txtSearch,0,1)
  $inputPanel.Controls.Add($lblHint,0,2)

  # Preview area
  $grpPreview = New-Object System.Windows.Forms.GroupBox
  $grpPreview.Text = 'Peripheral Preview'
  $grpPreview.AutoSize = $true
  $grpPreview.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $grpPreview.Dock = 'Top'
  $grpPreview.Padding = New-Object System.Windows.Forms.Padding(10)
  $tblPreview = New-Object System.Windows.Forms.TableLayoutPanel
  $tblPreview.AutoSize = $true
  $tblPreview.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $tblPreview.ColumnCount = 2
  $tblPreview.RowCount = 7
  $tblPreview.Dock = 'Top'
  $tblPreview.ColumnStyles.Clear()
  $tblPreview.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )
  $tblPreview.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )
  $tblPreview.RowStyles.Clear()
  for($i=0;$i -lt 7;$i++){ [void]$tblPreview.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) }

  $createLabel = {
    param($text)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.AutoSize = $true
    $lbl.Margin = '0,0,6,4'
    return $lbl
  }
  $createValue = {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $true
    $lbl.Margin = '0,0,0,4'
    return $lbl
  }

  $lblTypeVal = & $createValue
  $lblNameCurrentVal = & $createValue
  $lblNameProposedVal = & $createValue
  $lblParentCurrentVal = & $createValue
  $lblParentProposedVal = & $createValue
  $lblAssetVal = & $createValue
  $lblSerialVal = & $createValue
  $lblRITMVal = & $createValue
  $lblRetireVal = & $createValue

  $createArrow = {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $true
    $lbl.Margin = '6,0,6,4'
    $lbl.Text = '---->'
    $lbl.Visible = $false
    return $lbl
  }

  $lblNameArrow = & $createArrow
  $nameValuePanel = New-Object System.Windows.Forms.FlowLayoutPanel
  $nameValuePanel.AutoSize = $true
  $nameValuePanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $nameValuePanel.FlowDirection = 'LeftToRight'
  $nameValuePanel.WrapContents = $false
  $nameValuePanel.Margin = '0,0,0,0'
  $nameValuePanel.Padding = '0,0,0,0'
  $nameValuePanel.Controls.Add($lblNameCurrentVal)
  $nameValuePanel.Controls.Add($lblNameArrow)
  $nameValuePanel.Controls.Add($lblNameProposedVal)

  $lblParentArrow = & $createArrow
  $parentValuePanel = New-Object System.Windows.Forms.FlowLayoutPanel
  $parentValuePanel.AutoSize = $true
  $parentValuePanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $parentValuePanel.FlowDirection = 'LeftToRight'
  $parentValuePanel.WrapContents = $false
  $parentValuePanel.Margin = '0,0,0,0'
  $parentValuePanel.Padding = '0,0,0,0'
  $parentValuePanel.Controls.Add($lblParentCurrentVal)
  $parentValuePanel.Controls.Add($lblParentArrow)
  $parentValuePanel.Controls.Add($lblParentProposedVal)

  $tblPreview.Controls.Add((& $createLabel 'Type:'),0,0)
  $tblPreview.Controls.Add($lblTypeVal,1,0)
  $tblPreview.Controls.Add((& $createLabel 'Name:'),0,1)
  $tblPreview.Controls.Add($nameValuePanel,1,1)
  $tblPreview.Controls.Add((& $createLabel 'Parent:'),0,2)
  $tblPreview.Controls.Add($parentValuePanel,1,2)
  $tblPreview.Controls.Add((& $createLabel 'Asset Tag:'),0,3)
  $tblPreview.Controls.Add($lblAssetVal,1,3)
  $tblPreview.Controls.Add((& $createLabel 'Serial Number:'),0,4)
  $tblPreview.Controls.Add($lblSerialVal,1,4)
  $tblPreview.Controls.Add((& $createLabel 'RITM:'),0,5)
  $tblPreview.Controls.Add($lblRITMVal,1,5)
  $tblPreview.Controls.Add((& $createLabel 'Retire:'),0,6)
  $tblPreview.Controls.Add($lblRetireVal,1,6)

  $grpPreview.Controls.Add($tblPreview)

  # Buttons
  $buttonPanel = New-Object System.Windows.Forms.TableLayoutPanel
  $buttonPanel.AutoSize = $true
  $buttonPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $buttonPanel.ColumnCount = 3
  $buttonPanel.RowCount = 1
  $buttonPanel.Dock = 'Top'
  $buttonPanel.ColumnStyles.Clear()
  $buttonPanel.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)) )
  $buttonPanel.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )
  $buttonPanel.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )
  $buttonPanel.RowStyles.Clear()
  $buttonPanel.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
  $btnDialogAdd = New-Object ModernUI.RoundedButton
  $btnDialogAdd.Text = 'Add'
  $btnDialogAdd.Enabled = $false
  $btnDialogAdd.Margin = '0,8,6,0'
  $btnDialogCancel = New-Object ModernUI.RoundedButton
  $btnDialogCancel.Text = 'Cancel'
  $btnDialogCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  $btnDialogCancel.Margin = '0,8,0,0'
  $spacer = New-Object System.Windows.Forms.Panel
  $spacer.Dock = 'Fill'
  $buttonPanel.Controls.Add($spacer,0,0)
  $buttonPanel.Controls.Add($btnDialogAdd,1,0)
  $buttonPanel.Controls.Add($btnDialogCancel,2,0)
  $dialog.CancelButton = $btnDialogCancel

  $layout.Controls.Add($inputPanel,0,0)
  $layout.Controls.Add($grpPreview,0,1)
  $layout.Controls.Add($buttonPanel,0,2)
  $dialog.Controls.Add($layout)

  $valueLabels = @($lblTypeVal,$lblNameCurrentVal,$lblNameProposedVal,$lblParentCurrentVal,$lblParentProposedVal,$lblAssetVal,$lblSerialVal,$lblRITMVal,$lblRetireVal)
  $lookupResult = $null
  $clearPreview = {
    foreach($lbl in $valueLabels){ $lbl.Text = '' }
    $lblNameArrow.Visible = $false
    $lblParentArrow.Visible = $false
    $btnDialogAdd.Enabled = $false
  }
  $applyPreview = {
    param($result)
    $lookupResult = $result
    & $clearPreview
    if(-not $result){ return }
    $cand = $result.Candidate
    if(-not $cand){
      if($result.NormalizedInput){ $lblTypeVal.Text = '(not found)' }
      return
    }
    $lblTypeVal.Text = if($cand.Type){ $cand.Type } else { $cand.Kind }
    $lblNameCurrentVal.Text = if($cand.name){ $cand.name } else { '' }
    $lblAssetVal.Text = $cand.asset_tag
    $lblSerialVal.Text = $cand.serial_number
    $lblRITMVal.Text = $cand.RITM
    $lblRetireVal.Text = Fmt-DateLong $cand.Retire
    $lblParentCurrentVal.Text = Get-ParentDisplayName $cand.u_parent_asset
    $preview = Get-PeripheralLinkPreview $cand $parentRec
    $proposedName = ''
    $proposedParentDisplay = ''
    if($preview){
      if(-not [string]::IsNullOrWhiteSpace($preview.ProposedName)){
        $proposedName = $preview.ProposedName
      }
      if(-not [string]::IsNullOrWhiteSpace($preview.ProposedParentDisplay)){
        $proposedParentDisplay = $preview.ProposedParentDisplay
      }
    }
    if([string]::IsNullOrWhiteSpace($proposedName)){
      $proposedName = $lblNameCurrentVal.Text
    }
    $lblNameProposedVal.Text = $proposedName
    $lblNameArrow.Visible = -not [string]::IsNullOrWhiteSpace($proposedName)
    if([string]::IsNullOrWhiteSpace($proposedParentDisplay)){
      $proposedParentDisplay = $lblParentCurrentVal.Text
    }
    $lblParentProposedVal.Text = $proposedParentDisplay
    $lblParentArrow.Visible = -not [string]::IsNullOrWhiteSpace($proposedParentDisplay)
    if($cand.Kind -eq 'Peripheral'){ $btnDialogAdd.Enabled = $true }
  }

  $txtSearch.Add_TextChanged({
    $lookupResult = $null
    & $clearPreview
  })
  $txtSearch.Add_KeyDown({
    if($_.KeyCode -eq 'Enter'){
      $_.SuppressKeyPress = $true
      $result = Resolve-PeripheralLookup $txtSearch.Text
      & $applyPreview $result
    }
  })
  $btnDialogAdd.Add_Click({
    $result = $lookupResult
    if(-not $result){
      $result = Resolve-PeripheralLookup $txtSearch.Text
      & $applyPreview $result
    }
    if(-not $result){ return }
    if(Link-Peripheral $txtSearch.Text $parentRec $result){
      $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
      $dialog.Close()
    }
  })
  $dialog.Add_Shown({ $txtSearch.Focus() })
  Apply-ModernThemeToForm -Form $dialog
  try { [void]$dialog.ShowDialog($form) } finally { try { $dialog.Dispose() } catch {} }
}
function Remove-Selected-Associations($parentRec){
  if($dgv.SelectedRows.Count -eq 0){ return }
  foreach($row in $dgv.SelectedRows){
    $role = [string]$row.Cells['Role'].Value
    if(($role -ne 'Child') -and ($role -ne 'Grandchild')){ continue }
    $asset = [string]$row.Cells['AssetTag'].Value
    if([string]::IsNullOrWhiteSpace($asset)){ continue }
    $key=$asset.ToUpper()
    if(-not $script:IndexByAsset.ContainsKey($key)){ continue }
    $ch=$script:IndexByAsset[$key]
    $oldParent = $ch.u_parent_asset
    $oldName = $ch.name
    $ch.u_parent_asset = $null
    if($ch.serial_number){ $ch.name = $ch.serial_number }
    if($script:ChildrenByParent.ContainsKey($parentRec.asset_tag)){ [void]$script:ChildrenByParent[$parentRec.asset_tag].Remove($ch) }
    Log-AssocChange 'Unlink' (Get-DetectedType $ch) $ch.asset_tag $oldParent $null $oldName $ch.name
  }
  Build-Indices
  Refresh-AssocViews $parentRec
  Update-CartCheckbox-State $parentRec
  Update-ManualRoundButton $parentRec
  Validate-ParentAndName $script:CurrentDisplay $script:CurrentParent
  Update-FixNameButton $script:CurrentDisplay $script:CurrentParent
}
function Update-FixNameButton([object]$displayRec = $null, [object]$parentRec = $null){
  if(-not $displayRec){ $displayRec = $script:CurrentDisplay }
  if(-not $parentRec){  $parentRec  = $script:CurrentParent  }
  if(-not $displayRec -or -not $parentRec){ $btnFixName.Enabled = $false; return }
  if($displayRec.Type -eq 'Computer'){ $btnFixName.Enabled = $false; return }
  $expected = Compute-ProposedName $displayRec $parentRec
  if([string]::IsNullOrWhiteSpace($expected)){ $btnFixName.Enabled = $false; return }
  $cur = ''; if($displayRec.name){ $cur = $displayRec.name.Trim().ToUpper() }
  if($cur -ne $expected.Trim().ToUpper()){ $btnFixName.Enabled = $true } else { $btnFixName.Enabled = $false }
}
function Fix-DisplayName(){
  $disp   = $script:CurrentDisplay
  $parent = $script:CurrentParent
  if(-not $disp){ return }
  if(-not $parent){ [System.Windows.Forms.MessageBox]::Show("No parent computer found for this device. Scan a device with a valid parent first.","Fix Name") | Out-Null; return }
  if($disp.Type -eq 'Computer'){ [System.Windows.Forms.MessageBox]::Show("Fix Name applies to peripherals only.","Fix Name") | Out-Null; return }
  $expected = Compute-ProposedName $disp $parent
  if([string]::IsNullOrWhiteSpace($expected)){ [System.Windows.Forms.MessageBox]::Show("Could not compute the expected name for this device type.","Fix Name") | Out-Null; return }
  if($disp.name -and ($disp.name.Trim().ToUpper() -eq $expected.Trim().ToUpper())){
    [System.Windows.Forms.MessageBox]::Show("Name already matches the expected convention.","Fix Name") | Out-Null
    Update-FixNameButton $disp $parent
    return
  }
  $oldParent = if($disp.PSObject.Properties['u_parent_asset']) { $disp.u_parent_asset } else { $parent.asset_tag }
  $oldName   = $disp.name
  $disp.name = $expected
  Log-AssocChange 'Rename' (Get-DetectedType $disp) $disp.asset_tag $oldParent $oldParent $oldName $disp.name
  Build-Indices
  if($parent){ Refresh-AssocViews $parent }
  $txtHost.Text = $disp.name
  Validate-ParentAndName $disp $parent
  Update-FixNameButton $disp $parent
  $statusLabel.Text = ("Renamed to '" + $expected + "'.")
}
# ---- Populate Summary/UI ----
function Populate-UI($displayRec,$parentRec){
  try { Populate-Department-Combo $displayRec.u_department_location } catch {}
  try { Update-MaintenanceTypeSelection $displayRec $parentRec } catch {}
  $script:CurrentDisplay = $displayRec
  $script:CurrentParent  = $parentRec
  $txtType.Text = Get-DetectedType $displayRec
  $txtHost.Text=$displayRec.name
  $txtAT.Text=$displayRec.asset_tag
  $txtSN.Text=$displayRec.serial_number
   if($displayRec.Type -eq 'Computer'){
    $txtParent.Text='(n/a)'
  } else {
    if($displayRec.PSObject.Properties['u_parent_asset'] -and $displayRec.u_parent_asset){
      $txtParent.Text=$displayRec.u_parent_asset
    } else {
      $txtParent.Text='(blank)'
    }
  }
  $txtRITM.Text=$displayRec.RITM
  $txtRetire.Text = Fmt-DateLong $displayRec.Retire
  Show-RoundingStatus $parentRec
  if($parentRec){ Validate-Location $parentRec } else { Validate-Location $displayRec }
  if($parentRec){ Refresh-AssocViews $parentRec }
  Update-CartCheckbox-State $parentRec
  Update-ManualRoundButton   $parentRec
  if($btnAddPeripheral){ Set-SearchTextButtonBaseState -Button $btnAddPeripheral -BaseEnabled ([bool]$parentRec) }
  Validate-ParentAndName $displayRec $parentRec
  Update-FixNameButton $displayRec $parentRec
}
# ---- Location cascading (City > Location > Building > Floor > Room) ----
$script:IsPopulatingLocationCombos = $false

function Filter-LocationRows {
  param(
    [string]$city,
    [string]$location,
    [string]$building,
    [string]$floor
  )
  $rows = $script:LocationRows
  if($city){
    $nCity = Normalize-Field $city
    $rows = $rows | Where-Object { (Normalize-Field (Get-LocVal $_ 'City')) -eq $nCity }
  }
  if($location){
    $nLocation = Normalize-Field $location
    $rows = $rows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Location')) -eq $nLocation }
  }
  if($building){
    $nBuilding = Normalize-Field $building
    $rows = $rows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Building')) -eq $nBuilding }
  }
  if($floor){
    $nFloor = Normalize-Field $floor
    $rows = $rows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Floor')) -eq $nFloor }
  }
  return @($rows)
}

function Get-ValidLocationSelection {
  param(
    [string]$value,
    [object[]]$options
  )
  if(-not $value){ return $null }
  $norm = Normalize-Field $value
  foreach($opt in $options){
    if($null -eq $opt){ continue }
    $optText = '' + $opt
    if((Normalize-Field $optText) -eq $norm){ return $optText }
  }
  return $null
}

function Set-ComboSelectedIndexForText {
  param(
    [System.Windows.Forms.ComboBox]$Combo,
    [string]$Text
  )
  if(-not $Combo){ return }
  if([string]::IsNullOrEmpty($Text)){
    try {
      if($Combo.SelectedIndex -ne -1){ $Combo.SelectedIndex = -1 }
    } catch {}
    return
  }
  try {
    $index = $Combo.FindStringExact($Text)
    if($index -ge 0 -and $Combo.SelectedIndex -ne $index){
      $Combo.SelectedIndex = $index
    }
  } catch {}
}

function Restore-ComboSelectionLater {
  param(
    [System.Windows.Forms.ComboBox]$Combo,
    [object]$SelectionStart,
    [object]$SelectionLength
  )
  if(-not $Combo){ return }
  if($SelectionStart -eq $null){ return }
  $apply = {
    param($ctrl, $start, $length)
    try {
      $resolvedStart = [math]::Min($start, $ctrl.Text.Length)
      $ctrl.SelectionStart = $resolvedStart
      if($length -ne $null){
        $ctrl.SelectionLength = [math]::Min($length, $ctrl.Text.Length - $resolvedStart)
      }
    } catch {}
  }
  & $apply $Combo $SelectionStart $SelectionLength
  try {
    $null = $Combo.BeginInvoke($apply, @($Combo, $SelectionStart, $SelectionLength))
  } catch {}
}

function Set-ComboTextIfChanged {
  param(
    [System.Windows.Forms.ComboBox]$Combo,
    [string]$Text
  )
  if(-not $Combo){ return }
  $target = if($Text){ $Text } else { '' }
  $current = if($Combo.Text){ $Combo.Text } else { '' }
  if($current -eq $target){
    Set-ComboSelectedIndexForText $Combo $target
    return
  }
  $selStart = $null
  $selLength = $null
  try {
    $selStart = $Combo.SelectionStart
    $selLength = $Combo.SelectionLength
  } catch {}
  $Combo.Text = $target
  Set-ComboSelectedIndexForText $Combo $target
  Restore-ComboSelectionLater $Combo $selStart $selLength
}

function Populate-Location-Combos {
  param(
    [string]$city,
    [string]$loc,
    [string]$b,
    [string]$f,
    [string]$r,
    [switch]$PreserveInvalidSelections,
    [ValidateSet('City','Location','Building','Floor','Room')]
    [string]$ChangedLevel
  )
  $script:IsPopulatingLocationCombos = $true
  try {
    $cmbCity.Items.Clear(); $cmbLocation.Items.Clear(); $cmbBuilding.Items.Clear(); $cmbFloor.Items.Clear(); $cmbRoom.Items.Clear()

    $origLoc = $loc
    $origBuilding = $b
    $origFloor = $f
    $origRoom = $r

    $clearLocation = $false
    $clearBuilding = $false
    $clearFloor = $false
    $clearRoom = $false

    # City
    $cities = $script:LocationRows | ForEach-Object { Get-LocVal $_ 'City' } | Where-Object { $_ } | Select-Object -Unique | Sort-Object
    $cmbCity.Items.AddRange(@($cities))
    $validCityInput = Get-ValidLocationSelection $city $cities
    if($ChangedLevel -eq 'City'){
      $prevCityNorm = $script:LastLocationSelections.City
      if([string]::IsNullOrWhiteSpace($city)){
        if($prevCityNorm -ne ''){ $clearLocation = $true; $clearBuilding = $true; $clearFloor = $true; $clearRoom = $true }
      } elseif($validCityInput){
        $validCityNorm = Normalize-LocationComparisonValue $validCityInput
        if($validCityNorm -ne $prevCityNorm){ $clearLocation = $true; $clearBuilding = $true; $clearFloor = $true; $clearRoom = $true }
      }
    }
    Set-ComboTextIfChanged $cmbCity ($city)
    $validCity = Get-ValidLocationSelection $cmbCity.Text $cities
    if(-not $validCity -and -not $PreserveInvalidSelections){ Set-ComboTextIfChanged $cmbCity '' }
    $filterCity = if($validCity){ $validCity } else { $null }
    if($clearLocation){
      $loc = $null
      $clearBuilding = $true
      $clearFloor = $true
      $clearRoom = $true
    }

    # Location (filtered by City if present)
    $locRows = Filter-LocationRows $filterCity $null $null $null
    $locs = $locRows | ForEach-Object { Get-LocVal $_ 'Location' } | Where-Object { $_ } | Select-Object -Unique | Sort-Object
    $validLocationInput = Get-ValidLocationSelection $origLoc $locs
    if($ChangedLevel -eq 'Location'){
      $prevLocationNorm = $script:LastLocationSelections.Location
      if([string]::IsNullOrWhiteSpace($origLoc)){
        if($prevLocationNorm -ne ''){ $clearBuilding = $true; $clearFloor = $true; $clearRoom = $true }
      } elseif($validLocationInput){
        $validLocationNorm = Normalize-LocationComparisonValue $validLocationInput
        if($validLocationNorm -ne $prevLocationNorm){ $clearBuilding = $true; $clearFloor = $true; $clearRoom = $true }
      }
    }
    if($clearBuilding){
      $b = $null
      $clearFloor = $true
      $clearRoom = $true
    }
    $cmbLocation.Items.AddRange(@($locs))
    Set-ComboTextIfChanged $cmbLocation ($loc)
    $validLocation = Get-ValidLocationSelection $cmbLocation.Text $locs
    if(-not $validLocation -and -not $PreserveInvalidSelections){ Set-ComboTextIfChanged $cmbLocation '' }
    $filterLocation = if($validLocation){ $validLocation } else { $null }

    # Building
    $blds = @()
    if($filterLocation){
      $bldRows = Filter-LocationRows $filterCity $filterLocation $null $null
      $blds = $bldRows | ForEach-Object { Get-LocVal $_ 'Building' } | Where-Object { $_ } | Select-Object -Unique | Sort-Object
    }
    $validBuildingInput = Get-ValidLocationSelection $origBuilding $blds
    if($ChangedLevel -eq 'Building'){
      $prevBuildingNorm = $script:LastLocationSelections.Building
      if([string]::IsNullOrWhiteSpace($origBuilding)){
        if($prevBuildingNorm -ne ''){ $clearFloor = $true; $clearRoom = $true }
      } elseif($validBuildingInput){
        $validBuildingNorm = Normalize-LocationComparisonValue $validBuildingInput
        if($validBuildingNorm -ne $prevBuildingNorm){ $clearFloor = $true; $clearRoom = $true }
      }
    }
    if($clearFloor){
      $f = $null
      $clearRoom = $true
    }
    $cmbBuilding.Items.AddRange(@($blds))
    Set-ComboTextIfChanged $cmbBuilding ($b)
    $validBuilding = Get-ValidLocationSelection $cmbBuilding.Text $blds
    if(-not $validBuilding -and -not $PreserveInvalidSelections){ Set-ComboTextIfChanged $cmbBuilding '' }
    $filterBuilding = if($validBuilding){ $validBuilding } else { $null }

    # Floor
    $floors = @()
    if($filterLocation -and $filterBuilding){
      $floorRows = Filter-LocationRows $filterCity $filterLocation $filterBuilding $null
      $floors = $floorRows | ForEach-Object { Get-LocVal $_ 'Floor' } | Where-Object { $_ } | Select-Object -Unique
      $floors = Sort-Floors $floors
    }
    $validFloorInput = Get-ValidLocationSelection $origFloor $floors
    if($ChangedLevel -eq 'Floor'){
      $prevFloorNorm = $script:LastLocationSelections.Floor
      if([string]::IsNullOrWhiteSpace($origFloor)){
        if($prevFloorNorm -ne ''){ $clearRoom = $true }
      } elseif($validFloorInput){
        $validFloorNorm = Normalize-LocationComparisonValue $validFloorInput
        if($validFloorNorm -ne $prevFloorNorm){ $clearRoom = $true }
      }
    }
    if($clearRoom){ $r = $null }
    $cmbFloor.Items.AddRange(@($floors))
    Set-ComboTextIfChanged $cmbFloor ($f)
    $validFloor = Get-ValidLocationSelection $cmbFloor.Text $floors
    if(-not $validFloor -and -not $PreserveInvalidSelections){ Set-ComboTextIfChanged $cmbFloor '' }
    $filterFloor = if($validFloor){ $validFloor } else { $null }

    # Room
    $rooms = @()
    if($filterLocation -and $filterBuilding -and $filterFloor){
      $roomRows = Filter-LocationRows $filterCity $filterLocation $filterBuilding $filterFloor
      $rooms = $roomRows | ForEach-Object { Get-LocVal $_ 'Room' } | Where-Object { $_ } | Select-Object -Unique | Sort-Object
    }
    $cmbRoom.Items.AddRange(@($rooms))
    Set-ComboTextIfChanged $cmbRoom ($r)
    $validRoom = Get-ValidLocationSelection $cmbRoom.Text $rooms
    if(-not $validRoom -and -not $PreserveInvalidSelections){ Set-ComboTextIfChanged $cmbRoom '' }
  }
  finally {
    $script:IsPopulatingLocationCombos = $false
  }
}

function Should-SkipExcludedDevice {
  param([object]$device)
  if(-not $device){ return $false }
  try {
    if($chkShowExcluded -and -not $chkShowExcluded.Checked){
      $mt = ('' + $device.u_device_rounding).Trim()
      if($mt -match '^(?i)Excluded$'){ return $true }
    }
  } catch {}
  return $false
}
function Toggle-EditLocation(){
  $script:editing = -not $script:editing
  if($script:editing){
    Update-LastLocationSelections $txtCity.Text $txtLocation.Text $txtBldg.Text $txtFloor.Text $txtRoom.Text
    Populate-Location-Combos $txtCity.Text $txtLocation.Text $txtBldg.Text $txtFloor.Text $txtRoom.Text -PreserveInvalidSelections
    try { Populate-Department-Combo ($txtDept.Text) } catch {}
    $cmbCity.Visible=$true; $cmbLocation.Visible=$true; $cmbBuilding.Visible=$true; $cmbFloor.Visible=$true; $cmbRoom.Visible=$true
    $txtCity.Visible=$false; $txtLocation.Visible=$false; $txtBldg.Visible=$false; $txtFloor.Visible=$false; $txtRoom.Visible=$false
    if ($cmbDept) { $cmbDept.Visible=$true } ; if ($txtDept) { $txtDept.Visible=$false }
    foreach($t in @($txtDept,$txtDepartment)){ if ($t) { $t.Visible = $false } }
    foreach($c in @($cmbDept,$cmbDepartment,$ddlDept,$ddlDepartment)){ if ($c) { $c.Visible = $true } }
    $btnEditLoc.Text="Save Location"
  } else {
    $city = if($cmbCity.Text){ $cmbCity.Text.Trim() } else { '' }
    $loc  = if($cmbLocation.Text){ $cmbLocation.Text.Trim() } else { '' }
    $b    = if($cmbBuilding.Text){ $cmbBuilding.Text.Trim() } else { '' }
    $f    = if($cmbFloor.Text){ $cmbFloor.Text.Trim() } else { '' }
    $r    = if($cmbRoom.Text){ $cmbRoom.Text.Trim() } else { '' }
    $dept = ''
    if($cmbDept -and $cmbDept.Text){ $dept = $cmbDept.Text }
    elseif($txtDept -and $txtDept.Text){ $dept = $txtDept.Text }
    $hasCityAndLocation = (-not [string]::IsNullOrWhiteSpace($city)) -and (-not [string]::IsNullOrWhiteSpace($loc))
    $exists = $script:LocationRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'City')) -eq (Normalize-Field $city) -and (Normalize-Field (Get-LocVal $_ 'Location')) -eq (Normalize-Field $loc) -and (Normalize-Field (Get-LocVal $_ 'Building')) -eq (Normalize-Field $b) -and (Normalize-Field (Get-LocVal $_ 'Floor')) -eq (Normalize-Field $f) -and (Normalize-Field (Get-LocVal $_ 'Room')) -eq (Normalize-Field $r) }
    if($exists.Count -eq 0 -and $hasCityAndLocation){
      Save-LocationUserAdd $city $loc $b $f $r
    }
    $txtCity.Text=$city; $txtLocation.Text=$loc; $txtBldg.Text=$b; $txtFloor.Text=$f; $txtRoom.Text=$r
    Update-LastLocationSelections $city $loc $b $f $r
    if($dept){
      try { Save-DepartmentUserAdd $dept } catch {}
      try { Populate-Department-Combo $dept } catch {}
    }
    if($txtDept){ $txtDept.Text = $dept }
    if($cmbDept){ $cmbDept.Text = $dept }
    $targets = New-Object System.Collections.ArrayList
    foreach($cand in @($script:CurrentParent,$script:CurrentDisplay)){
      if(-not $cand){ continue }
      if(-not $targets.Contains($cand)){ [void]$targets.Add($cand) }
    }
    foreach($tgt in $targets){
      try { $tgt | Add-Member -NotePropertyName location -NotePropertyValue $loc -Force } catch {}
      try { $tgt | Add-Member -NotePropertyName u_building -NotePropertyValue $b -Force } catch {}
      try { $tgt | Add-Member -NotePropertyName u_floor -NotePropertyValue $f -Force } catch {}
      try { $tgt | Add-Member -NotePropertyName u_room -NotePropertyValue $r -Force } catch {}
      try { $tgt | Add-Member -NotePropertyName u_department_location -NotePropertyValue $dept -Force } catch {}
      try { $tgt | Add-Member -NotePropertyName Department -NotePropertyValue $dept -Force } catch {}
    }
    $script:editing = $false
    if($targets.Count -gt 0){
      Validate-Location $targets[0]
    } else {
      $tmp=[pscustomobject]@{location=$loc;u_building=$b;u_floor=$f;u_room=$r;u_department_location=$dept;Department=$dept}
      Validate-Location $tmp
    }
    $cmbCity.Visible=$false; $cmbLocation.Visible=$false; $cmbBuilding.Visible=$false; $cmbFloor.Visible=$false; $cmbRoom.Visible=$false
    $txtCity.Visible=$true; $txtLocation.Visible=$true; $txtBldg.Visible=$true; $txtFloor.Visible=$true; $txtRoom.Visible=$true
    if ($cmbDept) { $cmbDept.Visible=$false }
    if ($txtDept) { $txtDept.Visible=$true }
    $btnEditLoc.Text="Edit Location"
  }
}
$btnRemove.Add_Click({
  $pc = $script:CurrentParent
  if(-not $pc){ return }
  if(Should-SkipExcludedDevice $pc){ return }
  Remove-Selected-Associations $pc
})
$cmbCity.Add_TextChanged({
  if($script:IsPopulatingLocationCombos){ return }
  $selStart = $null; $selLength = $null
  try { $selStart = $cmbCity.SelectionStart; $selLength = $cmbCity.SelectionLength } catch {}
  Populate-Location-Combos $cmbCity.Text $cmbLocation.Text $cmbBuilding.Text $cmbFloor.Text $cmbRoom.Text -ChangedLevel 'City' -PreserveInvalidSelections
  Restore-ComboSelectionLater $cmbCity $selStart $selLength
})
$cmbLocation.Add_TextChanged({
  if($script:IsPopulatingLocationCombos){ return }
  $selStart = $null; $selLength = $null
  try { $selStart = $cmbLocation.SelectionStart; $selLength = $cmbLocation.SelectionLength } catch {}
  Populate-Location-Combos $cmbCity.Text $cmbLocation.Text $cmbBuilding.Text $cmbFloor.Text $cmbRoom.Text -ChangedLevel 'Location' -PreserveInvalidSelections
  Restore-ComboSelectionLater $cmbLocation $selStart $selLength
})
$cmbBuilding.Add_TextChanged({
  if($script:IsPopulatingLocationCombos){ return }
  $selStart = $null; $selLength = $null
  try { $selStart = $cmbBuilding.SelectionStart; $selLength = $cmbBuilding.SelectionLength } catch {}
  Populate-Location-Combos $cmbCity.Text $cmbLocation.Text $cmbBuilding.Text $cmbFloor.Text $cmbRoom.Text -ChangedLevel 'Building' -PreserveInvalidSelections
  Restore-ComboSelectionLater $cmbBuilding $selStart $selLength
})
$cmbFloor.Add_TextChanged({
  if($script:IsPopulatingLocationCombos){ return }
  $selStart = $null; $selLength = $null
  try { $selStart = $cmbFloor.SelectionStart; $selLength = $cmbFloor.SelectionLength } catch {}
  Populate-Location-Combos $cmbCity.Text $cmbLocation.Text $cmbBuilding.Text $cmbFloor.Text $cmbRoom.Text -ChangedLevel 'Floor' -PreserveInvalidSelections
  Restore-ComboSelectionLater $cmbFloor $selStart $selLength
})
# ---- Actions ----
function Focus-ScanInput(){
  if ($script:NewAssetToolSearchTextBox) {
    try { [void]$script:NewAssetToolSearchTextBox.Focus() } catch {}
  } elseif ($txtScan) {
    try { $txtScan.Focus() } catch {}
  }
}
function Do-Lookup(){
  Stop-RoundingTimer
  $raw = Find-RecordRaw $txtScan.Text
  if(-not $raw){
    Reset-RoundingTimer
    $statusLabel.Text=("No match for '" + $txtScan.Text + "'")
    return
  }
  $parent = Resolve-ParentComputer $raw
  Populate-UI $raw $parent
  Start-RoundingTimer
  $statusLabel.Text=("Found " + $raw.Kind + " / " + $raw.Type)
}
function Clear-UI(){
  Reset-RoundingTimer
  $script:CurrentDisplay = $null; $script:CurrentParent  = $null
  foreach($tb in @($txtType,$txtHost,$txtAT,$txtSN,$txtParent,$txtRITM,$txtRetire,$txtRound,$txtCity,$txtLocation,$txtBldg,$txtFloor,$txtRoom,$txtDept,$txtDepartment,$txtComments)){
    if($tb -is [System.Windows.Forms.Control]){
      $tb.Text = ''
      $tb.BackColor = [System.Drawing.Color]::White
    }
  }
  foreach($combo in @($cmbDept,$cmbDepartment,$ddlDept,$ddlDepartment)){
    try {
      if($combo){
        $combo.SelectedIndex = -1
        $combo.Text = ''
      }
    } catch {}
  }
  Update-LastLocationSelections '' '' '' '' ''
  if($cmbMaintType){
    try {
      $cmbMaintType.SelectedIndex = -1
      $cmbMaintType.Text = ''
    } catch {}
  }
  try { $dgv.Rows.Clear() } catch {}
  try { $cards.Controls.Clear() } catch {}
  Update-ManualRoundButton $null; Update-CartCheckbox-State $null
  foreach($cb in @($chkCable,$chkCableNeeded,$chkLabels,$chkCart,$chkPeriph)){ $cb.Checked=$false }
  $btnFixName.Enabled = $false
  if($btnAddPeripheral){ Set-SearchTextButtonBaseState -Button $btnAddPeripheral -BaseEnabled $false }
  $statusLabel.Text = "Ready - scan or enter a device."
  Size-AssocForRows(1) | Out-Null
}
# ---- Events ----
$txtScan.Add_KeyDown({ if($_.KeyCode -eq 'Enter'){ Do-Lookup; $_.SuppressKeyPress=$true } })
$txtScan.Add_TextChanged({
  if([string]::IsNullOrWhiteSpace($txtScan.Text)){ Clear-UI }
  try {
    if(Get-Command Update-RoundNowButtonState -ErrorAction SilentlyContinue){ Update-RoundNowButtonState }
  } catch {}
})
$btnEditLoc.Add_Click({ Toggle-EditLocation })
$btnAddPeripheral.Add_Click({
  $pc = $script:CurrentParent
  if(-not $pc){ return }
  try{
    if(-not $chkShowExcluded.Checked){
      $mt = ('' + $pc.u_device_rounding).Trim()
      if($mt -match '^(?i)Excluded$'){ return }
    }
  } catch {}
  Show-AddPeripheralDialog $pc
})
# Double-click a grid row to open that record
$dgv.Add_CellDoubleClick({
  if($_.RowIndex -lt 0){ return }
  $row = $dgv.Rows[$_.RowIndex]
  $asset = [string]$row.Cells['AssetTag'].Value
  $serial= [string]$row.Cells['Serial'].Value
  $name  = [string]$row.Cells['Name'].Value
  $rec = $null
  if($asset){ $key=$asset.Trim().ToUpper(); if($script:IndexByAsset.ContainsKey($key)){ $rec=$script:IndexByAsset[$key] } }
  if(-not $rec -and $serial){ $key=$serial.Trim().ToUpper(); if($script:IndexBySerial.ContainsKey($key)){ $rec=$script:IndexBySerial[$key] } }
  if(-not $rec -and $name){
    foreach($k in (HostnameKeyVariants $name)){ if($script:IndexByName.ContainsKey($k)){ $rec=$script:IndexByName[$k]; break } }
  }
  if($rec){
    $par=Resolve-ParentComputer $rec
    Populate-UI $rec $par
    Start-RoundingTimer
  }
})
$btnCheckComplete.Add_Click({
  $checkboxes = @($chkCable,$chkLabels,$chkPeriph,$chkCart)
  $enabledBoxes = $checkboxes | Where-Object { $_.Enabled }
  $allEnabledChecked = $false
  if($enabledBoxes.Count -gt 0){
    $allEnabledChecked = (($enabledBoxes | Where-Object { -not $_.Checked }).Count -eq 0)
  }
  if($allEnabledChecked){
    foreach($cb in $checkboxes){ $cb.Checked = $false }
    Update-CartCheckbox-State $script:CurrentParent
  } else {
    foreach($cb in $checkboxes){
      if($cb.Enabled){ $cb.Checked = $true }
    }
    $pc = $script:CurrentParent
    if($pc -and $pc.name -match '^(?i)AO'){
      $chkCart.Enabled = $true
      $chkCart.Checked = $true
    }
  }
})
$btnSave.Add_Click({
  Stop-RoundingTimer
  $out = $script:OutputFolder
  if(-not (Test-Path $out)){ New-Item -ItemType Directory -Path $out -Force | Out-Null }
$file = Join-Path ($(if($script:OutputFolder){$script:OutputFolder}else{$script:DataFolder})) 'RoundingEvents.csv'
  $exists = Test-Path $file
  if($exists){ Ensure-RoundingCommentsColumn $file }
  $pc = $script:CurrentParent
  if(-not $pc){ $pc = Resolve-ParentComputer (Find-RecordRaw $txtAT.Text) }
  if(-not $pc){ $pc = $script:CurrentDisplay }
  $shouldSkip = $false
  if($pc){ $shouldSkip = Should-SkipExcludedDevice $pc }
  if($shouldSkip){
    [System.Windows.Forms.MessageBox]::Show("This device is marked as Excluded. Enable 'Excluded' to log rounding.","Save Event") | Out-Null
    Focus-ScanInput
    return
  }
  $url = if($pc){ Get-RoundingUrlForParent $pc } else { $null }
  $deptValue = ''
  if($cmbDept -and $cmbDept.Text){ $deptValue = $cmbDept.Text }
  elseif($txtDept -and $txtDept.Text){ $deptValue = $txtDept.Text }
  if($deptValue){
    try { Save-DepartmentUserAdd $deptValue } catch {}
    if($txtDept){ $txtDept.Text = $deptValue }
    if($cmbDept){ $cmbDept.Text = $deptValue }
    try { Populate-Department-Combo $deptValue } catch {}
  }
  $row = [pscustomobject]@{
    Timestamp        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    AssetTag         = if($pc){ $pc.asset_tag } else { $null }
    Name             = if($pc){ $pc.name } else { $null }
    Serial           = if($pc){ $pc.serial_number } else { $null }
    City             = $txtCity.Text
    Location         = $txtLocation.Text
    Building         = $txtBldg.Text
    Floor            = $txtFloor.Text
    Room             = $txtRoom.Text
    CheckStatus      = $cmbChkStatus.Text
    RoundingMinutes  = [int]$numTime.Value
    CableMgmtOK      = if ($chkCable.Checked) { 'Yes' } else { 'No' }
    CablingNeeded    = if ($chkCableNeeded.Checked) { 'Yes' } else { 'No' }
    LabelOK          = if ($chkLabels.Checked) { 'Yes' } else { 'No' }
    CartOK           = if ($chkCart.Checked) { 'Yes' } else { 'No' }
    PeripheralsOK    = if ($chkPeriph.Checked) { 'Yes' } else { 'No' }
    MaintenanceType = $cmbMaintType.Text
    Department       = $deptValue
    RoundingUrl      = $url
    Comments         = $txtComments.Text
    Rounded          = 'No'
  }
  $cmbDept.Visible = $false  # Hidden until Edit Location is active
  $rowOut = $row | Select-Object $script:RoundingEventColumns
  if(-not $exists){ $rowOut | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8 }
  else { $rowOut | Export-Csv -Path $file -NoTypeInformation -Append -Encoding UTF8 }
  if($chkCableNeeded.Checked){
    $excelPath = Join-Path $script:OutputFolder 'CablingNeeded.xlsx'
    $excel = $null
    $workbook = $null
    $worksheet = $null
    $excelPathExists = Test-Path $excelPath
    try {
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = $false
      $excel.DisplayAlerts = $false
      if($excelPathExists){
        $workbook = $excel.Workbooks.Open($excelPath)
      } else {
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $headers = @('Timestamp','Name','Asset Tag','Location','Room')
        for($i = 0; $i -lt $headers.Count; $i++){
          $worksheet.Cells.Item(1,$i + 1).Value2 = $headers[$i]
        }
      }
      if(-not $worksheet){ $worksheet = $workbook.Worksheets.Item(1) }
      $usedRange = $worksheet.UsedRange
      $lastRow = $usedRange.Rows.Count
      if($lastRow -eq 1 -and -not $worksheet.Cells.Item(1,1).Value2){ $lastRow = 0 }
      $nextRow = $lastRow + 1
      $worksheet.Cells.Item($nextRow,1).Value2 = $row.Timestamp
      $worksheet.Cells.Item($nextRow,2).Value2 = $row.Name
      $worksheet.Cells.Item($nextRow,3).Value2 = $row.AssetTag
      $worksheet.Cells.Item($nextRow,4).Value2 = $row.Location
      $worksheet.Cells.Item($nextRow,5).Value2 = $row.Room
      if($excelPathExists){ $workbook.Save() }
      else { $workbook.SaveAs($excelPath) }
    } catch {
      Write-Warning ("Failed to log cabling needed entry: " + $_.Exception.Message)
    } finally {
      if($worksheet){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null }
      if($workbook){
        try { $workbook.Close($true) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
      }
      if($excel){
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
      }
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
    }
  }
  foreach($cb in @($chkCable,$chkCableNeeded,$chkLabels,$chkCart,$chkPeriph)){ $cb.Checked = $false }
  [System.Windows.Forms.MessageBox]::Show(("Saved rounding event to
" + $file),"Save Event") | Out-Null
  $cmbChkStatus.SelectedIndex = 0
  $txtScan.Clear()
  Clear-UI
  Focus-ScanInput
  # -- Nearby: add Location-only scope and rebuild
  try {
    if ($row -and $row.Location) {
      Add-NearbyScope $null $row.Location $null $null
      Update-ScopeLabel
      Rebuild-Nearby
      Write-Host ("Main Save: Added Location scope '" + $row.Location + "' -> Count=" + $script:ActiveNearbyScopes.Count)
    } else {
      Write-Host "Main Save: Row.Location missing; not adding scope."
    }
  } catch { Write-Host ("Main Save: Error - " + $_.Exception.Message) }
  try { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.UseWaitCursor = $false } catch {}
})
$btnManualRound.Add_Click({
  if($btnManualRound.Tag){ Start-Process -FilePath $btnManualRound.Tag }
  else { [System.Windows.Forms.MessageBox]::Show("No rounding URL found for this device.","Manual Round") | Out-Null }
})
$btnFixName.Add_Click({ Fix-DisplayName })


# -------- Hardcode paths and auto-load on startup --------
try{
  $__ownDir = Get-OwnScriptDir
  $script:DataFolder   = Join-Path $__ownDir 'Data'
  $script:OutputFolder = Join-Path $__ownDir 'Output'
  if (-not (Test-Path $script:OutputFolder)) { New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null }
  if (-not (Test-Path $script:DataFolder)) {
    throw "Data folder not found:`r
$script:DataFolder`r
Create a 'Data' folder next to the script and add your CSVs."
  }
  Load-DataFolder $script:DataFolder
  Update-Counters
  try { Populate-Department-Combo ($txtDept.Text) } catch {}
  $lblDataPath.Visible=$false; $lblOutputPath.Visible=$false; $lblDataStatus.Visible=$false; $statusLabel.Text = ("Data: " + $script:DataFolder + " | Output: " + $script:OutputFolder); $statusLabel.ForeColor=[System.Drawing.Color]::DarkGreen
  if ($statusPathLabel) {
    $statusPathLabel.Text = "Data: " + $script:DataFolder + "    |    Output: " + $script:OutputFolder
  }
  $lblOutputPath.Text = "Output: " + $script:OutputFolder
  $statusLabel.Text   = "Data OK"; $statusLabel.ForeColor=[System.Drawing.Color]::DarkGreen
} catch {
  $lblDataPath.Visible=$false; $lblOutputPath.Visible=$false; $lblDataStatus.Visible=$false; $statusLabel.Text = "Data files missing or error"; $statusLabel.ForeColor=[System.Drawing.Color]::Crimson
  if ($statusPathLabel) {
    $statusPathLabel.Text = "Data: " + ($(if($script:DataFolder){$script:DataFolder}else{'(not set)'})) + "    |    Output: " + ($(if($script:OutputFolder){$script:OutputFolder}else{'(not set)'}))
  }
  $lblOutputPath.Text = "Output: " + ($(if($script:OutputFolder){$script:OutputFolder}else{'(not set)'}))
  $statusLabel.Text   = "Failed to load Data folder. See error dialog."
  $err = $_.Exception
  $diag = @()
  $diag += "Resolver diagnostics:"
  $diag += ("  PSScriptRoot: " + ($(if($PSScriptRoot){$PSScriptRoot}else{'(null)'})))
  $diag += ("  PSCommandPath: " + ($(if($PSCommandPath){$PSCommandPath}else{'(null)'})))
  $diag += ("  MyInvocation.MyCommand.Path: " + ($(if($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path){$MyInvocation.MyCommand.Path}else{'(null)'})))
  $diag += ("  env:__ScriptDir: " + ($(if($env:__ScriptDir){$env:__ScriptDir}else{'(null)'})))
  $diag += ("  Get-Location: " + (Get-Location).Path)
  $msg = "Failed to load data:
" + $err.Message + "
" + ($diag -join "
") + "
Type: " + $err.GetType().FullName + "
Stack:
" + $_.ScriptStackTrace
  [System.Windows.Forms.MessageBox]::Show($msg,"Load Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
}
$form.Add_KeyDown({ if($_.Control -and $_.KeyCode -eq 'S'){ Save-AllCSVs; $_.Handled=$true } })
# -------- Launch --------
# (ShowDialog moved to end by combiner)
# ======================= NEARBY TAB INJECTION START =======================
# (identical to the builder's injection body)
# ---- Globals for Nearby ----
if (-not $script:ActiveNearbyScopes) {
  $script:ActiveNearbyScopes = New-Object System.Collections.Generic.HashSet[string]
}
if (-not (Get-Variable -Scope Script -Name NearbyShowAllChanges -ErrorAction SilentlyContinue)) {
  $script:NearbyShowAllChanges = New-Object System.Collections.Generic.List[string]
}
if (-not (Get-Variable -Scope Script -Name NearbyLastSortColumn -ErrorAction SilentlyContinue)) {
  $script:NearbyLastSortColumn = $null
}
if (-not (Get-Variable -Scope Script -Name NearbyLastSortDirection -ErrorAction SilentlyContinue)) {
  $script:NearbyLastSortDirection = 'Asc'
}
if (-not $script:NEAR_STATUSES) {
  # Full set minus "Complete"
  $script:NEAR_STATUSES = @(
    "—",
    "Inaccessible - Asset not found",
    "Inaccessible - In storage",
    "Inaccessible - In use by Customer",
    "Inaccessible - Laptop is not onsite",
    "Inaccessible - Other",
    "Inaccessible - Restricted area",
    "Inaccessible - Room locked - Card Swipe",
    "Inaccessible - Room locked - Key Lock",
    "Inaccessible - Under renovation",
    "Inaccessible - User working at home",
    "Pending Repair"
  )
}
# In-memory cache of rounding events
if (-not (Get-Variable -Scope Script -Name RoundingEvents -ErrorAction SilentlyContinue)) {
  $script:RoundingEvents = @()
}
function Ensure-RoundingCommentsColumn([string]$file){
  try {
    if([string]::IsNullOrWhiteSpace($file)){ return }
    if(-not (Test-Path $file)){ return }
    $header = $null
    try { $header = Get-Content -Path $file -TotalCount 1 -Encoding UTF8 } catch { $header = $null }
    $hasComments = $false
    $hasRounded  = $false
    if($header){
      $hasComments = ($header -match '(^|,)"?Comments"?(,|$)')
      $hasRounded  = ($header -match '(^|,)"?Rounded"?(,|$)')
    }
    if($hasComments -and $hasRounded){ return }
    $rows = @()
    try { $rows = Import-Csv -Path $file } catch { $rows = @() }
    $columns = @()
    if($rows -and $rows.Count -gt 0){
      $columns = @($rows[0].PSObject.Properties.Name)
    } elseif($header){
      $columns = @($header -split ',')
    }
    $targetColumns = @($script:RoundingEventColumns)
    if(-not $columns -or $columns.Count -eq 0){
      $columns = $targetColumns
    } else {
      foreach($col in $targetColumns){
        if(-not ($columns -contains $col)){
          $columns += $col
        }
      }
    }
    if($rows -and $rows.Count -gt 0){
      foreach($r in $rows){
        if(-not $r.PSObject.Properties['Comments']){
          $r | Add-Member -NotePropertyName Comments -NotePropertyValue '' -Force
        }
        if(-not $r.PSObject.Properties['CablingNeeded']){
          $r | Add-Member -NotePropertyName CablingNeeded -NotePropertyValue $false -Force
        }
        if(-not $r.PSObject.Properties['Rounded']){
          $r | Add-Member -NotePropertyName Rounded -NotePropertyValue 'No' -Force
        }
      }
      $rows | Select-Object $columns | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    } else {
      $headerLine = ($columns -join ',')
      Set-Content -Path $file -Value $headerLine -Encoding UTF8
    }
  } catch { }
}
function Load-RoundingEvents {
  $script:RoundingEvents = @()
  try {
    $base = if($script:OutputFolder){ $script:OutputFolder } else { $script:DataFolder }
    if(-not $base){ return }
    $file = Join-Path $base 'RoundingEvents.csv'
    if(Test-Path $file){
      Ensure-RoundingCommentsColumn $file
      try { $script:RoundingEvents = Import-Csv $file } catch { $script:RoundingEvents = @() }
    }
  } catch { $script:RoundingEvents = @() }
}
Load-RoundingEvents
# Filtering removed from Nearby grid; keep stub to refresh visibility/count labels.
function Apply-NearbyFilters {
  if (-not $dgvNearby) { Update-ScopeLabel; return }
  foreach ($row in $dgvNearby.Rows) {
    if ($row.IsNewRow) { continue }
    $row.Visible = $true
  }
  Update-ScopeLabel
}
function ScopeKey([string]$city,[string]$loc,[string]$b,[string]$f){
  $nl = if ($loc) { (Normalize-Field $loc) } else { "" }
  if ([string]::IsNullOrWhiteSpace($nl)) { return $null }
  # Location-only key
  return $nl
}

function Add-NearbyScope([string]$city,[string]$loc,[string]$b,[string]$f){
  if (-not $script:ActiveNearbyScopes) {
    $script:ActiveNearbyScopes = New-Object System.Collections.Generic.HashSet[string]
  }
  $k = ScopeKey $null $loc $null $null
  if ($k) { [void]$script:ActiveNearbyScopes.Add($k) }
}

function Add-NearbyScopeFromDevice($pc){
  if (-not $pc) { return }
  $city = Get-City-ForLocation $pc.location
  Add-NearbyScope $city $pc.location $pc.u_building $pc.u_floor
}
function Get-RoundedToday-Set {
  $set = New-Object 'System.Collections.Generic.HashSet[string]'
  if(-not $script:RoundingEvents){ return $set }
  $today = (Get-Date).Date
  foreach ($e in $script:RoundingEvents) {
    try {
      $timestamp = $null
      if($e -and $e.PSObject.Properties['Timestamp'] -and $e.Timestamp){
        try { $timestamp = [datetime]::Parse($e.Timestamp) } catch {}
      } elseif($e.Timestamp){
        try { $timestamp = Get-Date $e.Timestamp } catch {}
      }
      if($timestamp -and $timestamp.Date -eq $today){
        $assetTag = ''
        if($e.PSObject.Properties['AssetTag'] -and $e.AssetTag){ $assetTag = $e.AssetTag }
        elseif($e.AssetTag){ $assetTag = $e.AssetTag }
        $normalized = $assetTag.Trim().ToUpper()
        if($normalized){ [void]$set.Add($normalized) }
      }
    } catch {}
  }
  return $set
}
function Get-ExcludedDevices-Set {
  $set = New-Object 'System.Collections.Generic.HashSet[string]'
  if(-not $script:Computers){ return $set }
  foreach ($pc in $script:Computers) {
    try {
      $status = ('' + $pc.u_device_rounding).Trim()
      if ($status -match '^(?i)Excluded$') {
        $assetTag = ('' + $pc.asset_tag)
        $normalized = $assetTag.Trim().ToUpper()
        if ($normalized) { [void]$set.Add($normalized) }
      }
    } catch {}
  }
  return $set
}
function Update-NearbyCheckboxLabels {
  param(
    [int]$todayCount,
    [int]$excludedCount
  )

  try {
    if ($chkTodayRounded) {
      $chkTodayRounded.Text = "Today's Rounded ({0})" -f $todayCount
    }
  } catch {}

  try {
    if ($chkShowExcluded) {
      $chkShowExcluded.Text = "Excluded ({0})" -f $excludedCount
    }
  } catch {}
}
function Get-LatestRoundForAsset([string]$assetTag,[Nullable[datetime]]$fallback){
  $best = $null
if($assetTag -and $script:RoundingEvents){
    $needle = $assetTag.Trim().ToUpper()
    foreach($e in $script:RoundingEvents){
      try {
        $assetRaw = ''
        if($e.PSObject.Properties['AssetTag'] -and $e.AssetTag){ $assetRaw = $e.AssetTag }
        elseif($e.AssetTag){ $assetRaw = $e.AssetTag }
        $candidate = $assetRaw.Trim().ToUpper()
        if($candidate -ne $needle){ continue }
        $timestamp = $null
        if($e.PSObject.Properties['Timestamp'] -and $e.Timestamp){
          try { $timestamp = [datetime]::Parse($e.Timestamp) } catch {}
        } elseif($e.Timestamp){
          try { $timestamp = Get-Date $e.Timestamp } catch {}
        }
        if($timestamp -and (-not $best -or $timestamp -gt $best)){ $best = $timestamp }
      } catch {}
    }
  }
  if($best){ return $best }
  if($fallback){ return $fallback }
  return $null
}
# ---- Build Nearby UI ----
$nearToolbar = New-Object System.Windows.Forms.Panel
$nearToolbar.Dock = 'Top'
$nearToolbar.Height = 68
$nearToolbar.BackColor = $script:ThemeColors.Header
$lblScopes = New-Object System.Windows.Forms.Label
$lblScopes.AutoSize = $true
$lblScopes.Text = "Nearby scopes: 0"
$lblScopes.Location = '8,10'
$btnNearbyShowAll = New-Object ModernUI.RoundedButton
$btnNearbyShowAll.Text = 'Show All'
$btnNearbyShowAll.AutoSize = $true
$btnNearbyShowAll.Location = '8,32'
$chkTodayRounded = New-Object System.Windows.Forms.CheckBox
$chkTodayRounded.Text = "Today's Rounded"
$chkTodayRounded.AutoSize = $true
$chkTodayRounded.Location = '120,32'
$chkTodayRounded.Checked = $false
$chkShowExcluded = New-Object System.Windows.Forms.CheckBox
$chkShowExcluded.Text = "Excluded"
$chkShowExcluded.AutoSize = $true
$chkShowExcluded.Location = '280,32'
$chkShowExcluded.Checked = $false
$chkShowExcluded.Add_CheckedChanged({ Rebuild-Nearby })
$chkTodayRounded.Add_CheckedChanged({ Rebuild-Nearby })

$lblSort = New-Object System.Windows.Forms.Label
$lblSort.AutoSize = $true
$lblSort.Text = "Sort:"
$lblSort.Location = '430,10'
$cmbSort = New-Object System.Windows.Forms.ComboBox
$cmbSort.DropDownStyle = 'DropDownList'
$cmbSort.Items.AddRange(@(
  "Host Name (A→Z)",
  "Host Name (Z→A)",
  "Room (A→Z)",
  "Room (Z→A)",
  "Last Rounded (oldest first)",
  "Last Rounded (newest first)"
try { if ($cmbSort -and $cmbSort.Items -and $cmbSort.Items.Count -gt 4) { $cmbSort.SelectedIndex = 4 } else { $cmbSort.SelectedIndex = -1 } } catch {}
$cmbSort.Visible = $false; $cmbSort.Enabled = $false
))
$cmbSort.Location = '470,6'
$cmbSort.Width = 210
$btnClearScopes = New-Object ModernUI.RoundedButton
$btnClearScopes.Text = "Clear List"
$btnClearScopes.AutoSize = $true
$btnClearScopes.Location = '700,6'
$nearToolbar.Controls.AddRange(@($lblScopes,$btnNearbyShowAll,$chkTodayRounded,$chkShowExcluded,$btnClearScopes))
Update-NearbyCheckboxLabels 0 0
$btnNearbyShowAll.Add_Click({
  try {
    if (-not $script:NearbyShowAllChanges) {
      $script:NearbyShowAllChanges = New-Object System.Collections.Generic.List[string]
    }
    if ($btnNearbyShowAll.Text -eq 'Show All') {
      $script:NearbyShowAllChanges.Clear()
      if (-not $chkTodayRounded.Checked) {
        $chkTodayRounded.Checked = $true
        [void]$script:NearbyShowAllChanges.Add('Today')
      }
      if (-not $chkShowExcluded.Checked) {
        $chkShowExcluded.Checked = $true
        [void]$script:NearbyShowAllChanges.Add('Excluded')
      }
      $btnNearbyShowAll.Text = 'Hide Again'
    } else {
      foreach ($entry in @($script:NearbyShowAllChanges)) {
        switch ($entry) {
          'Today' {
            if ($chkTodayRounded.Checked) { $chkTodayRounded.Checked = $false }
          }
          'Excluded' {
            if ($chkShowExcluded.Checked) { $chkShowExcluded.Checked = $false }
          }
        }
      }
      $script:NearbyShowAllChanges.Clear()
      $btnNearbyShowAll.Text = 'Show All'
    }
  } catch {}
})
# --- Multi-Status (apply one status to selected rows) ---
$btnSetStatus = New-Object ModernUI.RoundedButton
$btnSetStatus.Text = 'Multi-Status'
$btnSetStatus.AutoSize = $true
$multiStatusPreferred = $btnSetStatus.PreferredSize
$clearPreferred = $btnClearScopes.PreferredSize
$btnSetStatus.AutoSize = $false
$btnSetStatus.AutoEllipsis = $false
$extraWidth = 12
$desiredWidth = [Math]::Max($multiStatusPreferred.Width + $extraWidth, $clearPreferred.Width)
$btnSetStatus.MinimumSize = New-Object System.Drawing.Size($multiStatusPreferred.Width, $clearPreferred.Height)
$btnSetStatus.Size = New-Object System.Drawing.Size($desiredWidth, $clearPreferred.Height)
$btnSetStatus.Location = New-Object System.Drawing.Point(560, $btnClearScopes.Location.Y)
$nearToolbar.Controls.Add($btnSetStatus)
if (-not $menuStatus) { $menuStatus = New-Object System.Windows.Forms.ContextMenuStrip }

function Set-NearbySelectedStatus {
  param(
    [string]$Value,
    [switch]$ShowConfirmation
  )
  if ([string]::IsNullOrWhiteSpace($Value)) { return }
  try {
    $col = $dgvNearby.Columns['Status']
    if ($col -and $col -is [System.Windows.Forms.DataGridViewComboBoxColumn]) {
      if (-not $col.Items.Contains($Value)) { [void]$col.Items.Add($Value) }
    }
  } catch {}
  $count = 0
  foreach ($row in $dgvNearby.SelectedRows) {
    try {
      $cell = $row.Cells['Status']
      if ($cell -and $cell -is [System.Windows.Forms.DataGridViewComboBoxCell]) {
        if (-not $cell.Items.Contains($Value)) { [void]$cell.Items.Add($Value) }
      }
      if ($row -and $row.Cells['Status']) {
        $row.Cells['Status'].Value = $Value
        $count++
      }
    } catch {}
  }
  if ($ShowConfirmation) {
    [System.Windows.Forms.MessageBox]::Show("Updated status for $count row(s).","Multi-Status",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
  }
}

function Show-NearbyStatusMenu {
  param(
    [System.Windows.Forms.Control]$Anchor,
    [System.Drawing.Point]$Location,
    [switch]$ShowConfirmation
  )
  try {
    $menuStatus.Items.Clear()
    $options = Get-StatusOptionsFromGrid
    foreach ($opt in $options) {
      $item = New-Object System.Windows.Forms.ToolStripMenuItem($opt)
      $item.Add_Click({
        param($s,$e)
        if ($ShowConfirmation) {
          Set-NearbySelectedStatus -Value $s.Text -ShowConfirmation
        } else {
          Set-NearbySelectedStatus -Value $s.Text
        }
      })
      [void]$menuStatus.Items.Add($item)
    }
    [void]$menuStatus.Items.Add('-')
    $custom = New-Object System.Windows.Forms.ToolStripMenuItem('Custom...')
    $custom.Add_Click({
      Add-Type -AssemblyName Microsoft.VisualBasic
      $val = [Microsoft.VisualBasic.Interaction]::InputBox('Enter custom status for selected rows:','Multi-Status','')
      if ([string]::IsNullOrWhiteSpace($val)) { return }
      if ($ShowConfirmation) {
        Set-NearbySelectedStatus -Value $val -ShowConfirmation
      } else {
        Set-NearbySelectedStatus -Value $val
      }
    })
    [void]$menuStatus.Items.Add($custom)
    if (-not $Location) {
      $Location = New-Object System.Drawing.Point(0,0)
    }
    if ($Anchor) {
      $menuStatus.Show($Anchor, $Location)
    } else {
      $menuStatus.Show($btnSetStatus, (New-Object System.Drawing.Point(0,$btnSetStatus.Height)))
    }
  } catch {
    [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)","Multi-Status",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
  }
}

$btnSetStatus.Add_Click({
  $pt = New-Object System.Drawing.Point(0,$btnSetStatus.Height)
  Show-NearbyStatusMenu $btnSetStatus $pt -ShowConfirmation
})
$dgvNearby = New-Object System.Windows.Forms.DataGridView
$dgvNearby.Dock='Fill'
$dgvNearby.AllowUserToAddRows=$false
$dgvNearby.ReadOnly=$false
$dgvNearby.SelectionMode=[System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgvNearby.MultiSelect=$true
$dgvNearby.RowHeadersVisible=$false
$dgvNearby.BackgroundColor=[System.Drawing.Color]::White
$dgvNearby.BorderStyle='FixedSingle'
$dgvNearby.add_CellMouseClick({
  param($sender,$e)
  if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
  if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $dgvNearby.Rows.Count) {
    try {
      $row = $dgvNearby.Rows[$e.RowIndex]
      if ($row -and -not $row.Selected) {
        $dgvNearby.ClearSelection()
        $row.Selected = $true
      }
    } catch {}
  }
  if ($dgvNearby.SelectedRows.Count -le 0) { return }
  $clientPt = $dgvNearby.PointToClient([System.Windows.Forms.Cursor]::Position)
  Show-NearbyStatusMenu $dgvNearby $clientPt
})
$dgvNearby.AutoSizeColumnsMode='DisplayedCells'
$dgvNearby.AutoGenerateColumns=$false
try { $dgvNearby.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags] 'NonPublic,Instance').SetValue($dgvNearby, $true, $null) } catch {}
function New-NearCol([string]$name,[string]$header,[int]$width,[bool]$ro=$true){
  $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $col.Name=$name; $col.HeaderText=$header; $col.Width=[math]::Max($width,60); $col.MinimumWidth=60; $col.ReadOnly=$ro
  return $col
}
# Visible columns
$dgvNearby.Columns.Add((New-NearCol 'Host' 'Host Name' 140))         | Out-Null
$dgvNearby.Columns.Add((New-NearCol 'Asset' 'Asset Tag' 110))        | Out-Null
$dgvNearby.Columns.Add((New-NearCol 'Location' 'Location' 120))      | Out-Null
$dgvNearby.Columns.Add((New-NearCol 'Building' 'Building' 110))      | Out-Null
$dgvNearby.Columns.Add((New-NearCol 'Floor' 'Floor' 80))             | Out-Null
$dgvNearby.Columns.Add((New-NearCol 'Room' 'Room' 90))               | Out-Null
$dgvNearby.Columns.Add((New-NearCol 'Department' 'Department' 160)) | Out-Null
$dgvNearby.Columns.Add((New-NearCol 'MaintenanceType' 'Maintenance Type' 160)) | Out-Null
$dgvNearby.Columns.Add((New-NearCol 'LastRounded' 'Last Rounded' 130)) | Out-Null
$dgvNearby.Columns.Add((New-NearCol 'DaysAgo' 'Days Ago' 90))        | Out-Null
# Status Combo column
$colStatus = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colStatus.Name = 'Status'
$colStatus.HeaderText = 'Status'
$colStatus.FlatStyle = 'Popup'
$colStatus.Width = 220
$colStatus.MinimumWidth = 160
$colStatus.DataSource = $script:NEAR_STATUSES
$colStatus.ReadOnly = $false
$dgvNearby.Columns.Add($colStatus) | Out-Null
# Hidden helper columns
# --- Enable header-click sorting on the unbound grid ---
try {
  # Programmatic sort allows us to control compare logic
  foreach ($col in $dgvNearby.Columns) {
    if ($col -and -not $col.Name.StartsWith('__')) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic }
  }
  # Custom comparer so we can sort dates numerically via LRRAW while showing LR
  $dgvNearby.add_SortCompare({
    param($sender, $e)
    try {
      $colName = $e.Column.Name
      $v1 = $e.CellValue1
      $v2 = $e.CellValue2
      # Prefer raw date for Last Rounded
      if ($colName -eq 'LR') {
        $r1 = $sender.Rows[$e.RowIndex1].Cells['LRRAW'].Value
        $r2 = $sender.Rows[$e.RowIndex2].Cells['LRRAW'].Value
        if ($r1 -and $r2) {
          $t1 = [datetime]::Parse($r1); $t2 = [datetime]::Parse($r2)
          $e.SortResult = [System.DateTime]::Compare($t1, $t2)
          $e.Handled = $true
          return
        }
      }
      # Default string compare (case-insensitive)
      $s1 = if ($v1) { [string]$v1 } else { "" }
      $s2 = if ($v2) { [string]$v2 } else { "" }
      $e.SortResult = [string]::Compare($s1, $s2, $true)  # ignore case
      $e.Handled = $true
    } catch {
      $e.SortResult = 0; $e.Handled = $true
    }
  })
  if (-not $script:NearbySortDir) { $script:NearbySortDir = @{} }
  $dgvNearby.add_ColumnHeaderMouseClick({
    param($sender, $e)
    $col = $sender.Columns[$e.ColumnIndex]
    if (-not $col) { return }
    $name = $col.Name
    $dir = if ($script:NearbySortDir[$name] -eq 'Asc') { 'Desc' } else { 'Asc' }
    $script:NearbySortDir[$name] = $dir
    $script:NearbyLastSortColumn = $name
    $script:NearbyLastSortDirection = $dir
    $lsd = if ($dir -eq 'Asc') { [System.ComponentModel.ListSortDirection]::Ascending } else { [System.ComponentModel.ListSortDirection]::Descending }
    $sender.Sort($col, $lsd)
    $col.HeaderCell.SortGlyphDirection = if ($dir -eq 'Asc') { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending }
  })
} catch {}
$colHiddenAT = New-NearCol 'AT_KEY' '__ATKEY' 10 $true; $colHiddenAT.Visible=$false; $dgvNearby.Columns.Add($colHiddenAT) | Out-Null
$colHiddenToday = New-NearCol 'TODAY' '__TODAY' 10 $true; $colHiddenToday.Visible=$false; $dgvNearby.Columns.Add($colHiddenToday) | Out-Null
$colHiddenLRRaw = New-NearCol 'LRRAW' '__LRRAW' 10 $true; $colHiddenLRRaw.Visible=$false; $dgvNearby.Columns.Add($colHiddenLRRaw) | Out-Null
$dgvNearby.add_CurrentCellDirtyStateChanged({
  if ($dgvNearby.CurrentCell -and $dgvNearby.CurrentCell.OwningColumn -and $dgvNearby.CurrentCell.OwningColumn.Name -eq 'Status') {
    try { $dgvNearby.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } catch {}
  }
})
$dgvNearby.add_CellValueChanged({
  param($sender,$e)
  if ($e.ColumnIndex -lt 0) { return }
  $col = $sender.Columns[$e.ColumnIndex]
  if ($col -and $col.Name -eq 'Status') { Apply-NearbyFilters }
})
# Bottom bar
$nearBottom = New-Object System.Windows.Forms.Panel
$nearBottom.Dock = 'Bottom'
$nearBottom.Height = 48
$lblNearNote = New-Object System.Windows.Forms.Label
$lblNearNote.AutoSize = $true
$lblNearNote.Text = "Bulk Save adds events with selected Status, 3 min each. Today's devices are shown only when ""View all"" is on and are not saved again."
$lblNearNote.Location = '8,14'
$btnNearSave = New-Object ModernUI.RoundedButton
$btnNearSave.Text = "Save"
$btnNearSave.Anchor = 'Bottom,Right'
$btnNearSave.Size = '120,30'
$btnNearSave.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 160), 8)
$nearBottom.Add_Resize({ $btnNearSave.Location = New-Object System.Drawing.Point(($nearBottom.ClientSize.Width - 128 - 12), 8) })
$nearBottom.Controls.AddRange(@($lblNearNote,$btnNearSave))
# Build Nearby tab page
$tabTop = New-Object System.Windows.Forms.TabControl
$tabTop.Dock = 'Fill'
$tabPageMain = New-Object System.Windows.Forms.TabPage
$tabPageMain.Text = 'Main'
$tabPageMain.UseVisualStyleBackColor = $false
$tabPageNear = New-Object System.Windows.Forms.TabPage
$tabPageNear.Text = 'Nearby'
$tabPageNear.UseVisualStyleBackColor = $false
# Move existing main UI into a panel and then into TabPageMain
$pageMain = New-Object System.Windows.Forms.Panel
$pageMain.Dock = 'Fill'
# Re-parent header + main table into panel
$form.Controls.Remove($panelTop)
$form.Controls.Remove($splitter)
$pageMain.Controls.Add($splitter); $splitter.Dock='Fill'
$pageMain.Controls.Add($panelTop); $panelTop.Dock='Top'
$tabPageMain.Controls.Add($pageMain)
# Compose Nearby page
$nearPage = New-Object System.Windows.Forms.Panel
$nearPage.Dock = 'Fill'
$nearPage.Controls.Add($dgvNearby)
$nearPage.Controls.Add($nearBottom)
$nearPage.Controls.Add($nearToolbar)
$tabPageNear.Controls.Add($nearPage)
$tabTop.TabPages.AddRange(@($tabPageMain,$tabPageNear))
# Put the TabControl on the form (above status strip)
$form.Controls.Add($tabTop)
$form.Controls.SetChildIndex($tabTop, 0)  # ensure it's above the status strip
# ---- Nearby logic ----
function Update-ScopeLabel {
  $scopeCount = if ($script:ActiveNearbyScopes) { $script:ActiveNearbyScopes.Count } else { 0 }
  $text = "Nearby scopes (Location): $scopeCount"
  if ($dgvNearby) {
    $total = 0
    $visible = 0
    foreach ($row in $dgvNearby.Rows) {
      if ($row.IsNewRow) { continue }
      $total++
      if ($row.Visible) { $visible++ }
    }
    if ($total -gt 0) {
      if ($visible -ne $total) {
        $text += (" — Showing {0} of {1}" -f $visible, $total)
      } else {
        $text += (" — Showing {0}" -f $total)
      }
    }
  }
  $lblScopes.Text = $text
}
function Should-Include-PC-InScopes($pc){
  if (-not $pc) { return $false }
  $k = ScopeKey $null $pc.location $null $null
  if (-not $k) { return $false }
  if (-not $script:ActiveNearbyScopes) { return $false }
  return $script:ActiveNearbyScopes.Contains($k)
}

function Get-RoundingStatusColor([Nullable[DateTime]]$dt){
  $s = Get-RoundingStatus $dt
  if ($s -eq 'Green') { return [System.Drawing.Color]::PaleGreen }
  if ($s -eq 'Yellow'){ return [System.Drawing.Color]::LightYellow }
  return [System.Drawing.Color]::MistyRose
}
function Rebuild-Nearby {
try { $dgvNearby.SuspendLayout() } catch {}
try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor } catch {}
  try { Write-Host ("Rebuild-Nearby: Active scopes=" + ($(if($script:ActiveNearbyScopes){$script:ActiveNearbyScopes.Count}else{0}))) } catch {}
  Load-RoundingEvents
  $todaySet = Get-RoundedToday-Set
  $excludedSet = Get-ExcludedDevices-Set
  $todayCount = 0
  $excludedCount = 0
  $dgvNearby.Rows.Clear()
  $seen = New-Object System.Collections.Generic.HashSet[string]
  foreach ($pc in $script:Computers) {
    if (-not (Should-Include-PC-InScopes $pc)) { continue }
    $at = $pc.asset_tag
    $atKey = ($at + "").Trim().ToUpper()
    if ($atKey) {
      if ($seen.Contains($atKey)) { continue } else { [void]$seen.Add($atKey) }
    }
    $lr = Get-LatestRoundForAsset $at $pc.LastRounded
    $days = ""
    if ($lr) {
      $d = [int]((Get-Date).Date - $lr.Date).TotalDays
      $days = $d
    }
    $isToday = $false
    if ($atKey -and $todaySet.Contains($atKey)) { $isToday = $true }
    $isExcluded = $false
    try {
      $roundingFlag = ('' + $pc.u_device_rounding).Trim()
      if ($roundingFlag -match '^(?i)Excluded$') { $isExcluded = $true }
    } catch {}
    if (-not $isExcluded -and $atKey -and $excludedSet.Contains($atKey)) { $isExcluded = $true }
    if ($isToday) { $todayCount++ }
    if ($isExcluded) { $excludedCount++ }
    if (-not $chkTodayRounded.Checked -and $isToday) { continue }
    if (-not $chkShowExcluded.Checked -and $isExcluded) { continue }
    $rowIdx = $dgvNearby.Rows.Add()
    $r = $dgvNearby.Rows[$rowIdx]
    $r.Cells['Host'].Value      = $pc.name
    $r.Cells['Asset'].Value     = $pc.asset_tag
    $r.Cells['Location'].Value  = $pc.location
    $r.Cells['Building'].Value  = $pc.u_building
    $r.Cells['Floor'].Value     = $pc.u_floor
    $r.Cells['Room'].Value      = $pc.u_room
    $r.Cells['Department'].Value = $pc.u_department_location
    $mtRaw = ''
    try {
      if($pc.PSObject.Properties['u_device_rounding']){ $mtRaw = '' + $pc.u_device_rounding }
    } catch {}
    $r.Cells['MaintenanceType'].Value = Get-MaintenanceTypeOrDefault $mtRaw ([string]$pc.name)
    $r.Cells['LastRounded'].Value = (Fmt-DateLong $lr)
    $r.Cells['DaysAgo'].Value   = $days
    $r.Cells['Status'].Value    = "—"
    $r.Cells['AT_KEY'].Value    = $atKey
    $r.Cells['TODAY'].Value     = if ($isToday) { "1" } else { "0" }
    $r.Cells['LRRAW'].Value     = if ($lr) { $lr.ToString("o") } else { "" }
    # Style
    if ($lr) {
      $r.Cells['LastRounded'].Style.BackColor = Get-RoundingStatusColor $lr
    } else {
      $r.Cells['LastRounded'].Style.BackColor = [System.Drawing.Color]::MistyRose
    }
    if ($isToday) {
      $r.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
    }
  }
  # Apply sort and filters
  Apply-NearbySort
  Apply-NearbyFilters
  Update-NearbyCheckboxLabels $todayCount $excludedCount
try { $dgvNearby.ResumeLayout() } catch {}
try { $form.Cursor = [System.Windows.Forms.Cursors]::Default } catch {}}
function Apply-NearbySort {
  if ($dgvNearby -and $script:NearbyLastSortColumn) {
    $col = $dgvNearby.Columns[$script:NearbyLastSortColumn]
    if ($col) {
      try {
        $lsd = if ($script:NearbyLastSortDirection -eq 'Desc') { [System.ComponentModel.ListSortDirection]::Descending } else { [System.ComponentModel.ListSortDirection]::Ascending }
        $dgvNearby.Sort($col, $lsd)
        foreach ($otherCol in $dgvNearby.Columns) {
          if (-not $otherCol -or $otherCol -eq $col) { continue }
          try { $otherCol.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None } catch {}
        }
        $col.HeaderCell.SortGlyphDirection = if ($lsd -eq [System.ComponentModel.ListSortDirection]::Ascending) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending }
        return
      } catch {
        $script:NearbyLastSortColumn = $null
      }
    } else {
      $script:NearbyLastSortColumn = $null
    }
  }
  $items = @()
  foreach ($row in $dgvNearby.Rows) {
    if ($row.IsNewRow) { continue }
    $items += [pscustomobject]@{
      Row=$row
      Host=[string]$row.Cells['Host'].Value
      Room=[string]$row.Cells['Room'].Value
      LR = $(try { if ($row.Cells['LRRAW'].Value) { [datetime]::Parse($row.Cells['LRRAW'].Value) } else { $null } } catch { $null })
    }
  }
  $sorted = $items
  switch ($cmbSort.SelectedItem) {
    'Host Name (A→Z)' { $sorted = $items | Sort-Object Host }
    'Host Name (Z→A)' { $sorted = $items | Sort-Object Host -Descending }
    'Room (A→Z)'      { $sorted = $items | Sort-Object Room }
    'Room (Z→A)'      { $sorted = $items | Sort-Object Room -Descending }
    'Last Rounded (oldest first)' {
      $sorted = $items | Sort-Object @{Expression={ if ($_.LR) { $_.LR } else { Get-Date "1900-01-01" } }}
    }
    'Last Rounded (newest first)' {
      $sorted = $items | Sort-Object @{Expression={ if ($_.LR) { $_.LR } else { Get-Date "1900-01-01" } }} -Descending
    }
    default { $sorted = $items | Sort-Object @{Expression={ if ($_.LR) { $_.LR } else { Get-Date "1900-01-01" } }} }
  }
  # Reorder rows in the grid
  $idx = 0
  foreach ($it in $sorted) {
    if ($it.Row -and $it.Row.PSObject.Properties['DisplayIndex']) {
      try {
        $it.Row.DisplayIndex = $idx
      } catch {}
      $idx++
    }
  }
}
# Double-click: open on Main and switch
$dgvNearby.Add_CellDoubleClick({
  if ($_.RowIndex -lt 0) { return }
  $row = $dgvNearby.Rows[$_.RowIndex]
  $asset = [string]$row.Cells['Asset'].Value
  $serial = ""  # not needed
  $name = [string]$row.Cells['Host'].Value
  $rec = $null
  if ($asset) {
    $key = $asset.Trim().ToUpper()
    if ($script:IndexByAsset.ContainsKey($key)) { $rec = $script:IndexByAsset[$key] }
  }
  if (-not $rec -and $name) {
    foreach ($k in (HostnameKeyVariants $name)) { if ($script:IndexByName.ContainsKey($k)) { $rec = $script:IndexByName[$k]; break } }
  }
  if ($rec) {
    $par = Resolve-ParentComputer $rec
    Populate-UI $rec $par
    Start-RoundingTimer
    $tabTop.SelectedTab = $tabPageMain
  }
})
# React to toolbar changes
$cmbSort.Add_SelectedIndexChanged({})
$btnClearScopes.Add_Click({
  $script:ActiveNearbyScopes.Clear()
  $dgvNearby.Rows.Clear()
  Update-ScopeLabel
})
# Bulk Save from Nearby
$btnNearSave.Add_Click({
  $out = $script:OutputFolder
  if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out -Force | Out-Null }
$file = Join-Path ($(if($script:OutputFolder){$script:OutputFolder}else{$script:DataFolder})) 'RoundingEvents.csv'
  $exists = Test-Path $file
  if($exists){ Ensure-RoundingCommentsColumn $file }
  $todaySet = Get-RoundedToday-Set
  $saved = 0
  foreach ($row in $dgvNearby.Rows) {
    if ($row.IsNewRow) { continue }
    $status = [string]$row.Cells['Status'].Value
    if (-not $status -or $status -eq '—') { continue }
    $asset = [string]$row.Cells['Asset'].Value
    $atKey = if ($asset) { $asset.Trim().ToUpper() } else { "" }
    if ($atKey -and $todaySet.Contains($atKey)) { continue } # don't duplicate today's
    # lookup pc for city/other fields
    $pc = $null
    if ($asset) {
      $k = $asset.Trim().ToUpper()
      if ($script:IndexByAsset.ContainsKey($k)) { $pc = $script:IndexByAsset[$k] }
    }
    if (-not $pc -and $row.Cells['Host'].Value) {
      foreach ($k2 in (HostnameKeyVariants ([string]$row.Cells['Host'].Value))) { if ($script:IndexByName.ContainsKey($k2)) { $pc = $script:IndexByName[$k2]; break } }
    }
    $city = if ($pc) { Get-City-ForLocation $pc.location } else { "" }
$url = $null
if ($pc) {
  try { $url = Get-RoundingUrlForParent $pc } catch { $url = $null }
}
    $mtRaw = ''
    if ($pc) {
      try {
        if ($pc.PSObject.Properties['u_device_rounding']) { $mtRaw = '' + $pc.u_device_rounding }
      } catch {}
    }
    $mtDeviceName = ''
    if ($pc) {
      try { $mtDeviceName = [string]$pc.name } catch {}
    }
    if (-not $mtDeviceName) {
      try { $mtDeviceName = [string]$row.Cells['Host'].Value } catch {}
    }
    $mtValue = Get-MaintenanceTypeOrDefault $mtRaw $mtDeviceName
    $ev = [pscustomobject]@{
      Timestamp        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
      AssetTag         = $asset
      Name             = $row.Cells['Host'].Value
      Serial           = if ($pc) { $pc.serial_number } else { $null }
      City             = $city
      Location         = $row.Cells['Location'].Value
      Building         = $row.Cells['Building'].Value
      Floor            = $row.Cells['Floor'].Value
      Room             = $row.Cells['Room'].Value
      CheckStatus      = $status
      RoundingMinutes  = 3
      CableMgmtOK      = 'No'
      CablingNeeded    = 'No'
      LabelOK          = 'No'
      CartOK           = 'No'
      PeripheralsOK    = 'No'
      MaintenanceType  = $mtValue
      Department       = $row.Cells['Department'].Value
      RoundingUrl      = $url
      Comments         = ''
      Rounded          = 'No'
    }
    $evOut = $ev | Select-Object $script:RoundingEventColumns
    if (-not $exists) { $evOut | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8 }
    else { $evOut | Export-Csv -Path $file -NoTypeInformation -Append -Encoding UTF8 }
    # Update in-memory
    if (-not ($script:RoundingEvents -is [System.Collections.IList])) {
      $script:RoundingEvents = @($script:RoundingEvents)
    }
    $script:RoundingEvents += $ev
    if ($pc) { $pc | Add-Member -NotePropertyName LastRounded -NotePropertyValue (Get-Date) -Force }
    $saved++
  }
  if ($saved -gt 0) {
    [System.Windows.Forms.MessageBox]::Show(("Saved {0} rounding event(s)." -f $saved),"Nearby Save") | Out-Null
    Rebuild-Nearby
  } else {
    [System.Windows.Forms.MessageBox]::Show("Nothing to save. Pick a Status first.","Nearby Save") | Out-Null
  }
})
# Hook Save on Main to accumulate scopes (based on last event appended)
$btnSave.Add_Click({
  Start-Sleep -Milliseconds 120
  try {
    $out = $script:OutputFolder
$file = Join-Path ($(if($script:OutputFolder){$script:OutputFolder}else{$script:DataFolder})) 'RoundingEvents.csv'
    if (Test-Path $file) {
      $rows = Import-Csv $file
      if ($rows.Count -gt 0) {
        $last = $rows[-1]
        # only add scope if it's today's event
        $dt = Get-Date $last.Timestamp
        if ($dt.Date -eq (Get-Date).Date) {
          Add-NearbyScope $last.City $last.Location $last.Building $last.Floor
          Rebuild-Nearby
        }
      }
    }
  } catch { }
})
# Initial build
Update-ScopeLabel
# ======================== NEARBY TAB INJECTION END ========================

function Get-StatusOptionsFromGrid {
  $opts = New-Object System.Collections.Generic.List[string]
  try {
    $col = $dgvNearby.Columns['Status']
    if ($col -and $col -is [System.Windows.Forms.DataGridViewComboBoxColumn]) {
      foreach ($i in $col.Items) { $opts.Add([string]$i) }
    } elseif ($dgvNearby.Rows.Count -gt 0) {
      $cell0 = $dgvNearby.Rows[0].Cells['Status']
      if ($cell0 -and $cell0 -is [System.Windows.Forms.DataGridViewComboBoxCell]) {
        foreach ($i in $cell0.Items) { $opts.Add([string]$i) }
      }
    }
  } catch {}
  if ($opts.Count -eq 0) {
    foreach ($v in @('OK','Checked','In Progress','Needs Attention','Out of Service','Retire','Escalated','Unknown')) { $opts.Add($v) }
  }
  # Distinct, sorted
  return ($opts | Where-Object { $_ -and $_.Trim().Length -gt 0 } | Sort-Object -Unique)
}

Apply-ModernThemeToForm -Form $form

# The theming utility enforces single-selection on all DataGridViews. Restore
# multi-selection for the Nearby grid so Shift/Ctrl selection works as
# intended.
try { if ($dgvNearby) { $dgvNearby.MultiSelect = $true } } catch {}

$form.BackColor = $script:ThemeColors.Background
try {
  $panelTop.BackColor = $script:ThemeColors.Header
  $nearToolbar.BackColor = $script:ThemeColors.Header
  $status.BackColor = $script:ThemeColors.Header
  $status.ForeColor = $script:ThemeColors.Text
} catch {}
try {
  $splitter.BackColor = $script:ThemeColors.Background
  $pageMain.BackColor = $script:ThemeColors.Background
  $nearPage.BackColor = $script:ThemeColors.Background
  $tabPageMain.BackColor = $script:ThemeColors.Background
  $tabPageNear.BackColor = $script:ThemeColors.Background
  $tlpAssoc.BackColor = $script:ThemeColors.Surface
  $assocToolbarPanel.BackColor = $script:ThemeColors.Surface
  $assocGridPanel.BackColor = $script:ThemeColors.Surface
  $cards.BackColor = $script:ThemeColors.Background
} catch {}

$script:NewAssetToolMainForm = $form

$shouldShowForm = $true
try {
  if ($global:NewAssetToolSuppressShow) { $shouldShowForm = $false }
} catch {}
try {
  if ($script:NewAssetToolSuppressShow) { $shouldShowForm = $false }
} catch {}

if ($shouldShowForm) {
  [void]$form.ShowDialog()
} else {
  return $form
}
