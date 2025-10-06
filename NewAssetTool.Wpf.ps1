Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Xaml
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName WindowsFormsIntegration

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

  [DllImport("shcore.dll")]
  public static extern int SetProcessDpiAwareness(ProcessDpiAwareness value);

  [DllImport("user32.dll")]
  public static extern IntPtr GetThreadDpiAwarenessContext();

  [DllImport("user32.dll")]
  public static extern int GetAwarenessFromDpiAwarenessContext(IntPtr value);
}

public enum ProcessDpiAwareness
{
  ProcessDpiUnaware = 0,
  ProcessSystemDpiAware = 1,
  ProcessPerMonitorDpiAware = 2
}
"@ -ErrorAction Stop
  } catch {
    Write-Verbose "[DPI] Failed to define native DPI helpers: $($_.Exception.Message)" -Verbose
  }
}

$script:NewAssetToolPerMonitorDpiContextEnabled = $false
if (-not (Get-Variable -Scope Global -Name NewAssetToolPerMonitorDpiContextEnabled -ErrorAction SilentlyContinue)) {
  $global:NewAssetToolPerMonitorDpiContextEnabled = $false
}

function Set-NewAssetToolProcessDpiAwareness {
  $perMonitorV2Context = [System.IntPtr]-4
  $contextEnabled = $false

  if ('NewAssetTool.NativeMethods.Dpi' -as [Type]) {
    try {
      if ([NewAssetTool.NativeMethods.Dpi]::SetProcessDpiAwarenessContext($perMonitorV2Context)) {
        $contextEnabled = $true
        Write-Verbose "[DPI] Opted into PerMonitorV2 awareness via SetProcessDpiAwarenessContext." -Verbose
      } else {
        $lastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Verbose "[DPI] SetProcessDpiAwarenessContext returned false (LastError=$lastError)." -Verbose
      }
    } catch [System.EntryPointNotFoundException] {
      Write-Verbose "[DPI] SetProcessDpiAwarenessContext not available; falling back." -Verbose
    } catch {
      Write-Verbose "[DPI] Failed to call SetProcessDpiAwarenessContext: $($_.Exception.Message)" -Verbose
    }

    if (-not $contextEnabled) {
      try {
        $hresult = [NewAssetTool.NativeMethods.Dpi]::SetProcessDpiAwareness([NewAssetTool.NativeMethods.ProcessDpiAwareness]::ProcessPerMonitorDpiAware)
        if ($hresult -eq 0) {
          $contextEnabled = $true
          Write-Verbose "[DPI] Opted into PerMonitor awareness via SetProcessDpiAwareness fallback." -Verbose
        } else {
          Write-Verbose ("[DPI] SetProcessDpiAwareness fallback failed (HRESULT=0x{0:X8})." -f $hresult) -Verbose
        }
      } catch {
        Write-Verbose "[DPI] Failed to call SetProcessDpiAwareness fallback: $($_.Exception.Message)" -Verbose
      }
    }
  } else {
    Write-Verbose "[DPI] Native DPI helper type unavailable; skipping opt-in." -Verbose
  }

  if ($contextEnabled) {
    $script:NewAssetToolPerMonitorDpiContextEnabled = $true
    $global:NewAssetToolPerMonitorDpiContextEnabled = $true
  }

  return $contextEnabled
}

function Get-LoaderScriptDir {
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
  return (Get-Location).Path
}

$scriptDir = Get-LoaderScriptDir
$ps1Path   = Join-Path $scriptDir 'NewAssetTool.ps1'
$xamlPath  = Join-Path $scriptDir 'NewAssetTool.xaml'

if (-not (Test-Path $ps1Path)) {
  throw "Could not find NewAssetTool.ps1 at '$ps1Path'."
}
if (-not (Test-Path $xamlPath)) {
  throw "Could not find NewAssetTool.xaml at '$xamlPath'."
}

$dpiAwarenessApplied = Set-NewAssetToolProcessDpiAwareness
if (-not $dpiAwarenessApplied) {
  Write-Verbose "[DPI] Per-monitor DPI awareness opt-in failed; continuing with default context." -Verbose
}

$previousSuppress = $null
$hadPrevious = $false
try {
  if (Test-Path Variable:\global:NewAssetToolSuppressShow) {
    $previousSuppress = $global:NewAssetToolSuppressShow
    $hadPrevious = $true
  }
} catch {}
$global:NewAssetToolSuppressShow = $true

try {
  $scriptOutput = . $ps1Path
} catch {
  if ($hadPrevious) {
    $global:NewAssetToolSuppressShow = $previousSuppress
  } else {
    Remove-Variable -Scope Global -Name NewAssetToolSuppressShow -ErrorAction SilentlyContinue
  }
  throw
}

if ($scriptOutput -is [System.Windows.Forms.Form]) {
  $form = $scriptOutput
} else {
  $form = ($scriptOutput | Where-Object { $_ -is [System.Windows.Forms.Form] } | Select-Object -First 1)
}

if (-not $form -and $script:NewAssetToolMainForm -is [System.Windows.Forms.Form]) {
  $form = $script:NewAssetToolMainForm
}

if (-not $form -or -not ($form -is [System.Windows.Forms.Form])) {
  throw "NewAssetTool.ps1 did not expose a main form instance."
}

$form.TopLevel = $false
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$form.Dock = [System.Windows.Forms.DockStyle]::Fill

if (-not (Test-Path $xamlPath)) {
  throw "XAML window definition missing at '$xamlPath'."
}

$reader = [System.Xml.XmlReader]::Create($xamlPath)
try {
  $window = [System.Windows.Markup.XamlReader]::Load($reader)
} finally {
  $reader.Close()
}

$windowsFormsHost = $window.FindName('WinFormsHost')
if (-not $windowsFormsHost) {
  throw "Could not locate the WindowsFormsHost named 'WinFormsHost' in XAML."
}
$windowsFormsHost.Child = $form
$form.Visible = $true

$rootGrid = $window.FindName('RootGrid')
$rootScaleTransform = $window.FindName('RootScaleTransform')
$applyWpfScale = {
  param([string]$Source = 'unspecified')

  if (-not $global:NewAssetToolPerMonitorDpiContextEnabled) { return }
  if (-not $rootScaleTransform) {
    Write-Verbose "[DPI][WPF] Root scale transform not located; skipping scale." -Verbose
    return
  }

  $contextDescription = 'unknown'
  try {
    if (Get-Command Get-NewAssetToolDpiContextDescription -ErrorAction SilentlyContinue) {
      $contextDescription = Get-NewAssetToolDpiContextDescription
    }
  } catch {}

  try {
    $rootScaleTransform.ScaleX = 0.8
    $rootScaleTransform.ScaleY = 0.8
    Write-Verbose (
      "[DPI][WPF] Applied root layout scale 0.8 ({0}) context={1}" -f $Source, $contextDescription
    ) -Verbose
  } catch {
    Write-Verbose "[DPI][WPF] Failed to apply layout scale ({0}): $($_.Exception.Message)" -Verbose
  }
}

if ($global:NewAssetToolPerMonitorDpiContextEnabled -and $window) {
  try {
    $window.Add_DpiChanged({ param($sender,$eventArgs) & $applyWpfScale 'Window.DpiChanged' })
  } catch {
    Write-Verbose "[DPI][WPF] Failed to attach DpiChanged handler: $($_.Exception.Message)" -Verbose
  }
  & $applyWpfScale 'Window.Initial'
}

if ($window) {
  if ($window.WindowState -ne [System.Windows.WindowState]::Normal) {
    $window.WindowState = [System.Windows.WindowState]::Normal
  }
  if ($window.SizeToContent -ne [System.Windows.SizeToContent]::Height) {
    $window.SizeToContent = [System.Windows.SizeToContent]::Height
  }
  $sizeLockApplied = $false
  $window.Add_ContentRendered({
    param($sender, $eventArgs)
    if (-not $sizeLockApplied) {
      if ($sender.SizeToContent -ne [System.Windows.SizeToContent]::Manual) {
        $sender.SizeToContent = [System.Windows.SizeToContent]::Manual
      }
      $sizeLockApplied = $true
    }
  })
}

$searchTextBox = $window.FindName('SearchTextBox')
if ($searchTextBox) {
  try { Set-ScanSearchControl $searchTextBox } catch {}
  if ($txtScan) {
    if ($searchTextBox.Text -ne $txtScan.Text) { $searchTextBox.Text = $txtScan.Text }
    $txtScan.Add_TextChanged({
      param($sender,$eventArgs)
      $target = $script:NewAssetToolSearchTextBox
      if ($target -and $target.Text -ne $sender.Text) {
        $target.Text = $sender.Text
        try { $target.CaretIndex = $target.Text.Length } catch {}
      }
    })
  }
  $searchTextBox.Add_TextChanged({
    param($sender,$eventArgs)
    if ($txtScan -and $txtScan.Text -ne $sender.Text) {
      $txtScan.Text = $sender.Text
    }
  })
  $searchTextBox.Add_KeyDown({
    param($sender,$eventArgs)
    if ($eventArgs.Key -eq [System.Windows.Input.Key]::Enter) {
      $eventArgs.Handled = $true
      try { Do-Lookup } catch {}
    }
  })
  try { Focus-ScanInput } catch {}
}

$app = [System.Windows.Application]::Current
if (-not $app) { $app = New-Object System.Windows.Application }
$app.ShutdownMode = [System.Windows.ShutdownMode]::OnMainWindowClose
$app.MainWindow = $window

$script:__newAssetToolCleanupRan = $false
$cleanupAction = {
  if (-not $script:__newAssetToolCleanupRan) {
    $script:__newAssetToolCleanupRan = $true
    try { if ($form) { $form.Dispose() } } catch {}
    try {
      if ($hadPrevious) {
        $global:NewAssetToolSuppressShow = $previousSuppress
      } else {
        Remove-Variable -Scope Global -Name NewAssetToolSuppressShow -ErrorAction SilentlyContinue
      }
    } catch {}
  }
}

$window.Add_Closed($cleanupAction)

try {
  [void]$app.Run($window)
} finally {
  & $cleanupAction
}
