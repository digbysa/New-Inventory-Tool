Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Xaml
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName WindowsFormsIntegration
Add-Type -Namespace Win32 -Name DpiAwarenessHelper -MemberDefinition @"
using System;
using System.Runtime.InteropServices;

public static class DpiAwarenessHelper
{
    private static readonly IntPtr PerMonitorV2Context = new IntPtr(-4);

    [DllImport("user32.dll")]
    private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiFlag);

    [DllImport("shcore.dll")]
    private static extern int SetProcessDpiAwareness(PROCESS_DPI_AWARENESS value);

    private enum PROCESS_DPI_AWARENESS
    {
        PROCESS_DPI_UNAWARE = 0,
        PROCESS_SYSTEM_DPI_AWARE = 1,
        PROCESS_PER_MONITOR_DPI_AWARE = 2
    }

    public static bool TrySetPerMonitorV2()
    {
        try
        {
            if (SetProcessDpiAwarenessContext(PerMonitorV2Context))
            {
                return true;
            }
        }
        catch (EntryPointNotFoundException) { }
        catch (DllNotFoundException) { }

        try
        {
            return SetProcessDpiAwareness(PROCESS_DPI_AWARENESS.PROCESS_PER_MONITOR_DPI_AWARE) == 0;
        }
        catch (EntryPointNotFoundException) { }
        catch (DllNotFoundException) { }

        return false;
    }
}
"@

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

$dpiAware = [Win32.DpiAwarenessHelper]::TrySetPerMonitorV2()
if (-not $dpiAware) {
  Write-Warning 'Failed to enable Per-Monitor V2 DPI awareness via Win32 APIs. High-DPI behavior may be degraded.'
}

if (-not (Test-Path $ps1Path)) {
  throw "Could not find NewAssetTool.ps1 at '$ps1Path'."
}
if (-not (Test-Path $xamlPath)) {
  throw "Could not find NewAssetTool.xaml at '$xamlPath'."
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
