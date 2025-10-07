Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Xaml
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName WindowsFormsIntegration

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
$targetChromeScale = 0.8
try {
  $candidate = $null
  if (Test-Path Variable:\global:NewAssetToolTargetChromeScale) {
    $candidate = $global:NewAssetToolTargetChromeScale
  } elseif (Test-Path Variable:\script:NewAssetToolTargetChromeScale) {
    $candidate = $script:NewAssetToolTargetChromeScale
  }
  if ($null -ne $candidate) {
    try { $targetChromeScale = [double]$candidate } catch {}
  }
} catch {}
$applyWpfScale = {
  param(
    [string]$Source = 'unspecified',
    [double]$dpiX = 0,
    [double]$dpiY = 0
  )

  if (-not $global:NewAssetToolPerMonitorDpiContextEnabled) { return }
  if (-not $rootScaleTransform) {
    Write-Verbose "[DPI][WPF] Root scale transform not located; skipping scale." -Verbose
    return
  }

  $effectiveTarget = $targetChromeScale
  if (-not $effectiveTarget -or [double]::IsNaN([double]$effectiveTarget) -or $effectiveTarget -le 0) {
    $effectiveTarget = 1.0
  }

  $contextDescription = 'unknown'
  try {
    if (Get-Command Get-NewAssetToolDpiContextDescription -ErrorAction SilentlyContinue) {
      $contextDescription = Get-NewAssetToolDpiContextDescription
    }
  } catch {}

  try {
    if ($dpiX -le 0 -or $dpiY -le 0) {
      try {
        if ($window) {
          $presentationSource = [System.Windows.PresentationSource]::FromVisual($window)
          if ($presentationSource -and $presentationSource.CompositionTarget) {
            if ($dpiX -le 0) { $dpiX = 96.0 * $presentationSource.CompositionTarget.TransformToDevice.M11 }
            if ($dpiY -le 0) { $dpiY = 96.0 * $presentationSource.CompositionTarget.TransformToDevice.M22 }
          }
        }
      } catch {}
    }

    if ($dpiX -le 0) { $dpiX = 96.0 }
    if ($dpiY -le 0) { $dpiY = 96.0 }

    $scaleX = $effectiveTarget
    $scaleY = $effectiveTarget

    $dpiScaleX = $dpiX / 96.0
    $dpiScaleY = $dpiY / 96.0
    if ($dpiScaleX -gt 0) { $scaleX = $effectiveTarget / $dpiScaleX }
    if ($dpiScaleY -gt 0) { $scaleY = $effectiveTarget / $dpiScaleY }

    $rootScaleTransform.ScaleX = $scaleX
    $rootScaleTransform.ScaleY = $scaleY
    Write-Verbose (
      "[DPI][WPF] Applied root layout scale X={0:n3}, Y={1:n3} (dpiX={2:n1}, dpiY={3:n1}) ({4}) context={5}" -f
      $scaleX, $scaleY, $dpiX, $dpiY, $Source, $contextDescription
    ) -Verbose
  } catch {
    Write-Verbose "[DPI][WPF] Failed to apply layout scale ({0}): $($_.Exception.Message)" -Verbose
  }
}

if ($global:NewAssetToolPerMonitorDpiContextEnabled -and $window) {
  try {
    $window.Add_DpiChanged({
      param($sender,$eventArgs)
      & $applyWpfScale 'Window.DpiChanged' $eventArgs.NewDpi.PixelsPerInchX $eventArgs.NewDpi.PixelsPerInchY
    })
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
