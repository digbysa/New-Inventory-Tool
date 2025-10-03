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

$script:__newAssetToolDesignDpi = $form.AutoScaleDimensions
if (-not $script:__newAssetToolDesignDpi -or $script:__newAssetToolDesignDpi.Width -le 0 -or $script:__newAssetToolDesignDpi.Height -le 0) {
  $script:__newAssetToolDesignDpi = New-Object System.Drawing.SizeF(96.0, 96.0)
}
$script:__newAssetToolLastDpi = $script:__newAssetToolDesignDpi

function Invoke-NewAssetToolWinFormsDpiUpdate {
  param(
    [double]$DpiX,
    [double]$DpiY
  )

  if (-not $form) { return }

  if ($DpiX -le 0) { $DpiX = $script:__newAssetToolDesignDpi.Width }
  if ($DpiY -le 0) { $DpiY = $script:__newAssetToolDesignDpi.Height }

  $previous = $script:__newAssetToolLastDpi
  if (-not $previous -or $previous.Width -le 0 -or $previous.Height -le 0) {
    $previous = $script:__newAssetToolDesignDpi
  }

  try { $form.SuspendLayout() } catch {}
  try { $form.AutoScaleDimensions = New-Object System.Drawing.SizeF($DpiX, $DpiY) } catch {}
  try { $form.PerformAutoScale() } catch {}

  $scaleX = if ($previous.Width -gt 0) { $DpiX / $previous.Width } else { 1 }
  $scaleY = if ($previous.Height -gt 0) { $DpiY / $previous.Height } else { 1 }
  if ($scaleX -le 0) { $scaleX = 1 }
  if ($scaleY -le 0) { $scaleY = 1 }

  try { $form.Scale($scaleX, $scaleY) } catch {}
  try { $form.ResumeLayout($true) } catch {}

  $script:__newAssetToolLastDpi = New-Object System.Drawing.SizeF($DpiX, $DpiY)
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
$window.Add_SourceInitialized({
  param($sender, $eventArgs)

  $dpiX = 96.0
  $dpiY = 96.0
  try {
    $hwndSource = [System.Windows.Interop.HwndSource]::FromVisual($sender)
    if ($hwndSource -and $hwndSource.CompositionTarget) {
      $transform = $hwndSource.CompositionTarget.TransformToDevice
      if ($transform) {
        $dpiX = 96.0 * $transform.M11
        $dpiY = 96.0 * $transform.M22
      }
    }
  } catch {}

  try { Invoke-NewAssetToolWinFormsDpiUpdate -DpiX $dpiX -DpiY $dpiY } catch {}
})
$window.Add_DpiChanged({
  param($sender, $args)

  $dpiX = $null
  $dpiY = $null

  try {
    if ($args -and $args.NewDpi) {
      $dpiX = $args.NewDpi.PixelsPerInchX
      $dpiY = $args.NewDpi.PixelsPerInchY
    }
  } catch {}

  if (-not $dpiX -or $dpiX -le 0 -or -not $dpiY -or $dpiY -le 0) {
    try {
      $hwndSource = [System.Windows.Interop.HwndSource]::FromVisual($sender)
      if ($hwndSource -and $hwndSource.CompositionTarget) {
        $transform = $hwndSource.CompositionTarget.TransformToDevice
        if ($transform) {
          $dpiX = 96.0 * $transform.M11
          $dpiY = 96.0 * $transform.M22
        }
      }
    } catch {}
  }

  try { Invoke-NewAssetToolWinFormsDpiUpdate -DpiX $dpiX -DpiY $dpiY } catch {}
})
$null = $window.Dispatcher.BeginInvoke(
  [System.Action]{
    $dpiX = 96.0
    $dpiY = 96.0
    try {
      $presentationSource = [System.Windows.Media.PresentationSource]::FromVisual($window)
      if ($presentationSource -and $presentationSource.CompositionTarget) {
        $transform = $presentationSource.CompositionTarget.TransformToDevice
        if ($transform) {
          $dpiX = 96.0 * $transform.M11
          $dpiY = 96.0 * $transform.M22
        }
      }
    } catch {}

    try { Invoke-NewAssetToolWinFormsDpiUpdate -DpiX $dpiX -DpiY $dpiY } catch {}
  },
  [System.Windows.Threading.DispatcherPriority]::Loaded)
$form.Visible = $true

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
