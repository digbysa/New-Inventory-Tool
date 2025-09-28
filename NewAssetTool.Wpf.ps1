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

$host = $window.FindName('WinFormsHost')
if (-not $host) {
  throw "Could not locate the WindowsFormsHost named 'WinFormsHost' in XAML."
}
$host.Child = $form
$form.Visible = $true

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
