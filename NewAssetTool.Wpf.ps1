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

function Get-FirstWinFormsTabControl {
  param([System.Windows.Forms.Control]$parent)
  if (-not $parent) { return $null }
  if ($parent -is [System.Windows.Forms.TabControl]) { return $parent }
  foreach ($child in $parent.Controls) {
    $found = Get-FirstWinFormsTabControl -parent $child
    if ($found) { return $found }
  }
  return $null
}

$topTabs = $window.FindName('TopTabs')
$embeddedTabControl = Get-FirstWinFormsTabControl -parent $form
if ($topTabs -and $embeddedTabControl) {
  try { $embeddedTabControl.Appearance = [System.Windows.Forms.TabAppearance]::Buttons } catch {}
  try { $embeddedTabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed } catch {}
  try { $embeddedTabControl.ItemSize = New-Object System.Drawing.Size -ArgumentList 0,1 } catch {}
  try { $embeddedTabControl.Padding = New-Object System.Drawing.Point -ArgumentList 0,0 } catch {}
  try { $embeddedTabControl.TabStop = $false } catch {}

  $script:__tabSyncInProgress = $false

  $syncAction = {
    param($setWpf,$index)
    if ($setWpf) {
      if ($index -ge 0 -and $index -lt $topTabs.Items.Count) {
        if ($topTabs.SelectedIndex -ne $index) { $topTabs.SelectedIndex = $index }
      }
    } else {
      if ($index -ge 0 -and $index -lt $embeddedTabControl.TabPages.Count) {
        if ($embeddedTabControl.SelectedIndex -ne $index) { $embeddedTabControl.SelectedIndex = $index }
      }
    }
  }

  $setIndexSafely = {
    param($isWpf, $index)
    if ($script:__tabSyncInProgress) { return }
    $script:__tabSyncInProgress = $true
    try {
      & $syncAction $isWpf $index
    } catch {}
    $script:__tabSyncInProgress = $false
  }

  try {
    & $setIndexSafely $true $embeddedTabControl.SelectedIndex
  } catch {}

  $topTabs.Add_SelectionChanged({
    param($sender,$args)
    if (-not $sender) { return }
    & $setIndexSafely $false $sender.SelectedIndex
  })

  $embeddedTabControl.add_SelectedIndexChanged({
    param($sender,$eventArgs)
    if (-not $sender) { return }
    & $setIndexSafely $true $sender.SelectedIndex
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
