function Get-ScriptDirectory {
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

$scriptDir = Get-ScriptDirectory
$moduleRoot = Join-Path $scriptDir 'NewAssetTool'
$manifest = Join-Path $moduleRoot 'NewAssetTool.psd1'
if (-not (Test-Path $manifest)) {
  throw "Could not find module manifest at '$manifest'."
}

Import-Module $manifest -Force

Start-NewAssetTool
