Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
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
$script:LocationRows = @()
$script:RoundingByAssetTag = @{} 
$script:CurrentDisplay = $null
$script:CurrentParent  = $null
$script:PreviewCandidate = $null
$script:PreviewIsLinkable = $false
$script:editing = $false
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
  if($u -match '^(AO)-?(.+)$'){
    [void]$out.Add(('{0}-{1}' -f $matches[1],$matches[2]))
    [void]$out.Add(('{0}{1}'  -f $matches[1],$matches[2]))
  }
  return $out
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
  $m = [regex]::Match($po, '(RITM\d+)')
  if($m.Success){ return $m.Groups[1].Value } else { return $po.Trim() }
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
function Build-Indices {
  $script:IndexByAsset.Clear(); $script:IndexBySerial.Clear(); $script:IndexByName.Clear()
  $script:ComputerByAsset.Clear(); $script:ComputerByName.Clear()
  foreach($rec in $script:Computers){
    if($rec.asset_tag){
      $rawKey = $rec.asset_tag.Trim().ToUpper()
      $canon  = Canonical-Asset $rawKey
      $script:IndexByAsset[$rawKey] = $rec
      if($canon){ $script:IndexByAsset[$canon] = $rec }
      $script:ComputerByAsset[$rawKey] = $rec
      if($canon){ $script:ComputerByAsset[$canon] = $rec }
    }
    if($rec.serial_number){ $script:IndexBySerial[$rec.serial_number.ToUpper()] = $rec }
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
        $rawKey = $rec.asset_tag.Trim().ToUpper()
        $canon  = Canonical-Asset $rawKey
        $script:IndexByAsset[$rawKey] = $rec
        if($canon){ $script:IndexByAsset[$canon] = $rec }
      }
      if($rec.serial_number){ $script:IndexBySerial[$rec.serial_number.ToUpper()] = $rec }
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
  $kids = New-Object System.Collections.ArrayList
  if(-not $parentRec){ return $kids }
  $parATKey = (Canonical-Asset $parentRec.asset_tag)
  if([string]::IsNullOrWhiteSpace($parATKey)){ return $kids }
  # Direct children by canonical AssetTag
  if($script:ChildrenByParent.ContainsKey($parATKey)){
    foreach($ch in $script:ChildrenByParent[$parATKey]){ if(-not $kids.Contains($ch)){ [void]$kids.Add($ch) } }
  }
  # Hostname/serial token matching (direct links)
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
      if($matchHost -and -not $kids.Contains($rec)){ [void]$kids.Add($rec) }
    }
  }
  # Pull in grandchildren under any Cart child (Mic/Scanner only)
  $cartKids = @($kids | Where-Object { $_.Type -eq 'Cart' })
  foreach($cart in $cartKids){
    $cartKey = Canonical-Asset $cart.asset_tag
    if($cartKey -and $script:ChildrenByParent.ContainsKey($cartKey)){
      foreach($gch in $script:ChildrenByParent[$cartKey]){
        if(($gch.Type -eq 'Mic' -or $gch.Type -eq 'Scanner') -and -not $kids.Contains($gch)){ [void]$kids.Add($gch) }
      }
    }
    foreach($tbl in @('Mics','Scanners')){
      $collection = (Get-Variable -Scope Script -Name $tbl -ErrorAction SilentlyContinue).Value
      if(-not $collection){ continue }
      foreach($rec in $collection){
        if(-not $rec.u_parent_asset){ continue }
        $upa = $rec.u_parent_asset.Trim().ToUpper()
        $cartNameU = if($cart.name){ $cart.name.Trim().ToUpper() } else { '' }
        if($cartNameU -and $upa -eq $cartNameU){ if(-not $kids.Contains($rec)){ [void]$kids.Add($rec) } }
      }
    }
  }
  return $kids
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
  Ensure-PrevPropParentLabel
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
function Load-DepartmentMaster(){
  try{
    $script:DepartmentMaster = @()
    $script:DepartmentUserAdds = @()
    $script:DepartmentList = New-Object System.Collections.Generic.List[string]
    $script:DepartmentListNorm = New-Object System.Collections.Generic.HashSet[string]
    $fileD = $null; try { if($script:DataFolder){ $fileD = Join-Path $script:DataFolder 'DepartmentMaster.csv' } } catch {}
    if(-not $fileD -or -not (Test-Path $fileD)){
      try { $fileD = Join-Path $PSScriptRoot 'Data/DepartmentMaster.csv' } catch {}
    }
    if(-not $fileD -or -not (Test-Path $fileD)){
      try { $fileD = 'DepartmentMaster.csv' } catch {}
    }
    if(Test-Path $fileD){
      $script:DepartmentMaster = Import-Csv -Path $fileD
      foreach($row in $script:DepartmentMaster){
        $d = Normalize-Field $row.Department
        if($d -and -not $script:DepartmentListNorm.Contains($d)){
          [void]$script:DepartmentList.Add($row.Department)
          [void]$script:DepartmentListNorm.Add($d)
        }
      }
    }
    $fileDU = $null; try { if($script:DataFolder){ $fileDU = Join-Path $script:DataFolder 'DepartmentMaster-UserAdds.csv' } } catch {}
    if(-not $fileDU -or -not (Test-Path $fileDU)){
      try { $fileDU = Join-Path $PSScriptRoot 'Data/DepartmentMaster-UserAdds.csv' } catch {}
    }
    if(-not $fileDU -or -not (Test-Path $fileDU)){
      try { $fileDU = 'DepartmentMaster-UserAdds.csv' } catch {}
    }
    if(Test-Path $fileDU){
      $script:DepartmentUserAdds = Import-Csv -Path $fileDU
      foreach($row in $script:DepartmentUserAdds){
        $d = Normalize-Field $row.Department
        if($d -and -not $script:DepartmentListNorm.Contains($d)){
          [void]$script:DepartmentList.Add($row.Department)
          [void]$script:DepartmentListNorm.Add($d)
        }
      }
    }
    $script:DepartmentList = ($script:DepartmentList | Sort-Object -Unique)
  } catch { }
}
function Save-DepartmentUserAdd([string]$dept){
  try{
    if([string]::IsNullOrWhiteSpace($dept)){ return }
    $n = Normalize-Field $dept
    if($script:DepartmentListNorm.Contains($n)){ return }
    $fileDU = $null; try { if($script:DataFolder){ $fileDU = Join-Path $script:DataFolder 'DepartmentMaster-UserAdds.csv' } } catch {}
    if(-not $fileDU -or -not (Test-Path $fileDU)){
      try { $fileDU = Join-Path $PSScriptRoot 'Data/DepartmentMaster-UserAdds.csv' } catch {}
    }
    if(-not $fileDU -or -not (Test-Path $fileDU)){
      try { $fileDU = 'DepartmentMaster-UserAdds.csv' } catch {}
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
    if(-not $cmbDept){ return }
    $cmbDept.Items.Clear()
    if(-not $script:DepartmentList){ Load-DepartmentMaster }
    if($script:DepartmentList){
      [void]$cmbDept.Items.AddRange($script:DepartmentList)
    }
    if($current){
      $cmbDept.Text = $current
    } elseif($cmbDept.Items.Count -gt 0){
      $cmbDept.SelectedIndex = 0
    }
  } catch {}
}
function Rebuild-RoomCaches(){
  $rooms = New-Object System.Collections.Generic.List[string]
  $codes = New-Object System.Collections.Generic.List[string]
  foreach($row in $script:LocationRows){
    $raw = Get-LocVal $row 'Room'
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
  $script:LocationRows=@()
  $lm = Join-Path $folder 'LocationMaster.csv'
  if(Test-Path $lm){ $script:LocationRows += Import-Csv $lm }
  $lm2 = Join-Path $folder 'LocationMaster-UserAdds.csv'
  if(Test-Path $lm2){ $script:LocationRows += Import-Csv $lm2 }
  $script:LocCols = @{}
  Rebuild-RoomCaches
}
function Save-LocationUserAdd([string]$city,[string]$loc,[string]$b,[string]$f,[string]$r){
  if(-not $script:OutputFolder){ $script:OutputFolder = $script:DataFolder }
  $file = Join-Path $script:OutputFolder 'LocationMaster-UserAdds.csv'
  if(-not (Test-Path $file)){ 'City,Location,Building,Floor,Room' | Out-File -FilePath $file -Encoding UTF8 }
  ('"{0}","{1}","{2}","{3}","{4}"' -f $city,$loc,$b,$f,$r) | Add-Content -Path $file -Encoding UTF8
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
$H_ASSOC_PREVIEW    = 74
$MAX_ASSOC_GRID_ROWS = 4
$form = New-Object System.Windows.Forms.Form
$form.Text = "Inventory Assoc Finder - OMI"
$lblPaths = New-Object System.Windows.Forms.Label
$lblPaths.AutoSize = $true
$lblPaths.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$lblPaths.TextAlign = "MiddleRight"
$lblPaths.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 720), ($form.ClientSize.Height - 20))
$lblPaths.Size = New-Object System.Drawing.Size(560, 18)
$form.Controls.Add($lblPaths)
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
# Row 2: Scan box
$grpScan = New-Object System.Windows.Forms.GroupBox
$grpScan.Text="Scan / enter Name, SN# or Asset Tag"
$grpScan.Dock='Top'
$grpScan.Margin = '0,6,0,0'
$grpScan.Padding = '8,8,8,8'
$grpScan.AutoSize = $false
$grpScan.Height = 56
$txtScan = New-Object System.Windows.Forms.TextBox
$txtScan.Location='16,22'
$txtScan.Anchor='Top,Left,Right'
$txtScan.Size='900,24'
$btnLookup = New-Object System.Windows.Forms.Button
$btnLookup.Text="Lookup"
$btnLookup.Location='930,20'
$btnLookup.Anchor='Top,Right'
$btnLookup.Size='110,26'
$grpScan.Controls.AddRange(@($txtScan,$btnLookup))
# Add SCAN first, then TOP row second so the TOP row ends up above the scan row
$panelTop.Controls.Add($grpScan)
$panelTop.Controls.Add($flpTop)
# ---------- END HEADER ----------
# Main 2-col table
$tlpMain = New-Object System.Windows.Forms.TableLayoutPanel
$LEFT_COL_WIDTH  = 520
$RIGHT_COL_WIDTH = 640
$tlpMain.Dock = 'Fill'
$tlpMain.ColumnCount = 2
$tlpMain.RowCount = 1
$tlpMain.Padding = New-Object System.Windows.Forms.Padding($GAP, 0, $GAP, $GAP)
$tlpMain.BackColor = [System.Drawing.Color]::White
$tlpMain.ColumnStyles.Clear()
$tlpMain.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, $LEFT_COL_WIDTH)) )
$tlpMain.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, $RIGHT_COL_WIDTH)) )
$tlpMain.RowStyles.Clear()
$tlpMain.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)) )
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
$grpSummary = New-Object System.Windows.Forms.GroupBox; $grpSummary.Text="Device Summary"; $grpSummary.Dock='Fill'
$grpSummary.Margin = New-Object System.Windows.Forms.Padding($GAP)
$grpSummary.Padding = New-Object System.Windows.Forms.Padding($GAP)
$grpSummary.Controls.AddRange(@(
  (New-L "Detected Type:" 12 28),
  (New-L "Name:" 12 58),
  (New-L "Asset Tag:" 12 88),
  (New-L "Serial:" 12 118),
  (New-L "Parent:" 12 148),
  (New-L "PO RITM:" 12 178),
  (New-L "Retire Date:" 12 208),
  (New-L "Last Rounded:" 12 238)
$txtType=New-RO 120 25 340
))
$txtHost=New-RO 120 55 340
$txtAT=New-RO 120 85 340
$txtSN=New-RO 120 115 340
$txtParent=New-RO 120 145 340
$txtRITM=New-RO 120 175 340
$txtRetire=New-RO 120 205 340
$txtRound=New-RO 120 235 340
$grpSummary.Controls.AddRange(@($txtType,$txtHost,$txtAT,$txtSN,$txtParent,$txtRITM,$txtRetire,$txtRound))
# Fix button beside "Name:"
$btnFixName = New-Object System.Windows.Forms.Button
$btnFixName.Text = "Fix"
$btnFixName.Size = '45,24'
$btnFixName.Location = New-Object System.Drawing.Point(70, 55)
$btnFixName.Anchor = 'Top,Left'
$btnFixName.Enabled = $false
$grpSummary.Controls.Add($btnFixName)
# Device Location (with City)
$grpLoc = New-Object System.Windows.Forms.GroupBox; $grpLoc.Text="Device Location"; $grpLoc.Dock='Fill'
$grpLoc.Margin = New-Object System.Windows.Forms.Padding($GAP)
$grpLoc.Padding = New-Object System.Windows.Forms.Padding($GAP)
$grpLoc.Controls.AddRange(@(
  (New-L "City:" 12 28),
  (New-L "Location:" 12 58),
  (New-L "Building:" 12 88),
  (New-L "Floor:" 12 118),
  (New-L "Room:" 12 148),
  (New-L "Department:" 12 178)
))
if(-not $btnEditLoc){
  $btnEditLoc = New-Object System.Windows.Forms.Button
  $btnEditLoc.Text = 'Edit Location'
  $btnEditLoc.Size = '120,26'
  $grpLoc.Controls.Add($btnEditLoc)
}
$txtCity=New-RO 120 25 340
$txtLocation=New-RO 120 55 340
$txtBldg=New-RO 120 85 340
$txtFloor=New-RO 120 115 340
$txtRoom=New-RO 120 145 340
$txtDept=New-RO 120 175 340; $txtDept.Visible=$false
$txtDept.Visible = $true
$txtDept.Visible=$true
$btnEditLoc.Location = New-Object System.Drawing.Point(340,195)
$grpLoc.Controls.AddRange(@($txtCity,$txtLocation,$txtBldg,$txtFloor,$txtRoom,$txtDept,$btnEditLoc))
# Editable combos
$cmbCity=New-Object System.Windows.Forms.ComboBox; $cmbCity.Location='120,25'; $cmbCity.Size='360,24'; $cmbCity.Visible=$false; $cmbCity.DropDownStyle='DropDown'
$cmbLocation=New-Object System.Windows.Forms.ComboBox; $cmbLocation.Location='120,55'; $cmbLocation.Size='360,24'; $cmbLocation.Visible=$false; $cmbLocation.DropDownStyle='DropDown'
$cmbBuilding=New-Object System.Windows.Forms.ComboBox; $cmbBuilding.Location='120,85'; $cmbBuilding.Size='360,24'; $cmbBuilding.Visible=$false; $cmbBuilding.DropDownStyle='DropDown'
$cmbFloor=New-Object System.Windows.Forms.ComboBox; $cmbFloor.Location='120,115'; $cmbFloor.Size='360,24'; $cmbFloor.Visible=$false; $cmbFloor.DropDownStyle='DropDown'
$cmbRoom=New-Object System.Windows.Forms.ComboBox; $cmbRoom.Location='120,145'; $cmbRoom.Size='360,24'; $cmbRoom.Visible=$false; $cmbRoom.DropDownStyle='DropDown'
$cmbDept=New-Object System.Windows.Forms.ComboBox; $cmbDept.Location='120,175'; $cmbDept.Width=360; $cmbDept.Visible=$false; $cmbDept.DropDownStyle='DropDownList'
$cmbDept.Visible = $false
$grpLoc.Controls.AddRange(@($cmbCity,$cmbLocation,$cmbBuilding,$cmbFloor,$cmbRoom,$cmbDept))
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
# Associated Devices
$grpAssoc = New-Object System.Windows.Forms.GroupBox; $grpAssoc.Text="Associated Devices"; $grpAssoc.Dock='Fill'
$grpAssoc.Margin = New-Object System.Windows.Forms.Padding($GAP)
$grpAssoc.Padding = New-Object System.Windows.Forms.Padding($GAP)
$tabAssoc = New-Object System.Windows.Forms.TabControl; $tabAssoc.Dock='Fill'
$tabGrid = New-Object System.Windows.Forms.TabPage; $tabGrid.Text='Grid'
$tabCards = New-Object System.Windows.Forms.TabPage; $tabCards.Text='Cards'
$tabAssoc.TabPages.AddRange(@($tabGrid,$tabCards))
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Dock='Fill'; $dgv.AutoGenerateColumns=$false; $dgv.AllowUserToAddRows=$false; $dgv.ReadOnly=$true
$dgv.SelectionMode=[System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgv.MultiSelect=$true; $dgv.RowHeadersVisible=$false; $dgv.BackgroundColor=[System.Drawing.Color]::White; $dgv.BorderStyle='FixedSingle'
$dgv.AutoSizeColumnsMode='DisplayedCells'
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
$dgv.Columns.Add((New-TextCol 'Retire' 'Retire  120')) | Out-Null
$tabGrid.Controls.Add($dgv)
try{ $dgv.Columns['Name'].AutoSizeMode='Fill'; $dgv.Columns['RITM'].Width=120 } catch {}
$cards = New-Object System.Windows.Forms.FlowLayoutPanel
$cards.Dock='Fill'; $cards.AutoScroll=$true; $cards.WrapContents=$true; $cards.FlowDirection='LeftToRight'
$tabCards.Controls.Add($cards)
$tlpAssoc = New-Object System.Windows.Forms.TableLayoutPanel
$tlpAssoc.Dock = 'Fill'
$tlpAssoc.ColumnCount = 1
$tlpAssoc.RowCount = 2
$tlpAssoc.RowStyles.Clear()
$tlpAssoc.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)) )
$tlpAssoc.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle -ArgumentList ([System.Windows.Forms.SizeType]::AutoSize)) )
$tlpAssoc.Padding = New-Object System.Windows.Forms.Padding($GAP)
$tlpAssoc.Margin  = New-Object System.Windows.Forms.Padding($GAP)
$tlpAssoc.Controls.Add($dgv,0,0)
# Toolbar row
$lblAdd = New-Object System.Windows.Forms.Label
$lblAdd.Text = "Add Peripheral (AssetTag/Serial):"
$lblAdd.AutoSize = $true
$lblAdd.Margin   = '6,8,6,6'
$txtAdd = New-Object System.Windows.Forms.TextBox
$txtAdd.Dock   = 'Fill'
$txtAdd.Margin = '0,4,6,4'
$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text   = 'Remove Selected'
$btnRemove.Dock   = 'Right'
$btnRemove.Margin = '0,4,6,4'
$btnRemove.AutoSize = $true
$btnRemove.AutoSizeMode = 'GrowAndShrink'
$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text   = 'Link to Computer'
$btnAdd.Dock   = 'Right'
$btnAdd.Margin = '0,4,6,4'
$btnAdd.AutoSize = $true
$btnAdd.AutoSizeMode = 'GrowAndShrink'
$tlpAssocTop = New-Object System.Windows.Forms.TableLayoutPanel
$tlpAssocTop.Dock = 'Fill'
$tlpAssocTop.RowCount = 1
$tlpAssocTop.ColumnCount = 4
$tlpAssocTop.Padding = '0,0,0,0'
$tlpAssocTop.Margin  = '0,0,0,0'
$tlpAssocTop.RowStyles.Clear()
$tlpAssocTop.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)) )
$tlpAssocTop.ColumnStyles.Clear()
$tlpAssocTop.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle -ArgumentList ([System.Windows.Forms.SizeType]::AutoSize)) )
$tlpAssocTop.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)) )
$tlpAssocTop.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle -ArgumentList ([System.Windows.Forms.SizeType]::AutoSize)) )
$tlpAssocTop.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle -ArgumentList ([System.Windows.Forms.SizeType]::AutoSize)) )
$tlpAssocTop.Controls.Add($lblAdd, 0, 0)
$tlpAssocTop.Controls.Add($txtAdd, 1, 0)
$tlpAssocTop.Controls.Add($btnRemove, 2, 0)
$tlpAssocTop.Controls.Add($btnAdd, 3, 0)
if(-not $valPrevSN){ $valPrevSN = New-Object System.Windows.Forms.Label; $valPrevSN.AutoSize=$true }
# Preview group (3x3 grid)
# Peripheral Preview (3 cols x 4 rows, inline "Label: value")
# Peripheral Preview (6 cols x 4 rows) with visible grid lines and equal columns
# Peripheral Preview (6x4) with separate label & value cells and visible borders
# Peripheral Preview (6x4) separate label & value cells, AutoSize columns, visible borders
# === Peripheral Preview (5 x 5) — label/value in separate cells ===
$grpPrev = New-Object System.Windows.Forms.GroupBox
$grpPrev.AutoSize=$true
$grpPrev.AutoSizeMode=[System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$grpPrev.Dock='Top'
$grpPrev.Text = 'Peripheral Preview'
# keep the preview height bound to its content instead of filling remaining space
$grpPrev.Dock = 'Top'
$grpPrev.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
    [System.Windows.Forms.AnchorStyles]::Left -bor `
    [System.Windows.Forms.AnchorStyles]::Right
# tighten the bottom margin so the next section sits closer
$grpPrev.Margin = New-Object System.Windows.Forms.Padding(8,8,8,2)
# Ensure all labels/values exist and use the requested captions
if(-not $lblPrevType){ $lblPrevType = New-Object System.Windows.Forms.Label }
$lblPrevType.Text='Type:';            $lblPrevType.AutoSize=$true
if(-not $valPrevType){ $valPrevType = New-Object System.Windows.Forms.Label }
$valPrevType.AutoSize=$true
if(-not $lblPrevName){ $lblPrevName = New-Object System.Windows.Forms.Label }
$lblPrevName.Text='Name:';            $lblPrevName.AutoSize=$true
if(-not $valPrevName){ $valPrevName = New-Object System.Windows.Forms.Label }
$valPrevName.AutoSize=$true
if(-not $valPrevProposed){ $valPrevProposed = New-Object System.Windows.Forms.Label }
$valPrevProposed.AutoSize=$true
if(-not $lblPrevParent){ $lblPrevParent = New-Object System.Windows.Forms.Label }
$lblPrevParent.Text='Parent:';        $lblPrevParent.AutoSize=$true
if(-not $valPrevParent){ $valPrevParent = New-Object System.Windows.Forms.Label }
$valPrevParent.AutoSize=$true
if(-not $valPrevPropParent){ $valPrevPropParent = New-Object System.Windows.Forms.Label }
$valPrevPropParent.AutoSize=$true
if(-not $lblPrevAT){ $lblPrevAT = New-Object System.Windows.Forms.Label }
$lblPrevAT.Text='Asset Tag:';         $lblPrevAT.AutoSize=$true
if(-not $valPrevAT){ $valPrevAT = New-Object System.Windows.Forms.Label }
$valPrevAT.AutoSize=$true
if(-not $lblPrevSN){ $lblPrevSN = New-Object System.Windows.Forms.Label }
$lblPrevSN.Text='Serial Number:';     $lblPrevSN.AutoSize=$true
if(-not $valPrevSN){ $valPrevSN = New-Object System.Windows.Forms.Label }
$valPrevSN.AutoSize=$true
if(-not $lblPrevRITM){ $lblPrevRITM = New-Object System.Windows.Forms.Label }
$lblPrevRITM.Text='RITM:';            $lblPrevRITM.AutoSize=$true
if(-not $valPrevRITM){ $valPrevRITM = New-Object System.Windows.Forms.Label }
$valPrevRITM.AutoSize=$true
if(-not $lblPrevRetire){ $lblPrevRetire = New-Object System.Windows.Forms.Label }
$lblPrevRetire.Text='Retire:';        $lblPrevRetire.AutoSize=$true
if(-not $valPrevRetire){ $valPrevRetire = New-Object System.Windows.Forms.Label }
$valPrevRetire.AutoSize=$true
# arrow labels for rows 1 and 2
if(-not $lblArrow1){ $lblArrow1 = New-Object System.Windows.Forms.Label }
$lblArrow1.Text='------>'; $lblArrow1.AutoSize=$true
if(-not $lblArrow2){ $lblArrow2 = New-Object System.Windows.Forms.Label }
$lblArrow2.Text='------>'; $lblArrow2.AutoSize=$true
# Consistent margins (trim vertical space); give values a small right pad
$lblMargin = New-Object System.Windows.Forms.Padding(0,0,6,0)
$valMargin = New-Object System.Windows.Forms.Padding(0,0,10,0)
$lblPrevType.Margin     = $lblMargin
$lblPrevName.Margin     = $lblMargin
$lblPrevParent.Margin   = $lblMargin
$lblPrevAT.Margin       = $lblMargin
$lblPrevSN.Margin       = $lblMargin
$lblPrevRITM.Margin     = $lblMargin
$lblPrevRetire.Margin   = $lblMargin
$lblArrow1.Margin       = New-Object System.Windows.Forms.Padding(4,0,4,0)
$lblArrow2.Margin       = New-Object System.Windows.Forms.Padding(4,0,4,0)
$valPrevType.Margin       = $valMargin
$valPrevName.Margin       = $valMargin
$valPrevProposed.Margin   = $valMargin
$valPrevParent.Margin     = $valMargin
$valPrevPropParent.Margin = $valMargin
$valPrevAT.Margin         = $valMargin
$valPrevSN.Margin         = $valMargin
$valPrevRITM.Margin       = $valMargin
$valPrevRetire.Margin     = $valMargin
# Build the 5x5 table — AutoSize columns/rows; guidelines OFF
$tlpPrev = New-Object System.Windows.Forms.TableLayoutPanel
$tlpPrev.Dock='Top'
$tlpPrev.AutoSize       = $true
$tlpPrev.AutoSizeMode   = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$tlpPrev.Dock           = 'Top'
$tlpPrev.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
    [System.Windows.Forms.AnchorStyles]::Left -bor `
    [System.Windows.Forms.AnchorStyles]::Right
$tlpPrev.CellBorderStyle= [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
$tlpPrev.Padding        = New-Object System.Windows.Forms.Padding(2)
$tlpPrev.Margin         = New-Object System.Windows.Forms.Padding(0)
$tlpPrev.ColumnCount    = 5
$tlpPrev.RowCount       = 5
$tlpPrev.ColumnStyles.Clear()
for($i=0; $i -lt 5; $i++){
    [void]$tlpPrev.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle -ArgumentList ([System.Windows.Forms.SizeType]::AutoSize)) )
}
$tlpPrev.RowStyles.Clear()
for($i=0; $i -lt 5; $i++){
    [void]$tlpPrev.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle -ArgumentList ([System.Windows.Forms.SizeType]::AutoSize)) )
}
# Row 0: Type: | value
$tlpPrev.Controls.Add($lblPrevType); $tlpPrev.SetColumn($lblPrevType,0); $tlpPrev.SetRow($lblPrevType,0)
$tlpPrev.Controls.Add($valPrevType); $tlpPrev.SetColumn($valPrevType,1); $tlpPrev.SetRow($valPrevType,0)
# Row 1: Name: | value | ----> | New name value
$tlpPrev.Controls.Add($lblPrevName);     $tlpPrev.SetColumn($lblPrevName,0);     $tlpPrev.SetRow($lblPrevName,1)
$tlpPrev.Controls.Add($valPrevName);     $tlpPrev.SetColumn($valPrevName,1);     $tlpPrev.SetRow($valPrevName,1)
$tlpPrev.Controls.Add($lblArrow1);       $tlpPrev.SetColumn($lblArrow1,2);       $tlpPrev.SetRow($lblArrow1,1)
$tlpPrev.Controls.Add($valPrevProposed); $tlpPrev.SetColumn($valPrevProposed,4); $tlpPrev.SetRow($valPrevProposed,1)
# Row 2: Parent: | value | ----> | New Parent value
$tlpPrev.Controls.Add($lblPrevParent);     $tlpPrev.SetColumn($lblPrevParent,0);     $tlpPrev.SetRow($lblPrevParent,2)
$tlpPrev.Controls.Add($valPrevParent);     $tlpPrev.SetColumn($valPrevParent,1);     $tlpPrev.SetRow($valPrevParent,2)
$tlpPrev.Controls.Add($lblArrow2);         $tlpPrev.SetColumn($lblArrow2,2);         $tlpPrev.SetRow($lblArrow2,2)
$tlpPrev.Controls.Add($valPrevPropParent); $tlpPrev.SetColumn($valPrevPropParent,4); $tlpPrev.SetRow($valPrevPropParent,2)
# Row 3: Asset Tag: | value | (empty) | RITM: | value
$tlpPrev.Controls.Add($lblPrevAT);    $tlpPrev.SetColumn($lblPrevAT,0);    $tlpPrev.SetRow($lblPrevAT,3)
$tlpPrev.Controls.Add($valPrevAT);    $tlpPrev.SetColumn($valPrevAT,1);    $tlpPrev.SetRow($valPrevAT,3)
$tlpPrev.Controls.Add($lblPrevRITM);  $tlpPrev.SetColumn($lblPrevRITM,3);  $tlpPrev.SetRow($lblPrevRITM,3)
$tlpPrev.Controls.Add($valPrevRITM);  $tlpPrev.SetColumn($valPrevRITM,4);  $tlpPrev.SetRow($valPrevRITM,3)
# Row 4: Serial Number: | value | (empty) | Retire: | value
$tlpPrev.Controls.Add($lblPrevSN);       $tlpPrev.SetColumn($lblPrevSN,0);       $tlpPrev.SetRow($lblPrevSN,4)
$tlpPrev.Controls.Add($valPrevSN);       $tlpPrev.SetColumn($valPrevSN,1);       $tlpPrev.SetRow($valPrevSN,4)
$tlpPrev.Controls.Add($lblPrevRetire);   $tlpPrev.SetColumn($lblPrevRetire,3);   $tlpPrev.SetRow($lblPrevRetire,4)
$tlpPrev.Controls.Add($valPrevRetire);   $tlpPrev.SetColumn($valPrevRetire,4);   $tlpPrev.SetRow($valPrevRetire,4)
$grpPrev.Controls.Clear()
$grpPrev.Controls.Add($tlpPrev)
# === /Peripheral Preview ===
$tlpAssocStrip = New-Object System.Windows.Forms.TableLayoutPanel
$tlpAssocStrip.Dock = 'Fill'
$tlpAssocStrip.AutoSize = $true
$tlpAssocStrip.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$tlpAssocStrip.ColumnCount = 1
$tlpAssocStrip.RowCount = 2
$tlpAssocStrip.RowStyles.Clear()
$tlpAssocStrip.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)) )
$tlpAssocStrip.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
$tlpAssocStrip.Controls.Add($tlpAssocTop,0,0)
$tlpAssocStrip.Controls.Add($grpPrev,   0,1)
$tlpAssoc.Controls.Add($tlpAssocStrip,0,1)
$grpAssoc.Controls.Add($tlpAssoc)
# Rounding group
$grpMaint = New-Object System.Windows.Forms.GroupBox; $grpMaint.Text="Device Rounding"; $grpMaint.Dock='Fill'
$grpMaint.Margin = New-Object System.Windows.Forms.Padding($GAP)
$grpMaint.Padding = New-Object System.Windows.Forms.Padding($GAP)
$lblChkStatus=New-Object System.Windows.Forms.Label; $lblChkStatus.Text="Check Status:"; $lblChkStatus.AutoSize=$true; $lblChkStatus.Location='12,28'
$cmbChkStatus=New-Object System.Windows.Forms.ComboBox; $cmbChkStatus.Location='120,25'; $cmbChkStatus.Size='380,24'; $cmbChkStatus.DropDownStyle='DropDownList'
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
$lblTime=New-Object System.Windows.Forms.Label; $lblTime.Text="Rounding Time (min):"; $lblTime.AutoSize=$true; $lblTime.Location='12,60'
$lblMaintType=New-Object System.Windows.Forms.Label; $lblMaintType.Text='Maintenance Type:'; $lblMaintType.AutoSize=$true; $lblMaintType.Location='260,60'
$cmbMaintType=New-Object System.Windows.Forms.ComboBox; $cmbMaintType.Location='370,58'; $cmbMaintType.Size='200,22'; $cmbMaintType.DropDownStyle='DropDownList'
$cmbMaintType.Items.AddRange(@('Excluded','General Rounding','Mobile Cart','Critical Clinical'))
$numTime=New-Object System.Windows.Forms.NumericUpDown; $numTime.Location='160,58'; $numTime.Size='80,24'; $numTime.Minimum=0; $numTime.Maximum=120; $numTime.Value=3
$chkCable=New-Object System.Windows.Forms.CheckBox; $chkCable.Text="Validate Cable Management"; $chkCable.Location='12,96'; $chkCable.AutoSize=$true
$chkLabels=New-Object System.Windows.Forms.CheckBox; $chkLabels.Text="Ensure monitor appropriately labelled"; $chkLabels.Location='12,124'; $chkLabels.AutoSize=$true
$chkCart=New-Object System.Windows.Forms.CheckBox; $chkCart.Text="Check Physical Cart Is Working"; $chkCart.Location='320,96'; $chkCart.AutoSize=$true
$chkPeriph=New-Object System.Windows.Forms.CheckBox; $chkPeriph.Text="Validate peripherals are connected and working"; $chkPeriph.Location='320,124'; $chkPeriph.AutoSize=$true
$btnCheckComplete=New-Object System.Windows.Forms.Button; $btnCheckComplete.Text="Check Complete"; $btnCheckComplete.Location='12,160'; $btnCheckComplete.Size='180,30'
$btnSave=New-Object System.Windows.Forms.Button; $btnSave.Text="Save Event"; $btnSave.Location='204,160'; $btnSave.Size='180,30'
$btnManualRound=New-Object System.Windows.Forms.Button; $btnManualRound.Text="Manual Round"; $btnManualRound.Location='396,160'; $btnManualRound.Size='180,30'; $btnManualRound.Enabled=$false
$grpMaint.Controls.AddRange(@($lblChkStatus,$cmbChkStatus,$lblTime,$numTime,$lblMaintType,$cmbMaintType,$chkCable,$chkLabels,$chkCart,$chkPeriph,$btnCheckComplete,$btnSave,$btnManualRound))
$grpMaint.Controls.Add($lblMaintType); $grpMaint.Controls.Add($cmbMaintType)
# Compose columns
$tlpRight.Controls.Add($grpAssoc,0,0)
$tlpRight.Controls.Add($grpMaint,0,1)
$tlpMain.Controls.Add($tlpLeft, 0, 0)
$tlpMain.Controls.Add($tlpRight,1, 0)
# StatusStrip
$status=New-Object System.Windows.Forms.StatusStrip
$statusLabel=New-Object System.Windows.Forms.ToolStripStatusLabel; $status.Items.Add($statusLabel) | Out-Null; $statusLabel.Text="Ready"
# Add to form
$form.SuspendLayout()
$form.Controls.Add($tlpMain)   # Fill
$form.Add_Shown({
  foreach($c in @($cmbDept,$cmbDepartment,$ddlDept,$ddlDepartment)){ if ($c) { $c.Visible = $false } }
  foreach($t in @($txtDept,$txtDepartment)){ if ($t) { $t.Visible = $true } }
  if ($cmbDept) { $cmbDept.Visible = $false }
  if ($txtDept) { $txtDept.Visible = $true }
  $lblPaths.Text = "Data: " + $DataFolder + "    |    Output: " + $OutputFolder
})
$form.Controls.Add($panelTop)  # Top
$form.Controls.Add($status)    # Bottom
$form.ResumeLayout($true)
$form.PerformLayout()
# -------- Responsive row sizing (DPI aware) --------
function Apply-ResponsiveHeights {
  try {
    # Fixed heights derived from preferred content sizes, with sensible minimums
    $minSummary  = [Math]::Max($grpSummary.PreferredSize.Height, 280)
    $minLocation = [Math]::Max($grpLoc.PreferredSize.Height, 200)
    $tlpLeft.RowStyles[0].SizeType = [System.Windows.Forms.SizeType]::Absolute
    $tlpLeft.RowStyles[0].Height   = $minSummary
    $tlpLeft.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::Absolute
    $tlpLeft.RowStyles[1].Height   = $minLocation
    $rowsShown = [Math]::Max([Math]::Min($dgv.Rows.Count, $MAX_ASSOC_GRID_ROWS), 1)
    $assocInfo = Size-AssocForRows $rowsShown
    $assocTarget = 0
    if($assocInfo -and $assocInfo.Target){
      $assocTarget = [Math]::Max([int]$assocInfo.Target, 0)
    }
    $stripHeight = $tlpAssocStrip.PreferredSize.Height
    if($stripHeight -le 0){
      $stripHeight = [Math]::Max($grpPrev.PreferredSize.Height, 0) + 34
    }
    $assocPadding = $grpAssoc.Padding.Vertical + $grpAssoc.Margin.Vertical + $tlpAssoc.Margin.Vertical
    $minAssoc   = [Math]::Max($assocTarget + $stripHeight + $assocPadding, 220)
    $minRound   = [Math]::Max($grpMaint.PreferredSize.Height, 220)
    $tlpRight.RowStyles[0].SizeType = [System.Windows.Forms.SizeType]::Absolute
    $tlpRight.RowStyles[0].Height   = $minAssoc
    $tlpRight.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::Absolute
    $tlpRight.RowStyles[1].Height   = $minRound
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
    $target = $gridH + $dgv.Margin.Vertical + $tlpAssoc.Padding.Vertical
    $result.Target = [Math]::Max([int]$target, 0)
    $result.Grid   = [Math]::Max([int]$gridH, 0)
  } catch { }
  return $result
}
function Size-AssocForRows([int]$rows){
  $info = Get-AssocSizing $rows
  try{
    if($info.Target -gt 0){
      $tlpAssoc.RowStyles[0].SizeType = [System.Windows.Forms.SizeType]::Absolute
      $tlpAssoc.RowStyles[0].Height   = $info.Target
      $tlpAssoc.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::AutoSize
      $tlpAssoc.RowStyles[1].Height   = 0
    }
  } catch { }
  return $info
}
$form.Add_Shown({
  Apply-ResponsiveHeights
  Size-AssocForRows([Math]::Min([Math]::Max($dgv.Rows.Count,1), $MAX_ASSOC_GRID_ROWS)) | Out-Null
})
# -------- UI logic ---------
function Update-Counters(){ $locCount = $script:LocationRows.Count; $lblDataStatus.Text = ("Computers: {0} | Monitors: {1} | Mics: {2} | Scanners: {3} | Carts: {4} | Locations: {5}" -f `
    $script:Computers.Count,$script:Monitors.Count,$script:Mics.Count,$script:Scanners.Count,$script:Carts.Count,$locCount) }
function Update-CartCheckbox-State([object]$parentRec){
  $chkCart.Checked = $false; $chkCart.Enabled = $false
  if($parentRec -and $parentRec.name -match '^(?i)AO'){ $chkCart.Enabled = $true }
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
  $row = $script:LocationRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Location')) -eq $nLoc } | Select-Object -First 1
  if($row){ return ([string](Get-LocVal $row 'City')) } else { return '' }
}
function Validate-Location($rec){
  # In edit mode, do not actively validate or repaint to avoid flicker; just reflect current text.
  if($script:editing){ return }
  # Show raw values in UI
  $txtCity.Text     = Get-City-ForLocation $rec.location
$txtDept.Text = $Department  # mirror read-only display; dropdown only when editing
  $txtLocation.Text = $rec.location
  $txtBldg.Text     = $rec.u_building
  $txtFloor.Text    = $rec.u_floor
  $txtRoom.Text     = $rec.u_room
  try{ $cmbDept.Text = $rec.u_department_location } catch {}

  $tip.SetToolTip($txtRoom, "")
  $okC=$false; $okL=$false; $okB=$false; $okF=$false; $okR=$false
  $nLoc  = Normalize-Field $rec.location
  $nBld  = Normalize-Field $rec.u_building
  $nFlr  = Normalize-Field ([string]$rec.u_floor)
  $nRoom = Normalize-Field $rec.u_room
  if($script:LocationRows.Count -gt 0){
    $rowsL = $script:LocationRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Location')) -eq $nLoc }
    $okL   = ($rowsL.Count -gt 0)
    # City based on Location
    if($okL){
      $c = Get-LocVal ($rowsL | Select-Object -First 1) 'City'
      if(-not [string]::IsNullOrWhiteSpace($c)){ $okC = $true; $txtCity.Text = $c }
    }
    $rowsB = @()
    if($okL -and $nBld){
      $rowsB = $rowsL | Where-Object { (Normalize-Field (Get-LocVal $_ 'Building')) -eq $nBld }
      $okB   = ($rowsB.Count -gt 0)
    }
    $rowsF = @()
    if($okB -and $nFlr){
      $rowsF = $rowsB | Where-Object { (Normalize-Field (Get-LocVal $_ 'Floor')) -eq $nFlr }
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
  $txtCity.BackColor     = if($okL){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
  $txtLocation.BackColor = if($okL){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
  $txtBldg.BackColor     = if($okB){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
  $txtFloor.BackColor    = if($okF){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
  $txtRoom.BackColor     = if($okR){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
  try{
    $d=$rec.u_department_location
    $okD = $false
    if($d -and $script:DepartmentListNorm){ $okD = $script:DepartmentListNorm.Contains((Normalize-Field $d)) }
    $cmbDept.BackColor = if($okD){ [System.Drawing.Color]::PaleGreen } else { [System.Drawing.Color]::MistyRose }
  } catch {}
}
function Refresh-AssocGrid($parentRec){
  $dgv.Rows.Clear(); if(-not $parentRec){ Size-AssocForRows(1) | Out-Null; return }
  $prow = $dgv.Rows.Add()
  $dgv.Rows[$prow].Cells['Role'].Value='Parent'
  $dgv.Rows[$prow].Cells['Type'].Value='Computer'
  $dgv.Rows[$prow].Cells['Name'].Value=$parentRec.name
  $dgv.Rows[$prow].Cells['AssetTag'].Value=$parentRec.asset_tag
  $dgv.Rows[$prow].Cells['Serial'].Value=$parentRec.serial_number
  $dgv.Rows[$prow].Cells['RITM'].Value=$parentRec.RITM
  $dgv.Rows[$prow].Cells['Retire'].Value= (Fmt-DateLong $parentRec.Retire)
  $dgv.Rows[$prow].DefaultCellStyle.BackColor=[System.Drawing.Color]::WhiteSmoke
  $kids = Get-ChildrenForParent $parentRec
  foreach($ch in $kids){
    $rowIdx = $dgv.Rows.Add()
    $r = $dgv.Rows[$rowIdx]
    $r.Cells['Role'].Value='Child'
    $r.Cells['Type'].Value=$ch.Type
    if($ch.name){ $r.Cells['Name'].Value = $ch.name } else { $r.Cells['Name'].Value = '' }
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
  Size-AssocForRows([Math]::Min([Math]::Max($dgv.Rows.Count,1), $MAX_ASSOC_GRID_ROWS)) | Out-Null
}
function Make-Card($title,$kvPairs,[System.Drawing.Color]$ritmColor,[bool]$showRITM,[bool]$showRetire,$tagPayload){
  $p = New-Object System.Windows.Forms.Panel
  $p.Width = 280; $p.Height = 160; $p.Margin = '6,6,6,6'
  $p.BackColor = [System.Drawing.Color]::WhiteSmoke; $p.BorderStyle='FixedSingle'
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
      if($rec){ $par = Resolve-ParentComputer $rec; Populate-UI $rec $par }
    }
  })
  $lblTitle = New-Object System.Windows.Forms.Label
  $lblTitle.Text = $title; $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
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
    foreach($ch in $kids){
      $ritm = $ch.RITM
      $col = [System.Drawing.Color]::Black
      if(-not [string]::IsNullOrWhiteSpace($ritm)){
        if($pRITM -and $ritm -eq $pRITM){ $col=[System.Drawing.Color]::ForestGreen } else { $col=[System.Drawing.Color]::IndianRed }
      }
      $showR = ($ch.Type -eq 'Monitor')
      $cards.Controls.Add( (Make-Card ("Child - " + (Get-DetectedType $ch)) @(
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
function Update-LinkButton(){ if($script:PreviewIsLinkable){ $btnAdd.Enabled=$true } else { $btnAdd.Enabled=$false } }
function Ensure-PrevPropParentLabel(){
  if($valPrevPropParent -and $valPrevParent){
    if([string]::IsNullOrWhiteSpace($valPrevPropParent.Text)){
      $valPrevPropParent.Text = $valPrevParent.Text
    }
  }
}
function Preview-Peripheral([string]$query,$parentRec){
  $valPrevType.Text=''; $valPrevName.Text=''; $valPrevAT.Text=''; $valPrevSN.Text=''; $valPrevParent.Text=''; $valPrevRITM.Text=''; $valPrevRetire.Text=''; $valPrevProposed.Text=''
  $valPrevPropParent.Text=''
  $script:PreviewCandidate = $null; $script:PreviewIsLinkable = $false
  $n = Normalize-Scan $query; if(-not $n){ Update-LinkButton; return }
  $key=$n.Value.ToUpper()
  $cand=$null
  if($script:IndexByAsset.ContainsKey($key)){$cand=$script:IndexByAsset[$key]}
  elseif($script:IndexBySerial.ContainsKey($key)){$cand=$script:IndexBySerial[$key]}
  elseif($script:IndexByName.ContainsKey($key)){$cand=$script:IndexByName[$key]}
  if(-not $cand){ $valPrevType.Text='(not found)'; Update-LinkButton; return }
  $script:PreviewCandidate = $cand
  $valPrevType.Text=$cand.Type; $valPrevName.Text=$cand.name; $valPrevAT.Text=$cand.asset_tag; $valPrevSN.Text=$cand.serial_number
  if([string]::IsNullOrWhiteSpace($cand.u_parent_asset)){ $valPrevParent.Text='(none)' } else { $valPrevParent.Text=$cand.u_parent_asset }
  Ensure-PrevPropParentLabel
  if($parentRec){
    $prop = Compute-ProposedName $cand $parentRec
    if($prop){ $valPrevProposed.Text = $prop }
  }
  if($cand.Type -eq 'Monitor'){ $valPrevRITM.Text=$cand.RITM; $valPrevRetire.Text = (Fmt-DateLong $cand.Retire) }
  else { $valPrevRITM.Text=''; $valPrevRetire.Text='' }
  if($cand.Kind -eq 'Peripheral' -and $parentRec){ $script:PreviewIsLinkable = $true }
  Update-LinkButton
}
function Log-AssocChange([string]$action,[string]$deviceType,[string]$childAT,[string]$oldParent,[string]$newParent,[string]$oldName,[string]$newName){
  $out = if($script:OutputFolder){$script:OutputFolder}else{$script:DataFolder}
  if(-not $out){ return }
  $file = Join-Path $out 'CMDBUpdates.csv'
  $row = [pscustomobject]@{
    Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Action    = $action
    DeviceType= $deviceType
    AssetTag  = $childAT
    OldParent = $oldParent
    NewParent = $newParent
    OldName   = $oldName
    NewName   = $newName
  }
  if(Test-Path $file){ $row | Export-Csv -Path $file -NoTypeInformation -Append -Encoding UTF8 }
  else { $row | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8 }
}
function Link-Peripheral([string]$query,$parentRec){
  if(-not $script:PreviewIsLinkable){ return }
  $n=Normalize-Scan $query; if(-not $n){ return }
  $key=$n.Value.ToUpper()
  $cand=$null
  if($script:IndexByAsset.ContainsKey($key)){$cand=$script:IndexByAsset[$key]}
  elseif($script:IndexBySerial.ContainsKey($key)){$cand=$script:IndexBySerial[$key]}
  elseif($script:IndexByName.ContainsKey($key)){$cand=$script:IndexByName[$key]}
  if(-not $cand){ [System.Windows.Forms.MessageBox]::Show("Peripheral not found.","Link") | Out-Null; return }
  if($cand.Kind -ne 'Peripheral'){ [System.Windows.Forms.MessageBox]::Show("Selected item is not a peripheral.","Link") | Out-Null; return }
  $oldParent = $cand.u_parent_asset
  $oldName = $cand.name
  if($cand.Type -eq 'Cart'){
    $cand.u_parent_asset = $parentRec.asset_tag
    $cand.name = "$($parentRec.name)-CRT"
  } elseif(($cand.Type -eq 'Mic') -or ($cand.Type -eq 'Scanner')){
    $targetParentAT = $parentRec.asset_tag
    if($parentRec.name -match '^(?i)AO'){
      $carts = Find-CartsForComputer $parentRec
      if($carts.Count -gt 0){
        $cart = $carts[0]
        if($cart.asset_tag){ $targetParentAT = $cart.asset_tag } else { $targetParentAT = $cart.name }
      }
    }
    $cand.u_parent_asset = $targetParentAT
    if($cand.Type -eq 'Mic'){ $cand.name = "$($parentRec.name)-Mic" }
    if($cand.Type -eq 'Scanner'){ $cand.name = "$($parentRec.name)-SCN" }
  } else {
    $cand.u_parent_asset = $parentRec.asset_tag
    switch($cand.Type){
      'Monitor' { $cand.name = $parentRec.name }
      default   { }
    }
  }
  Build-Indices
  Log-AssocChange 'Link' (Get-DetectedType $cand) $cand.asset_tag $oldParent $cand.u_parent_asset $oldName $cand.name
  Refresh-AssocViews $parentRec
  Update-CartCheckbox-State $parentRec
  Update-ManualRoundButton   $parentRec
  $txtAdd.Clear()
  $script:PreviewCandidate = $null
  $script:PreviewIsLinkable = $false
  Preview-Peripheral '' $parentRec
  Update-LinkButton
  Validate-ParentAndName $script:CurrentDisplay $script:CurrentParent
  Update-FixNameButton $script:CurrentDisplay $script:CurrentParent
}
function Remove-Selected-Associations($parentRec){
  if($dgv.SelectedRows.Count -eq 0){ return }
  foreach($row in $dgv.SelectedRows){
    $role = [string]$row.Cells['Role'].Value
    if($role -ne 'Child'){ continue }
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
  Preview-Peripheral $txtAdd.Text $parentRec
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
  Update-LinkButton
  $statusLabel.Text = ("Renamed to '" + $expected + "'.")
}
# ---- Populate Summary/UI ----
function Populate-UI($displayRec,$parentRec){
  try { Populate-Department-Combo $displayRec.u_department_location } catch {}
  try { if($displayRec.u_device_rounding){ $cmbMaintType.Text=$displayRec.u_device_rounding } } catch {}
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
  Preview-Peripheral $txtAdd.Text $parentRec
  Validate-ParentAndName $displayRec $parentRec
  Update-FixNameButton $displayRec $parentRec
}
# ---- Location cascading (City > Location > Building > Floor > Room) ----
function Populate-Location-Combos([string]$city,[string]$loc,[string]$b,[string]$f,[string]$r){
  $cmbCity.Items.Clear(); $cmbLocation.Items.Clear(); $cmbBuilding.Items.Clear(); $cmbFloor.Items.Clear(); $cmbRoom.Items.Clear()
  # City
  $cities = $script:LocationRows | ForEach-Object { Get-LocVal $_ 'City' } | Where-Object { $_ } | Select-Object -Unique | Sort-Object
  $cmbCity.Items.AddRange(@($cities))
  if($city){ $cmbCity.Text = $city }
  # Location (filtered by City if present)
  $locRows = $script:LocationRows
  if($cmbCity.Text){
    $locRows = $locRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'City')) -eq (Normalize-Field $cmbCity.Text) }
  }
  $locs = $locRows | ForEach-Object { Get-LocVal $_ 'Location' } | Where-Object { $_ } | Select-Object -Unique | Sort-Object
  $cmbLocation.Items.AddRange(@($locs))
  if($loc){ $cmbLocation.Text=$loc }
  if($cmbLocation.Text){
    $blds = $script:LocationRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Location')) -eq (Normalize-Field $cmbLocation.Text) -and ( -not $cmbCity.Text -or (Normalize-Field (Get-LocVal $_ 'City')) -eq (Normalize-Field $cmbCity.Text) ) } | ForEach-Object { Get-LocVal $_ 'Building' } | Select-Object -Unique | Sort-Object
    $cmbBuilding.Items.AddRange(@($blds))
  }
  if($b){ $cmbBuilding.Text=$b }
  if($cmbLocation.Text -and $cmbBuilding.Text){
    $fls = $script:LocationRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Location')) -eq (Normalize-Field $cmbLocation.Text) -and (Normalize-Field (Get-LocVal $_ 'Building')) -eq (Normalize-Field $cmbBuilding.Text) } | ForEach-Object { Get-LocVal $_ 'Floor' } | Select-Object -Unique
    $fls = Sort-Floors $fls
    $cmbFloor.Items.AddRange(@($fls))
  }
  if($f){ $cmbFloor.Text=$f }
  if($cmbLocation.Text -and $cmbBuilding.Text -and $cmbFloor.Text){
    $rms = $script:LocationRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Location')) -eq (Normalize-Field $cmbLocation.Text) -and (Normalize-Field (Get-LocVal $_ 'Building')) -eq (Normalize-Field $cmbBuilding.Text) -and (Normalize-Field (Get-LocVal $_ 'Floor')) -eq (Normalize-Field $cmbFloor.Text) } | ForEach-Object { Get-LocVal $_ 'Room' } | Select-Object -Unique | Sort-Object
    $cmbRoom.Items.AddRange(@($rms))
  }
  if($r){ $cmbRoom.Text=$r }
}
function Toggle-EditLocation(){
  $script:editing = -not $script:editing
  if($script:editing){
    Populate-Location-Combos $txtCity.Text $txtLocation.Text $txtBldg.Text $txtFloor.Text $txtRoom.Text
    $cmbCity.Visible=$true; $cmbLocation.Visible=$true; $cmbBuilding.Visible=$true; $cmbFloor.Visible=$true; $cmbRoom.Visible=$true
    $txtCity.Visible=$false; $txtLocation.Visible=$false; $txtBldg.Visible=$false; $txtFloor.Visible=$false; $txtRoom.Visible=$false
    if ($cmbDept) { $cmbDept.Visible=$true } ; if ($txtDept) { $txtDept.Visible=$false }
    foreach($t in @($txtDept,$txtDepartment)){ if ($t) { $t.Visible = $false } }
    foreach($c in @($cmbDept,$cmbDepartment,$ddlDept,$ddlDepartment)){ if ($c) { $c.Visible = $true } }
    $btnEditLoc.Text="Save Location"
  } else {
    $city=$cmbCity.Text; $loc=$cmbLocation.Text; $b=$cmbBuilding.Text; $f=$cmbFloor.Text; $r=$cmbRoom.Text
    $exists = $script:LocationRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'City')) -eq (Normalize-Field $city) -and (Normalize-Field (Get-LocVal $_ 'Location')) -eq (Normalize-Field $loc) -and (Normalize-Field (Get-LocVal $_ 'Building')) -eq (Normalize-Field $b) -and (Normalize-Field (Get-LocVal $_ 'Floor')) -eq (Normalize-Field $f) -and (Normalize-Field (Get-LocVal $_ 'Room')) -eq (Normalize-Field $r) }
    if($exists.Count -eq 0 -and $city -and $loc -and $b -and $f -and $r){
      $new=[pscustomobject]@{City=$city;Location=$loc;Building=$b;Floor=$f;Room=$r}
      $script:LocationRows += $new
      Save-LocationUserAdd $city $loc $b $f $r
      Rebuild-RoomCaches
    }
    $txtCity.Text=$city; $txtLocation.Text=$loc; $txtBldg.Text=$b; $txtFloor.Text=$f; $txtRoom.Text=$r
    $tmp=[pscustomobject]@{location=$loc;u_building=$b;u_floor=$f;u_room=$r}
    $script:editing = $false
    Validate-Location $tmp
    $cmbCity.Visible=$false; $cmbLocation.Visible=$false; $cmbBuilding.Visible=$false; $cmbFloor.Visible=$false; $cmbRoom.Visible=$false
    $txtCity.Visible=$true; $txtLocation.Visible=$true; $txtBldg.Visible=$true; $txtFloor.Visible=$true; $txtRoom.Visible=$true; if($cmbDept){ $txtDept.Text = $cmbDept.Text }
    if ($cmbDept) { $cmbDept.Visible=$false } ; if ($txtDept) { $txtDept.Visible=$true; $txtDept.Text = $cmbDept.Text }
    $btnEditLoc.Text="Edit Location"
  }
}
$cmbCity.Add_TextChanged({
  $cmbLocation.Items.Clear(); $cmbBuilding.Items.Clear(); $cmbFloor.Items.Clear(); $cmbRoom.Items.Clear()
  $locRows = $script:LocationRows
  if($cmbCity.Text){ $locRows = $locRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'City')) -eq (Normalize-Field $cmbCity.Text) } }
  $locs = $locRows | ForEach-Object { Get-LocVal $_ 'Location' } | Where-Object { $_ } | Select-Object -Unique | Sort-Object
  $cmbLocation.Items.AddRange(@($locs))
})
$cmbLocation.Add_TextChanged({
  $cmbBuilding.Items.Clear(); $cmbFloor.Items.Clear(); $cmbRoom.Items.Clear()
  if($cmbLocation.Text){
    $blds = $script:LocationRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Location')) -eq (Normalize-Field $cmbLocation.Text) -and ( -not $cmbCity.Text -or (Normalize-Field (Get-LocVal $_ 'City')) -eq (Normalize-Field $cmbCity.Text) ) } | ForEach-Object { Get-LocVal $_ 'Building' } | Select-Object -Unique | Sort-Object
    $cmbBuilding.Items.AddRange(@($blds))
  }
})
$cmbBuilding.Add_TextChanged({
  $cmbFloor.Items.Clear(); $cmbRoom.Items.Clear()
  if($cmbLocation.Text -and $cmbBuilding.Text){
    $fls = $script:LocationRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Location')) -eq (Normalize-Field $cmbLocation.Text) -and (Normalize-Field (Get-LocVal $_ 'Building')) -eq (Normalize-Field $cmbBuilding.Text) } | ForEach-Object { Get-LocVal $_ 'Floor' } | Select-Object -Unique
    $fls = Sort-Floors $fls
    $cmbFloor.Items.AddRange(@($fls))
    # Force user to pick a valid Floor/Room for the new Building
    $cmbFloor.Text=''; $cmbRoom.Text=''
  }
})
$cmbFloor.Add_TextChanged({
  $cmbRoom.Items.Clear()
  if($cmbLocation.Text -and $cmbBuilding.Text -and $cmbFloor.Text){
    $rms = $script:LocationRows | Where-Object { (Normalize-Field (Get-LocVal $_ 'Location')) -eq (Normalize-Field $cmbLocation.Text) -and (Normalize-Field (Get-LocVal $_ 'Building')) -eq (Normalize-Field $cmbBuilding.Text) -and (Normalize-Field (Get-LocVal $_ 'Floor')) -eq (Normalize-Field $cmbFloor.Text) } | ForEach-Object { Get-LocVal $_ 'Room' } | Select-Object -Unique | Sort-Object
    $cmbRoom.Items.AddRange(@($rms))
  }
})
# ---- Actions ----
function Do-Lookup(){
  $raw = Find-RecordRaw $txtScan.Text
  if(-not $raw){ $statusLabel.Text=("No match for '" + $txtScan.Text + "'"); return }
  $parent = Resolve-ParentComputer $raw
  Populate-UI $raw $parent
  $statusLabel.Text=("Found " + $raw.Kind + " / " + $raw.Type)
}
function Clear-UI(){
  $script:CurrentDisplay = $null; $script:CurrentParent  = $null
  foreach($tb in @($txtType,$txtHost,$txtAT,$txtSN,$txtParent,$txtRITM,$txtRetire,$txtRound,$txtCity,$txtLocation,$txtBldg,$txtFloor,$txtRoom)){
    $tb.Text = ''; $tb.BackColor = [System.Drawing.Color]::White
  }
  try { $dgv.Rows.Clear() } catch {}
  try { $cards.Controls.Clear() } catch {}
  $txtAdd.Text = ''
  foreach($l in @($valPrevType,$valPrevName,$valPrevAT,$valPrevSN,$valPrevParent,$valPrevRITM,$valPrevRetire,$valPrevProposed)){ $l.Text='' }
  $btnAdd.Enabled = $false; $script:PreviewCandidate = $null; $script:PreviewIsLinkable = $false
  Update-ManualRoundButton $null; Update-CartCheckbox-State $null
  foreach($cb in @($chkCable,$chkLabels,$chkCart,$chkPeriph)){ $cb.Checked=$false }
  $btnFixName.Enabled = $false
  $statusLabel.Text = "Ready - scan or enter a device."
  Size-AssocForRows(1) | Out-Null
}
# ---- Events ----
$btnLookup.Add_Click({ Do-Lookup })
$txtScan.Add_KeyDown({ if($_.KeyCode -eq 'Enter'){ Do-Lookup; $_.SuppressKeyPress=$true } })
$txtScan.Add_TextChanged({ if([string]::IsNullOrWhiteSpace($txtScan.Text)){ Clear-UI } })
$btnEditLoc.Add_Click({ Toggle-EditLocation })
$txtAdd.Add_TextChanged({
  $pc = $script:CurrentParent
  Preview-Peripheral $txtAdd.Text $pc
})
$btnAdd.Add_Click({
  $pc = $script:CurrentParent
  if($pc){
      try{ if(-not $chkShowExcluded.Checked){ $mt = ('' + $pc.u_device_rounding).Trim(); if($mt -match '^(?i)Excluded$'){ continue } } } catch {}
 Link-Peripheral $txtAdd.Text $pc }
})
$btnRemove.Add_Click({
  $pc = $script:CurrentParent
  if($pc){
      try{ if(-not $chkShowExcluded.Checked){ $mt = ('' + $pc.u_device_rounding).Trim(); if($mt -match '^(?i)Excluded$'){ continue } } } catch {}
 Remove-Selected-Associations $pc }
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
  if($rec){ $par=Resolve-ParentComputer $rec; Populate-UI $rec $par }
})
$btnCheckComplete.Add_Click({
  $chkCable.Checked=$true; $chkLabels.Checked=$true; $chkPeriph.Checked=$true
  $pc = $script:CurrentParent
  if($pc -and $pc.name -match '^(?i)AO'){ $chkCart.Enabled=$true; $chkCart.Checked=$true } else { $chkCart.Checked=$false; $chkCart.Enabled=$false }
})
$btnSave.Add_Click({
  $out = $script:OutputFolder
  if(-not (Test-Path $out)){ New-Item -ItemType Directory -Path $out -Force | Out-Null }
$file = Join-Path ($(if($script:OutputFolder){$script:OutputFolder}else{$script:DataFolder})) 'RoundingEvents.csv'
  $exists = Test-Path $file
  $pc = $script:CurrentParent
  if(-not $pc){ $pc = Resolve-ParentComputer (Find-RecordRaw $txtAT.Text) }
  if(-not $pc){ $pc = $script:CurrentDisplay }
  $url = $null
  if($pc){
      try{ if(-not $chkShowExcluded.Checked){ $mt = ('' + $pc.u_device_rounding).Trim(); if($mt -match '^(?i)Excluded$'){ continue } } } catch {}
 $url = Get-RoundingUrlForParent $pc }
  $row = [pscustomobject]@{
    Timestamp        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    AssetTag         = if($pc){
      try{ if(-not $chkShowExcluded.Checked){ $mt = ('' + $pc.u_device_rounding).Trim(); if($mt -match '^(?i)Excluded$'){ continue } } } catch {}
$pc.asset_tag}else{$null}
    Name             = if($pc){
      try{ if(-not $chkShowExcluded.Checked){ $mt = ('' + $pc.u_device_rounding).Trim(); if($mt -match '^(?i)Excluded$'){ continue } } } catch {}
$pc.name}else{$null}
    Serial           = if($pc){
      try{ if(-not $chkShowExcluded.Checked){ $mt = ('' + $pc.u_device_rounding).Trim(); if($mt -match '^(?i)Excluded$'){ continue } } } catch {}
$pc.serial_number}else{$null}
    City             = $txtCity.Text
    Location         = $txtLocation.Text
    Building         = $txtBldg.Text
    Floor            = $txtFloor.Text
    Room             = $txtRoom.Text
    CheckStatus      = $cmbChkStatus.Text
    RoundingMinutes  = [int]$numTime.Value
    CableMgmtOK      = $chkCable.Checked
    LabelOK          = $chkLabels.Checked
    CartOK           = $chkCart.Checked
    PeripheralsOK    = $chkPeriph.Checked
    MaintenanceType = $cmbMaintType.Text
    Department       = $cmbDept.Text
    RoundingUrl      = $url
  }
  $cmbDept.Visible = $false  # Hidden until Edit Location is active
  if(-not $exists){ $row | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8 }
  else { $row | Export-Csv -Path $file -NoTypeInformation -Append -Encoding UTF8 }
  foreach($cb in @($chkCable,$chkLabels,$chkCart,$chkPeriph)){ $cb.Checked = $false }
  [System.Windows.Forms.MessageBox]::Show(("Saved rounding event to
" + $file),"Save Event") | Out-Null
  $cmbChkStatus.SelectedIndex = 0
  $txtScan.Clear()
  Clear-UI
  $txtScan.Focus()
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
  $lblDataPath.Visible=$false; $lblOutputPath.Visible=$false; $lblDataStatus.Visible=$false; $statusLabel.Text = ("Data: " + $script:DataFolder + " | Output: " + $script:OutputFolder); $statusLabel.ForeColor=[System.Drawing.Color]::DarkGreen
  $lblOutputPath.Text = "Output: " + $script:OutputFolder
  $statusLabel.Text   = "Data OK"; $statusLabel.ForeColor=[System.Drawing.Color]::DarkGreen
} catch {
  $lblDataPath.Visible=$false; $lblOutputPath.Visible=$false; $lblDataStatus.Visible=$false; $statusLabel.Text = "Data files missing or error"; $statusLabel.ForeColor=[System.Drawing.Color]::Crimson
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
function Load-RoundingEvents {
  $script:RoundingEvents = @()
  try {
    $base = if($script:OutputFolder){ $script:OutputFolder } else { $script:DataFolder }
    if(-not $base){ return }
    $file = Join-Path $base 'RoundingEvents.csv'
    if(Test-Path $file){
      try { $script:RoundingEvents = Import-Csv $file } catch { $script:RoundingEvents = @() }
    }
  } catch { $script:RoundingEvents = @() }
}
Load-RoundingEvents
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
$nearToolbar.Height = 40
$lblScopes = New-Object System.Windows.Forms.Label
$lblScopes.AutoSize = $true
$lblScopes.Text = "Nearby scopes: 0"
$lblScopes.Location = '8,10'
$chkViewAll = New-Object System.Windows.Forms.CheckBox
$chkViewAll.Text = "View all (include today's rounded)"
$chkViewAll.AutoSize = $true
$chkViewAll.Location = '170,8'
$chkViewAll.Checked = $false

$chkShowExcluded = New-Object System.Windows.Forms.CheckBox
$chkShowExcluded.Text = "Show Excluded"
$chkShowExcluded.AutoSize = $true
$chkShowExcluded.Location = '390,8'
$chkShowExcluded.Checked = $false
$chkShowExcluded.Add_CheckedChanged({ Rebuild-Nearby })

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
$btnClearScopes = New-Object System.Windows.Forms.Button
$btnClearScopes.Text = "Clear List"
$btnClearScopes.AutoSize = $true
$btnClearScopes.Location = '700,6'
$nearToolbar.Controls.AddRange(@($lblScopes,$chkViewAll,$btnClearScopes))
# --- Multi-Status (apply one status to selected rows) ---
$btnSetStatus = New-Object System.Windows.Forms.Button
$btnSetStatus.Text = 'Multi-Status'
$btnSetStatus.Width = 110
$btnSetStatus.Height = 24
$btnSetStatus.Location = '560,4'
$nearToolbar.Controls.Add($btnSetStatus)
if (-not $menuStatus) { $menuStatus = New-Object System.Windows.Forms.ContextMenuStrip }
$btnSetStatus.Add_Click({
  try {
    $menuStatus.Items.Clear()
    $options = Get-StatusOptionsFromGrid
    foreach ($opt in $options) {
      $item = New-Object System.Windows.Forms.ToolStripMenuItem($opt)
      $item.Add_Click({
        param($s,$e)
        $chosen = $s.Text
        # Ensure the Status column/cells include this option
        try {
          $col = $dgvNearby.Columns['Status']
          if ($col -and $col -is [System.Windows.Forms.DataGridViewComboBoxColumn]) {
            if (-not $col.Items.Contains($chosen)) { [void]$col.Items.Add($chosen) }
          }
        } catch {}
        $count = 0
        foreach ($row in $dgvNearby.SelectedRows) {
          $cell = $row.Cells['Status']
          if ($cell -and $cell -is [System.Windows.Forms.DataGridViewComboBoxCell]) {
            if (-not $cell.Items.Contains($chosen)) { [void]$cell.Items.Add($chosen) }
          }
          if ($row -and $row.Cells['Status']) {
            $row.Cells['Status'].Value = $chosen
            $count++
          }
        }
        [System.Windows.Forms.MessageBox]::Show("Updated status for $count row(s).","Multi-Status",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
      })
      [void]$menuStatus.Items.Add($item)
    }
    [void]$menuStatus.Items.Add('-')
    $custom = New-Object System.Windows.Forms.ToolStripMenuItem('Custom...')
    $custom.Add_Click({
      Add-Type -AssemblyName Microsoft.VisualBasic
      $val = [Microsoft.VisualBasic.Interaction]::InputBox('Enter custom status for selected rows:','Multi-Status','')
      if ([string]::IsNullOrWhiteSpace($val)) { return }
      try {
        $col = $dgvNearby.Columns['Status']
        if ($col -and $col -is [System.Windows.Forms.DataGridViewComboBoxColumn]) {
          if (-not $col.Items.Contains($val)) { [void]$col.Items.Add($val) }
        }
      } catch {}
      $count = 0
      foreach ($row in $dgvNearby.SelectedRows) {
        $cell = $row.Cells['Status']
        if ($cell -and $cell -is [System.Windows.Forms.DataGridViewComboBoxCell]) {
          if (-not $cell.Items.Contains($val)) { [void]$cell.Items.Add($val) }
        }
        if ($row -and $row.Cells['Status']) {
          $row.Cells['Status'].Value = $val
          $count++
        }
      }
      [System.Windows.Forms.MessageBox]::Show("Updated status for $count row(s).","Multi-Status",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    })
    [void]$menuStatus.Items.Add($custom)
    $pt = New-Object System.Drawing.Point(0,$btnSetStatus.Height)
    $menuStatus.Show($btnSetStatus, $pt)
  } catch {
    [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)","Multi-Status",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
  }
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
$dgvNearby.Columns.Add((New-NearCol 'Department' 'Department' 160)) | Out-Null; $dgvNearby.Columns.Add((New-NearCol 'LastRounded' 'Last Rounded' 130))| Out-Null
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
    $lsd = if ($dir -eq 'Asc') { [System.ComponentModel.ListSortDirection]::Ascending } else { [System.ComponentModel.ListSortDirection]::Descending }
    $sender.Sort($col, $lsd)
    $col.HeaderCell.SortGlyphDirection = if ($dir -eq 'Asc') { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending }
  })
} catch {}
$colHiddenAT = New-NearCol 'AT_KEY' '__ATKEY' 10 $true; $colHiddenAT.Visible=$false; $dgvNearby.Columns.Add($colHiddenAT) | Out-Null
$colHiddenToday = New-NearCol 'TODAY' '__TODAY' 10 $true; $colHiddenToday.Visible=$false; $dgvNearby.Columns.Add($colHiddenToday) | Out-Null
$colHiddenLRRaw = New-NearCol 'LRRAW' '__LRRAW' 10 $true; $colHiddenLRRaw.Visible=$false; $dgvNearby.Columns.Add($colHiddenLRRaw) | Out-Null
# Bottom bar
$nearBottom = New-Object System.Windows.Forms.Panel
$nearBottom.Dock = 'Bottom'
$nearBottom.Height = 48
$lblNearNote = New-Object System.Windows.Forms.Label
$lblNearNote.AutoSize = $true
$lblNearNote.Text = "Bulk Save adds events with selected Status, 3 min each. Today's devices are shown only when ""View all"" is on and are not saved again."
$lblNearNote.Location = '8,14'
$btnNearSave = New-Object System.Windows.Forms.Button
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
$tabPageNear = New-Object System.Windows.Forms.TabPage
$tabPageNear.Text = 'Nearby'
# Move existing main UI into a panel and then into TabPageMain
$pageMain = New-Object System.Windows.Forms.Panel
$pageMain.Dock = 'Fill'
# Re-parent header + main table into panel
$form.Controls.Remove($panelTop)
$form.Controls.Remove($tlpMain)
$pageMain.Controls.Add($tlpMain); $tlpMain.Dock='Fill'
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
  $lblScopes.Text = "Nearby scopes (Location): " + $script:ActiveNearbyScopes.Count
}
function Should-Include-PC-InScopes($pc){
  if (-not $pc) { return $false }
  $k = ScopeKey $null $pc.location $null $null
  if (-not $k) { return $false }
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
    if (-not $chkViewAll.Checked -and $isToday) { continue }
    $rowIdx = $dgvNearby.Rows.Add()
    $r = $dgvNearby.Rows[$rowIdx]
    $r.Cells['Host'].Value      = $pc.name
    $r.Cells['Asset'].Value     = $pc.asset_tag
    $r.Cells['Location'].Value  = $pc.location
    $r.Cells['Building'].Value  = $pc.u_building
    $r.Cells['Floor'].Value     = $pc.u_floor
    $r.Cells['Room'].Value      = $pc.u_room
      $r.Cells['Department'].Value = $pc.u_department_location
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
  # Apply sort
  Apply-NearbySort
  Update-ScopeLabel
try { $dgvNearby.ResumeLayout() } catch {}
try { $form.Cursor = [System.Windows.Forms.Cursors]::Default } catch {}}
function Apply-NearbySort {
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
    $tabTop.SelectedTab = $tabPageMain
  }
})
# React to toolbar changes
$chkViewAll.Add_CheckedChanged({ Rebuild-Nearby })
$cmbSort.Add_SelectedIndexChanged({})
$btnClearScopes.Add_Click({
  $script:ActiveNearbyScopes.Clear()
  Update-ScopeLabel
  $dgvNearby.Rows.Clear()
})
# Bulk Save from Nearby
$btnNearSave.Add_Click({
  $out = $script:OutputFolder
  if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out -Force | Out-Null }
$file = Join-Path ($(if($script:OutputFolder){$script:OutputFolder}else{$script:DataFolder})) 'RoundingEvents.csv'
  $exists = Test-Path $file
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
      CableMgmtOK      = $false
      LabelOK          = $false
      CartOK           = $false
      PeripheralsOK    = $false
      MaintenanceType  = if ($pc) { $pc.u_device_rounding } else { $null }
      Department       = $row.Cells['Department'].Value
      RoundingUrl      = $url
    }
    if (-not $exists) { $ev | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8 }
    else { $ev | Export-Csv -Path $file -NoTypeInformation -Append -Encoding UTF8 }
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
[void]$form.ShowDialog()

  # Department
  $txtDept.Text = $rec.u_department_location
  $okD = $false
  $nDept = Normalize-Field $rec.u_department_location
  if($script:DepartmentListNorm -and $nDept){ $okD = $script:DepartmentListNorm.Contains($nDept) }
