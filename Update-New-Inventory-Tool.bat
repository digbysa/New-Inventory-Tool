@echo off
setlocal EnableExtensions DisableDelayedExpansion

set "PS1=%TEMP%\Update-New-Inventory-Tool.ps1"
del /q "%PS1%" 2>nul

>> "%PS1%" echo $ErrorActionPreference = "Stop"
>> "%PS1%" echo
>> "%PS1%" echo $RepoOwner        = "digbysa"
>> "%PS1%" echo $RepoName         = "New-Inventory-Tool"
>> "%PS1%" echo $Branch           = "VGHVersion1.0"
>> "%PS1%" echo $TargetFolderName = "New-Inventory-Tool"
>> "%PS1%" echo
>> "%PS1%" echo try ^{
>> "%PS1%" echo     try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
>> "%PS1%" echo
>> "%PS1%" echo     $Desktop    = [Environment]::GetFolderPath("Desktop")
>> "%PS1%" echo     $TargetPath = Join-Path $Desktop $TargetFolderName
>> "%PS1%" echo
>> "%PS1%" echo     $zipUrl = "https://github.com/$RepoOwner/$RepoName/archive/refs/heads/$Branch.zip"
>> "%PS1%" echo
>> "%PS1%" echo     $tempRoot   = Join-Path $env:TEMP ("gh_update_" + [guid]::NewGuid().ToString("N"))
>> "%PS1%" echo     $zipPath    = Join-Path $tempRoot "repo.zip"
>> "%PS1%" echo     $expandPath = Join-Path $tempRoot "expanded"
>> "%PS1%" echo
>> "%PS1%" echo     New-Item -ItemType Directory -Path $tempRoot   ^| Out-Null
>> "%PS1%" echo     New-Item -ItemType Directory -Path $expandPath ^| Out-Null
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host "Downloading latest from GitHub..."
>> "%PS1%" echo     Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host "Extracting..."
>> "%PS1%" echo     Expand-Archive -Path $zipPath -DestinationPath $expandPath -Force
>> "%PS1%" echo
>> "%PS1%" echo     $rootFolder = Get-ChildItem -Path $expandPath -Directory ^| Select-Object -First 1
>> "%PS1%" echo     if (-not $rootFolder) { throw "Extraction failed (no folder found after unzip)." }
>> "%PS1%" echo
>> "%PS1%" echo     $sourcePath = $rootFolder.FullName
>> "%PS1%" echo
>> "%PS1%" echo     if (Test-Path $TargetPath) ^{
>> "%PS1%" echo         $stamp      = Get-Date -Format "yyyyMMdd_HHmmss"
>> "%PS1%" echo         $backupPath = "${TargetPath}_backup_$stamp"
>> "%PS1%" echo         Write-Host "Backing up existing folder to: $backupPath"
>> "%PS1%" echo         Rename-Item -Path $TargetPath -NewName (Split-Path $backupPath -Leaf)
>> "%PS1%" echo     ^}
>> "%PS1%" echo
>> "%PS1%" echo     New-Item -ItemType Directory -Path $TargetPath -Force ^| Out-Null
>> "%PS1%" echo     Write-Host "Copying files to: $TargetPath"
>> "%PS1%" echo     Copy-Item -Path (Join-Path $sourcePath "*") -Destination $TargetPath -Recurse -Force
>> "%PS1%" echo
>> "%PS1%" echo     Remove-Item -Path $tempRoot -Recurse -Force
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host ""
>> "%PS1%" echo     Write-Host "Done! Updated New-Inventory-Tool on your Desktop." -ForegroundColor Green
>> "%PS1%" echo     Write-Host "If an old version existed, it was backed up with _backup_YYYYMMDD_HHMMSS." -ForegroundColor DarkGreen
>> "%PS1%" echo ^}
>> "%PS1%" echo catch ^{
>> "%PS1%" echo     Write-Host ""
>> "%PS1%" echo     Write-Host "UPDATE FAILED: $($_.Exception.Message)" -ForegroundColor Red
>> "%PS1%" echo     Write-Host "Tip: Close the tool first if files are in use, then try again." -ForegroundColor Yellow
>> "%PS1%" echo     exit 1
>> "%PS1%" echo ^}

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
set "RC=%ERRORLEVEL%"

del /q "%PS1%" 2>nul
exit /b %RC%
