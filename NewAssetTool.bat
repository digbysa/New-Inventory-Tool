@echo off
setlocal
cd /d "C:\Users\da1701_sa\Desktop\New-Inventory-Tool"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File ".\NewAssetTool.ps1"
echo.
echo (Done) Press any key to close...
pause >nul
endlocal
