@echo off
setlocal
:: This CMD file wraps PowerShell to allow a one-click install and displays the path.
SET "POWERSHELL_SCRIPT=^
$installPath = \"$env:APPDATA\ExcelAI\"; ^
Write-Host \"--- Starting ExcelAI Installation ---\" -ForegroundColor Cyan; ^
Write-Host \"Destination Path: $installPath\" -ForegroundColor Yellow; ^
if (!(Test-Path $installPath)) { ^
    New-Item -ItemType Directory -Path $installPath -Force ^| Out-Null; ^
    Write-Host \"Created new directory at $installPath\"; ^
} ^
Write-Host \"Copying build files...\"; ^
Copy-Item -Path \".\dist\*\" -Destination $installPath -Recurse -Force; ^
Copy-Item -Path \".\manifest.xml\" -Destination $installPath -Force; ^
Write-Host \"Registering Trusted Catalog in Windows Registry...\"; ^
$regPath = \"HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\"; ^
$guid = \"{$( [guid]::NewGuid().ToString() )}\"; ^
$keyPath = \"$regPath\$guid\"; ^
New-Item -Path $keyPath -Force ^| Out-Null; ^
New-ItemProperty -Path $keyPath -Name \"URL\" -Value $installPath -PropertyType String ^| Out-Null; ^
New-ItemProperty -Path $keyPath -Name \"Flags\" -Value 1 -PropertyType DWord ^| Out-Null; ^
Write-Host \"`nSUCCESS: ExcelAI installed!\" -ForegroundColor Green; ^
Write-Host \"Next Steps:\" -ForegroundColor White; ^
Write-Host \"1. Restart Excel.\" -ForegroundColor White; ^
Write-Host \"2. Go to Insert -> My Add-ins -> Shared Folder.\" -ForegroundColor White; ^
Write-Host \"3. Select 'ExcelAI' and click Add.\" -ForegroundColor White"

powershell -NoProfile -ExecutionPolicy Bypass -Command "& { %POWERSHELL_SCRIPT% }"
pause
