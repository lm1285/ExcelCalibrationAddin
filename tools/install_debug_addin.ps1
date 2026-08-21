$ErrorActionPreference = "Stop"

$workspace = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$vstoPath = Join-Path $workspace "src\ExcelCalibrationAddin.Vsto\bin\Debug\ExcelStandaloneComAddin.Vsto.vsto"
$buildScript = Join-Path $workspace "tools\build_latest_addin.ps1"
$installerCandidates = @(
    "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe",
    "C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
)

Write-Host "== Build a clean, current VSTO add-in =="
& powershell -ExecutionPolicy Bypass -File $buildScript -Configuration Debug
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

if (!(Test-Path $vstoPath)) {
    throw "VSTO manifest not found: $vstoPath"
}

$installer = $installerCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (!$installer) {
    throw "VSTOInstaller.exe not found. Install the Visual Studio Tools for Office Runtime."
}

Write-Host "== Install current VSTO manifest =="
& $installer /Install $vstoPath
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$addinKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelStandaloneComAddin.Vsto"
$manifestUri = ([Uri](Resolve-Path -LiteralPath $vstoPath).Path).AbsoluteUri + "|vstolocal"
New-Item -Path $addinKey -Force | Out-Null
New-ItemProperty -LiteralPath $addinKey -Name "Description" -Value "ExcelStandaloneComAddin.Vsto" -PropertyType String -Force | Out-Null
New-ItemProperty -LiteralPath $addinKey -Name "FriendlyName" -Value "ExcelStandaloneComAddin.Vsto" -PropertyType String -Force | Out-Null
New-ItemProperty -LiteralPath $addinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -LiteralPath $addinKey -Name "Manifest" -Value $manifestUri -PropertyType String -Force | Out-Null

$registeredManifest = [string](Get-ItemProperty -LiteralPath $addinKey).Manifest
if ($registeredManifest -ne $manifestUri) {
    throw "The installed Excel registration does not reference the current VSTO manifest: $registeredManifest"
}

Write-Host "Current add-in installed and registered: $registeredManifest"
