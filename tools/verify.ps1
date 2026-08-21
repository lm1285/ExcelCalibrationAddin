$ErrorActionPreference = "Stop"

$workspace = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $workspace

$msbuild = & (Join-Path $workspace "tools\find_vs_msbuild.ps1")
if (!$msbuild -or !(Test-Path $msbuild)) {
    throw "Visual Studio MSBuild not found. Install Visual Studio 2022 with Office/SharePoint development tools."
}

$solution = Join-Path $workspace "ExcelCalibrationAddin.sln"

Write-Host "== VSTO restore =="
& $msbuild $solution /t:Restore /p:Configuration=Debug /p:Platform="Any CPU"
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Write-Host "== VSTO build =="
& $msbuild $solution /t:Build /p:Configuration=Debug /p:Platform="Any CPU"
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
