$ErrorActionPreference = "Stop"

$workspace = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $workspace

$config = Join-Path $workspace "src\ExcelCalibrationAddin.Vsto\appsettings.json"
$settings = Get-Content -LiteralPath $config -Raw | ConvertFrom-Json
$healthUrl = "$($settings.Backend.BaseUrl.TrimEnd('/'))$($settings.Backend.TemplateApiPrefix.TrimEnd('/'))/health"
try {
    $healthResponse = Invoke-WebRequest -Uri $healthUrl -UseBasicParsing -TimeoutSec 2
    if ($healthResponse.StatusCode -eq 200) {
        Write-Host "Local backend is already running: $healthUrl"
        exit 0
    }
}
catch {
    # No healthy instance is listening, so continue with the startup sequence.
}

$msbuild = & (Join-Path $workspace "tools\find_vs_msbuild.ps1")
if (!$msbuild -or !(Test-Path $msbuild)) {
    throw "Visual Studio MSBuild not found. Install Visual Studio 2022 with Office/SharePoint development tools."
}

$serverProject = Join-Path $workspace "src\ExcelCalibrationAddin.LocalServer\ExcelCalibrationAddin.LocalServer.csproj"
& $msbuild $serverProject /t:Restore /p:Configuration=Debug
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

& $msbuild $serverProject /t:Build /p:Configuration=Debug
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$server = Join-Path $workspace "src\ExcelCalibrationAddin.LocalServer\bin\Debug\net48\ExcelCalibrationAddin.LocalServer.exe"
& $server "--config=$config"
