[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [string]$PublishDirectory = "outputs\vsto-publish"
)

$ErrorActionPreference = "Stop"

$workspace = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

$msbuild = & (Join-Path $workspace "tools\find_vs_msbuild.ps1")
if (!$msbuild -or !(Test-Path $msbuild)) {
    throw "Visual Studio MSBuild not found. Install Visual Studio 2022 with Office/SharePoint development tools."
}

$bootstrapperRoot = & (Join-Path $workspace "tools\prepare_vsto_bootstrapper.ps1")
$publishPath = if ([IO.Path]::IsPathRooted($PublishDirectory)) {
    [IO.Path]::GetFullPath($PublishDirectory)
} else {
    [IO.Path]::GetFullPath((Join-Path $workspace $PublishDirectory))
}
New-Item -ItemType Directory -Force -Path $publishPath | Out-Null
$publishStartedUtc = [DateTime]::UtcNow
$buildLog = Join-Path $publishPath "publish.log"

& $msbuild (Join-Path $workspace "src\ExcelCalibrationAddin.Vsto\ExcelCalibrationAddin.Vsto.csproj") `
    /t:Publish `
    /restore `
    /p:Configuration=$Configuration `
    /p:Platform=AnyCPU `
    "/p:PublishDir=$publishPath\" `
    "/p:PublishUrl=$publishPath\" `
    "/p:GenerateBootstrapperSdkPath=$bootstrapperRoot" `
    /p:IsWebBootstrapper=false `
    /warnaserror `
    /fl `
    "/flp:logfile=$buildLog;verbosity=normal" `
    /v:minimal
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$setup = Join-Path $publishPath "setup.exe"
$manifest = Join-Path $publishPath "ExcelStandaloneComAddin.Vsto.vsto"
if (!(Test-Path $setup)) {
    throw "Bootstrapper setup.exe was not generated: $setup"
}
if (!(Test-Path $manifest)) {
    throw "Published VSTO manifest was not generated: $manifest"
}
foreach ($file in @($setup, $manifest)) {
    if ((Get-Item $file).LastWriteTimeUtc -lt $publishStartedUtc) {
        throw "Publish output is stale: $file"
    }
}

Write-Host "Published VSTO installer: $setup"
