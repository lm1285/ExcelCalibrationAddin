[CmdletBinding()]
param(
    [string]$CacheRoot = (Join-Path $env:LOCALAPPDATA "ExcelCalibrationAddin\bootstrapper")
)

$ErrorActionPreference = "Stop"

$CacheRoot = [IO.Path]::GetFullPath($CacheRoot)
$systemRoot = Join-Path ${env:ProgramFiles(x86)} "Microsoft SDKs\ClickOnce Bootstrapper"
$packageSource = Join-Path $systemRoot "Packages\VSTOR40"
$frameworkSource = Join-Path $systemRoot "Packages\DotNetFX48"
$engineSource = Join-Path $systemRoot "Engine"
$packageRoot = Join-Path $CacheRoot "Packages\VSTOR40"
$frameworkRoot = Join-Path $CacheRoot "Packages\DotNetFX48"

if (!(Test-Path (Join-Path $packageSource "product.xml"))) {
    throw "ClickOnce Bootstrapper VSTOR40 product metadata is missing: $packageSource"
}
if (!(Test-Path (Join-Path $engineSource "setup.bin"))) {
    throw "ClickOnce Bootstrapper engine is missing: $engineSource"
}
if (!(Test-Path (Join-Path $frameworkSource "Product.xml"))) {
    throw ".NET Framework 4.8 bootstrapper metadata is missing: $frameworkSource"
}

New-Item -ItemType Directory -Force -Path $packageRoot | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "zh-Hans") | Out-Null
New-Item -ItemType Directory -Force -Path $frameworkRoot | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $CacheRoot "Engine") | Out-Null

Copy-Item (Join-Path $packageSource "product.xml") (Join-Path $packageRoot "product.xml") -Force
Copy-Item (Join-Path $packageSource "zh-Hans\package.xml") (Join-Path $packageRoot "zh-Hans\package.xml") -Force
Copy-Item (Join-Path $packageSource "zh-Hans\eula.rtf") (Join-Path $packageRoot "zh-Hans\eula.rtf") -Force
Get-ChildItem $frameworkSource -Force | Copy-Item -Destination $frameworkRoot -Recurse -Force

$runtimePayload = Join-Path $packageRoot "vstor_redist.exe"
if (!(Test-Path $runtimePayload) -or (Get-Item $runtimePayload).Length -lt 1000000) {
    Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/?LinkId=158918" -OutFile $runtimePayload -UseBasicParsing
}

# The SDK metadata predates the currently linked, Microsoft-signed VSTO runtime.
# Keep the bootstrapper's file integrity check enabled by recording the key from
# the downloaded Authenticode certificate instead of disabling PublicKey checks.
$signature = Get-AuthenticodeSignature -FilePath $runtimePayload
if ($signature.Status -ne "Valid" -or !$signature.SignerCertificate -or
    $signature.SignerCertificate.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $false) -ne "Microsoft Corporation") {
    throw "The VSTO runtime must have a valid Microsoft Corporation Authenticode signature: $($signature.Status)"
}
$runtimePublicKey = ($signature.SignerCertificate.GetPublicKeyString() -replace "\s", "").ToUpperInvariant()
$productPath = Join-Path $packageRoot "product.xml"
$product = New-Object System.Xml.XmlDocument
$product.PreserveWhitespace = $true
$product.Load($productPath)
$namespaces = New-Object System.Xml.XmlNamespaceManager($product.NameTable)
$namespaces.AddNamespace("b", "http://schemas.microsoft.com/developer/2004/01/bootstrapper")
$packageFile = $product.SelectSingleNode("/b:Product/b:PackageFiles/b:PackageFile[@Name='vstor_redist.exe']", $namespaces)
if (!$packageFile) {
    throw "VSTOR40 metadata does not declare vstor_redist.exe."
}
$packageFile.SetAttribute("PublicKey", $runtimePublicKey)
$product.Save($productPath)
Write-Host "Verified Microsoft VSTO runtime $((Get-Item $runtimePayload).VersionInfo.FileVersion); SHA256=$((Get-FileHash $runtimePayload -Algorithm SHA256).Hash)"

Copy-Item (Join-Path $engineSource "setup.bin") (Join-Path $CacheRoot "Engine\setup.bin") -Force
Copy-Item (Join-Path $engineSource "Launcher.exe") (Join-Path $CacheRoot "Engine\Launcher.exe") -Force
Get-ChildItem $engineSource -Directory | ForEach-Object {
    $destination = Join-Path $CacheRoot "Engine\$($_.Name)"
    New-Item -ItemType Directory -Force -Path $destination | Out-Null
    Copy-Item (Join-Path $_.FullName "setup.xml") (Join-Path $destination "setup.xml") -Force
}

Write-Output $CacheRoot
