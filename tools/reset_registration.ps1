$ErrorActionPreference = "Stop"

$toolsDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$closeExcelScript = Join-Path $toolsDirectory "close_excel.ps1"
& powershell -ExecutionPolicy Bypass -File $closeExcelScript
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$addinNames = @(
    "ExcelCalibrationAddin.Vsto",
    "ExcelCalibrationWorkbenchAddin.Vsto",
    "ExcelTemplateComAddin.Vsto",
    "ExcelStandaloneComAddin.Vsto"
)

$identityPatterns = @(
    "ExcelCalibrationAddin",
    "ExcelCalibrationWorkbenchAddin",
    "ExcelTemplateComAddin",
    "ExcelStandaloneComAddin",
    "excel_addin",
    "excel_com_addin"
)

foreach ($addinName in $addinNames) {
    $addinKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\$addinName"
    if (Test-Path $addinKey) {
        Remove-Item -LiteralPath $addinKey -Recurse -Force
        Write-Host "Removed Excel add-in registry key: $addinKey"
    }
}

$uninstallRoot = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
if (Test-Path $uninstallRoot) {
    Get-ChildItem -LiteralPath $uninstallRoot | ForEach-Object {
        $properties = Get-ItemProperty -LiteralPath $_.PSPath
        $text = (($properties.PSObject.Properties | ForEach-Object { [string]$_.Value }) -join " ")
        if ($identityPatterns | Where-Object { $text -like "*$_*" }) {
            Remove-Item -LiteralPath $_.PSPath -Recurse -Force
            Write-Host "Removed old uninstall entry: $($_.PSChildName)"
        }
    }
}

$solutionMetadataRoot = "HKCU:\Software\Microsoft\VSTO\SolutionMetadata"
if (Test-Path $solutionMetadataRoot) {
    Get-ChildItem -LiteralPath $solutionMetadataRoot | ForEach-Object {
        $properties = Get-ItemProperty -LiteralPath $_.PSPath
        $text = (($properties.PSObject.Properties | ForEach-Object { [string]$_.Value }) -join " ")
        if ($addinNames -contains $properties.addInName -or
            $addinNames -contains $properties.friendlyName -or
            ($identityPatterns | Where-Object { $text -like "*$_*" })) {
            Remove-Item -LiteralPath $_.PSPath -Recurse -Force
            Write-Host "Removed VSTO solution metadata: $($_.PSChildName)"
        }
    }
}

$inclusionRoot = "HKCU:\Software\Microsoft\VSTO\Security\Inclusion"
if (Test-Path $inclusionRoot) {
    Get-ChildItem -LiteralPath $inclusionRoot | ForEach-Object {
        $properties = Get-ItemProperty -LiteralPath $_.PSPath
        $text = (($properties.PSObject.Properties | ForEach-Object { [string]$_.Value }) -join " ")
        if ($identityPatterns | Where-Object { $text -like "*$_*" }) {
            Remove-Item -LiteralPath $_.PSPath -Recurse -Force
            Write-Host "Removed VSTO trust entry: $($_.PSChildName)"
        }
    }
}

$vstoCache = Join-Path $env:LOCALAPPDATA "Apps\2.0"
$rundll = Join-Path $env:WINDIR "System32\rundll32.exe"
if (Test-Path $rundll) {
    & $rundll dfshim CleanOnlineAppCache
    Write-Host "ClickOnce cache cleaned via $rundll"
} else {
    Write-Warning "rundll32.exe not found at $rundll. Skipped ClickOnce cache cleanup."
}

Write-Host "Registration reset complete."
Write-Host "If Office still reports a cached ClickOnce conflict, run:"
Write-Host "  $rundll dfshim CleanOnlineAppCache"
Write-Host "VSTO ClickOnce cache location: $vstoCache"
