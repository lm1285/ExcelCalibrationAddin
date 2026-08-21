[CmdletBinding()]
param(
    [ValidateRange(0, 60)]
    [int]$GracePeriodSeconds = 5
)

$ErrorActionPreference = "Stop"

$excelProcesses = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
if ($excelProcesses.Count -eq 0) {
    Write-Host "Excel is not running."
    exit 0
}

Write-Host "Closing Excel processes before add-in maintenance..."
$excelProcesses | Select-Object Id, ProcessName, MainWindowTitle | Format-Table -AutoSize

foreach ($process in $excelProcesses) {
    if ($process.MainWindowHandle -ne 0) {
        $null = $process.CloseMainWindow()
    }
}

$deadline = (Get-Date).AddSeconds($GracePeriodSeconds)
do {
    Start-Sleep -Milliseconds 250
    $remaining = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
} while ($remaining.Count -gt 0 -and (Get-Date) -lt $deadline)

if ($remaining.Count -gt 0) {
    Write-Host "Force-closing remaining Excel processes to prevent stale add-in files and registrations."
    $remaining | Stop-Process -Force
    $remaining | Wait-Process -Timeout 10 -ErrorAction SilentlyContinue
}

if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) {
    throw "Excel could not be closed. The latest add-in cannot be installed safely."
}

Write-Host "All Excel processes are closed."
