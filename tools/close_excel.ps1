[CmdletBinding()]
param(
    [ValidateRange(0, 60)]
    [int]$GracePeriodSeconds = 5,

    [ValidateRange(5, 120)]
    [int]$ForceExitTimeoutSeconds = 30
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

    # Excel can remain alive while VSTO/COM components unload. Retry against the
    # current process list because the original Process objects can become stale.
    $forceDeadline = (Get-Date).AddSeconds($ForceExitTimeoutSeconds)
    do {
        $remaining = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
        if ($remaining.Count -eq 0) {
            break
        }

        foreach ($process in $remaining) {
            try {
                Stop-Process -Id $process.Id -Force -ErrorAction Stop
            } catch {
                Write-Warning "Could not stop Excel process $($process.Id): $($_.Exception.Message)"
            }
        }

        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $forceDeadline)
}

$remaining = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
if ($remaining.Count -gt 0) {
    $details = ($remaining | ForEach-Object {
        $path = try { $_.Path } catch { "<unavailable>" }
        "PID=$($_.Id), Session=$($_.SessionId), Path=$path"
    }) -join "; "

    throw "Excel could not be closed after $ForceExitTimeoutSeconds seconds. Remaining process: $details. Close Excel in Task Manager, or run this updater with the same/elevated user account."
}

Write-Host "All Excel processes are closed."
