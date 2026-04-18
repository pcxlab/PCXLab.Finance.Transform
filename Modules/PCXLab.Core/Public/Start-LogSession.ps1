function Start-LogSession {
    param(
        [Parameter(Mandatory)]
        [string]$LogFolder
    )

    if (!(Test-Path $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

    $global:LogFile = Join-Path $LogFolder "Run_$timestamp.log"

    New-Item -Path $global:LogFile -ItemType File -Force | Out-Null

    Write-Host "Logging started: $global:LogFile" -ForegroundColor Gray
}