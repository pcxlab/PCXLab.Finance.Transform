param(
    [Parameter(Mandatory)]
    [string]$Folder,

    [string]$OutputFolder
)

# Validate input folder
if (-not (Test-Path $Folder)) {
    throw "Input folder does not exist: $Folder"
}

# Default output folder = same as input
if (-not $OutputFolder) {
    $OutputFolder = $Folder
}

# Ensure output folder exists
if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# Import modules
Import-Module PCXLab.Core -Force
Import-Module PCXLab.Excel -Force

# Start logging
Start-LogSession -LogFolder (Join-Path $Folder "logs")

# Get files
$files = Get-ChildItem $Folder -File

foreach ($file in $files) {

    # Skip already processed
    if ($file.Name -match "_ConvertedFromXls" -or $file.Name -match "_Transformed") {
        Write-Log "Skipping: $($file.Name)"
        continue
    }

    Write-Log "Processing: $($file.Name)"

    try {
        # Step 1: Convert if needed
        $workingFile = Convert-XlsToXlsx -File $file

        # Step 2: Transform
        $result = Convert-ICICIFormat -File $workingFile

        # Step 3: Output naming (clean centralized logic)
        $outFileName = Get-OutputFileName `
            -File $file `
            -Converted:$($file.Extension -eq ".xls") `
            -Transformed

        $outFile = Join-Path $OutputFolder $outFileName

        # Step 4: Export
        $result | Export-Excel -Path $outFile -AutoSize -BoldTopRow

        Write-Log "Saved: $outFile" "SUCCESS"
    }
    catch {
        Write-Log "Error processing $($file.Name): $($_.Exception.Message)" "ERROR"
    }
}