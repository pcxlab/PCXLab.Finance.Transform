param(
    [Parameter(Mandatory)]
    [string]$Folder,

    [Parameter(Mandatory=$false)]
    [string]$OutputFolder
)

# Validate input folder
if (-not (Test-Path $Folder)) {
    throw "Input folder does not exist: $Folder"
}

# If OutputFolder not provided → use same folder
if (-not $OutputFolder) {
    $OutputFolder = $Folder
}

# Ensure output folder exists
if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# Add module path (temporary)
$env:PSModulePath += ";C:\Projects\Automation\Modules"

# Import module (no hardcoding needed if installed properly)
Import-Module PCXLab.Excel -Force

# Get all files
$files = Get-ChildItem $Folder -File

foreach ($file in $files) {

    # Skip already processed files
    if ($file.Name -match "_ConvertedFromXls" -or $file.Name -match "_Transformed") {
        Write-Host "Skipping: $($file.Name)" -ForegroundColor DarkGray
        continue
    }

    Write-Host "Processing: $($file.Name)" -ForegroundColor Cyan

    try {
        # Step 1: Convert (if needed)
        $workingFile = Convert-XlsToXlsx -File $file

        # Step 2: Transform (ONLY transform here, no conversion inside)
        $result = Convert-ICICIFormat -File $workingFile

        # Step 3: Output file
        $outFileName = Get-OutputFileName -File $file -Converted:$($file.Extension -eq ".xls") -Transformed
        $outFile = Join-Path $OutputFolder $outFileName

        # Step 4: Export
        $result | Export-Excel -Path $outFile -AutoSize -BoldTopRow

        Write-Host "Saved: $outFile" -ForegroundColor Green
    }
    catch {
        Write-Host "Error processing $($file.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
}