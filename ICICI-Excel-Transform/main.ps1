param(
    [Parameter(Mandatory)]
    [string]$Folder,

    [Parameter(Mandatory=$false)]
    [string]$OutputFolder
)

# If OutputFolder not provided → use same folder
if (-not $OutputFolder) {
    $OutputFolder = $Folder
}

# Add module path (temporary)
$env:PSModulePath += ";C:\Projects\Automation\Modules"

Import-Module PCXLab.Excel -Force

# Ensure output folder exists
if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$files = Get-ChildItem $Folder -File

foreach ($file in $files) {

    # 🚫 Skip already processed files
    if ($file.Name -match "_ConvertedFromXls" -or $file.Name -match "_Transformed") {
        Write-Host "Skipping: $($file.Name)" -ForegroundColor DarkGray
        continue
    }

    Write-Host "Processing: $($file.Name)" -ForegroundColor Cyan

    try {
        # Step 1: Convert (if needed)
        $workingFile = Convert-XlsToXlsx -File $file

        # Step 2: Transform
        $result = Convert-ICICIFormat -File $workingFile

        # Step 3: Build output file name
        $outFile = Join-Path $OutputFolder ($file.BaseName + "_Transformed.xlsx")

        # Step 4: Export
        $result | Export-Excel -Path $outFile -AutoSize -BoldTopRow

        Write-Host "Saved: $outFile" -ForegroundColor Green
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}