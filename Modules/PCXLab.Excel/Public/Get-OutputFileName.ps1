function Get-OutputFileName {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File,

        [switch]$Converted,
        [switch]$Transformed
    )

    # Step 1: Start with base name
    $baseName = $File.BaseName

    # Step 2: Remove existing suffixes (avoid duplication)
    $baseName = $baseName -replace "_ConvertedFromXls", ""
    $baseName = $baseName -replace "_Transformed", ""

    # Step 3: Build suffix
    $suffix = ""

    if ($Converted) {
        $suffix += "_ConvertedFromXls"
    }

    if ($Transformed) {
        $suffix += "_Transformed"
    }

    # Step 4: Always output as .xlsx
    $finalName = "$baseName$suffix.xlsx"

    return $finalName
}