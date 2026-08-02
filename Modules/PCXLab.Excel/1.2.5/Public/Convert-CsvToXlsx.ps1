function Convert-CsvToXlsx {

    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    # Already XLSX
    if ($File.Extension -ieq ".xlsx") {
        return $File
    }

    # Only CSV supported
    if ($File.Extension -ine ".csv") {
        throw "File '$($File.Name)' is not a CSV file."
    }

    # Build output path
    $newFileName = Get-OutputFileName -File $File -Converted
    $newFile = Join-Path $File.DirectoryName $newFileName

    # Reuse existing conversion
    if (Test-Path $newFile) {
        return Get-Item $newFile
    }

    Write-Host "Converting CSV -> XLSX: $($File.Name)" -ForegroundColor Yellow

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {

        $workbook = $excel.Workbooks.Open($File.FullName)

        # xlOpenXMLWorkbook
        $workbook.SaveAs($newFile, 51)

        $workbook.Close($false)
    }
    catch {
        throw "Failed to convert '$($File.Name)'"
    }
    finally {

        if ($workbook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }

        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    Get-Item $newFile
}