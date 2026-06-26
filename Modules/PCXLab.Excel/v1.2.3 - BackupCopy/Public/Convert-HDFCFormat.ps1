function Convert-HDFCFormat {

    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    # Ensure XLS -> XLSX
    $workingFile = Convert-XlsToXlsx -File $File

    # Get MOP
    $mop = Get-MOPFromFileName -FileName $workingFile.Name

    # Read Excel
    $raw = Import-Excel $workingFile.FullName -NoHeader

    # Find header row
    $headerIndex = Get-HDFCHeader -RawData $raw

    if ($null -eq $headerIndex) {
        throw "HDFC Header not found in file [$($File.Name)]"
    }

    # Extract header row
    $headerRow = @($raw[$headerIndex].PSObject.Properties.Value)

    Write-Verbose "Header Row:"
    Write-Verbose ($headerRow -join " | ")

    # Build column map
    $colMap = @{
        DateTime = $null
        Details  = $null
        Amount   = $null
        DrCr     = $null
    }

    for ($i = 0; $i -lt $headerRow.Count; $i++) {

        $header = "$($headerRow[$i])".Trim()

        if ([string]::IsNullOrWhiteSpace($header)) {
            continue
        }

        switch -Regex ($header) {

            'Date' {
                if ($null -eq $colMap.DateTime) {
                    $colMap.DateTime = $i
                }
            }

            'Description' {
                if ($null -eq $colMap.Details) {
                    $colMap.Details = $i
                }
            }

            'Debit\s*/\s*Credit' {
                if ($null -eq $colMap.Amount) {
                    $colMap.Amount = $i
                }

                if ($null -eq $colMap.DrCr) {
                    $colMap.DrCr = $i
                }
            }

            'AMT|Amount' {
                if ($null -eq $colMap.Amount) {
                    $colMap.Amount = $i
                }
            }

            '^Debit$|^Credit$|Dr/Cr' {
                if ($null -eq $colMap.DrCr) {
                    $colMap.DrCr = $i
                }
            }
        }
    }

    Write-Verbose "DateTime Column : $($colMap.DateTime)"
    Write-Verbose "Details Column  : $($colMap.Details)"
    Write-Verbose "Amount Column   : $($colMap.Amount)"
    Write-Verbose "DrCr Column     : $($colMap.DrCr)"

    # Validation
    foreach ($requiredColumn in @('DateTime','Details','Amount')) {

        if ($null -eq $colMap[$requiredColumn]) {
            throw "Required column [$requiredColumn] not found in file [$($File.Name)]"
        }
    }

    # Data starts after header
    $data = $raw[($headerIndex + 1)..($raw.Count - 1)]

    foreach ($row in $data) {

        $values = @($row.PSObject.Properties.Value)

        if ($values.Count -le $colMap.Amount) {
            continue
        }

        $dateTime = $values[$colMap.DateTime]
        $details  = $values[$colMap.Details]
        $amount   = $values[$colMap.Amount]

        $drcr = ""

        if ($null -ne $colMap.DrCr) {
            $drcr = $values[$colMap.DrCr]
        }

        # Skip empty rows
        if ([string]::IsNullOrWhiteSpace($dateTime)) { continue }
        if ([string]::IsNullOrWhiteSpace($details))  { continue }

        # Extract date
        if ($dateTime -match '(\d{2}/\d{2}/\d{4})') {

            $date = [datetime]::ParseExact(
                $matches[1],
                'dd/MM/yyyy',
                $null
            ).ToString('dd-MM-yyyy')
        }
        else {
            continue
        }

        # Clean amount
        $amount = "$amount"
        $amount = $amount -replace ',', ''
        $amount = $amount.Trim()

        if ([string]::IsNullOrWhiteSpace($amount)) {
            continue
        }

        try {
            $amountValue = [decimal]$amount
        }
        catch {
            continue
        }

        $amtDr = 0
        $amtCr = 0

        if ($drcr -match 'Cr|Credit') {
            $amtCr = $amountValue
        }
        else {
            $amtDr = $amountValue
        }

        # Extract Ref#
        $ref = ""

        if ($details -match 'Ref#\s*([A-Za-z0-9]+)') {
            $ref = "Ref# $($matches[1])"
        }

        # Clean narration
        $cleanDetails = (
            $details -replace '\s*\(Ref#.*?\)', ''
        ).Trim()

        [PSCustomObject]@{
            Date           = $date
            Narration      = $cleanDetails
            Item           = ""
            Category       = ""
            Place          = ""
            Freq           = ""
            For            = ""
            MOP            = $mop
            'Amt (Dr)'     = $amtDr
            'Chq./Ref.No.' = $ref
            'Value Dt'     = if ($amtDr -gt 0) { 'Dr.' } elseif ($amtCr -gt 0) { 'Cr.' } else { '' }
            'Amt (Cr)'     = $amtCr
        }
    }
}