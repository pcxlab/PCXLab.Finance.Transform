function Convert-ICICICreditCardCSV {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    # Mode of Payment
    $mop = Get-MOPFromFileName -FileName $File.Name

    # Read Excel (converted from CSV)
    $raw = Import-Excel $File.FullName -NoHeader

    # Locate header
    $headerIndex = Get-ICICICreditCardHeader -RawData $raw

    if ($null -eq $headerIndex) {
        throw "Header not found in $($File.Name)"
    }

    # Build column map
    $headerRow = $raw[$headerIndex].PSObject.Properties.Value

    $colMap = @{}

    for ($i = 0; $i -lt $headerRow.Count; $i++) {

        switch ($headerRow[$i]) {

            "Date" {
                $colMap.Date = $i
            }

            "Transaction Details" {
                $colMap.Details = $i
            }

            "Amount(in Rs)" {
                $colMap.Amount = $i
            }

            "BillingAmountSign" {
                $colMap.Sign = $i
            }

            "Sr.No." {
                $colMap.Ref = $i
            }
        }
    }

    if ($colMap.Count -lt 5) {
        throw "Column mapping incomplete."
    }

    #
    # Data starts AFTER the masked card row
    #

    $data = $raw[($headerIndex + 2)..($raw.Count - 1)]

    foreach ($row in $data) {

        $values = $row.PSObject.Properties.Value

        #
        # Skip masked card number rows
        #

        if ($values[$colMap.Date] -match "XXXX") {
            continue
        }

        $date = $values[$colMap.Date]

        if ([string]::IsNullOrWhiteSpace($date)) {
            continue
        }

        $details = $values[$colMap.Details]
        $amount = $values[$colMap.Amount]
        $sign = $values[$colMap.Sign]
        $ref = $values[$colMap.Ref]

        #
        # Credit / Debit
        #

        if ($sign -eq "CR") {

            $amtCr = [decimal]$amount
            $amtDr = 0

        }
        else {

            $amtDr = [decimal]$amount
            $amtCr = 0
        }

        [PSCustomObject]@{

            Date = $date
            Narration = $details
            Item = ""
            Category = ""
            Place = ""
            Freq = ""
            For = ""
            MOP = $mop

            "Amt (Dr)" = $amtDr
            "Chq./Ref.No." = $ref
            "Value Dt" = if ($amtDr -gt 0) { "Dr." } else { "Cr." }
            "Amt (Cr)" = $amtCr
        }
    }
}