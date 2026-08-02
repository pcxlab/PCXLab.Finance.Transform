function Get-HDFCHeader {

    <#
    .SYNOPSIS
        Finds the header row in an HDFC Credit Card statement.

    .DESCRIPTION
        Scans the imported worksheet and locates the transaction
        header by verifying the presence of all mandatory column names.

        This approach is resilient to HDFC adding additional
        headings (for example "Domestic/International Transactions")
        or inserting extra rows before the transaction table.

    .PARAMETER RawData
        Excel data imported using Import-Excel.

    .OUTPUTS
        System.Int32

    .EXAMPLE
        $headerRow = Get-HDFCHeader -RawData $raw
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$RawData
    )

    Write-Verbose "Searching for HDFC transaction header..."

    for ($i = 0; $i -lt $RawData.Count; $i++) {

        $row = @($RawData[$i].PSObject.Properties.Value)

        # Ignore completely blank rows
        if (($row | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -eq 0) {
            continue
        }

        # Normalize values
        $cells = $row | ForEach-Object {
            if ($_) {
                $_.ToString().Trim()
            }
        }

        if (
            $cells -contains 'Transaction type' -and
            $cells -contains 'Date & Time' -and
            $cells -contains 'Description' -and
            $cells -contains 'AMT' -and
            $cells -contains 'Debit / Credit'
        ) {
            Write-Verbose "Transaction header found at row $($i + 1)."

            return $i
        }
    }

    throw "Unable to locate the HDFC transaction header. The statement format may have changed."
}