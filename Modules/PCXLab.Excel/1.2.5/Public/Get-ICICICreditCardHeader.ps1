function Get-ICICICreditCardHeader {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $RawData
    )

    for ($i = 0; $i -lt $RawData.Count; $i++) {

        $rowValues = $RawData[$i].PSObject.Properties.Value

        $hasDate = $rowValues | Where-Object { $_ -eq "Date" }

        $hasDetails = $rowValues | Where-Object {
            $_ -eq "Transaction Details"
        }

        $hasAmount = $rowValues | Where-Object {
            $_ -eq "Amount(in Rs)"
        }

        if ($hasDate -and $hasDetails -and $hasAmount) {
            return $i
        }
    }

    return $null
}