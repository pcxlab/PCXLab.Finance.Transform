function Convert-ICICICreditCardStatement {

    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    Convert-ICICIFormat -File $File
}