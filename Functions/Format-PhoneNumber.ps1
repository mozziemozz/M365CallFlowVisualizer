<#
    .SYNOPSIS
    Extracts and optionally obfuscates a phone number from a tel: URI.

    .DESCRIPTION
    Author:             Patrik Lleshaj (https://github.com/realgarit)
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function Format-PhoneNumber {
    param (
        [Parameter(Mandatory = $true)][String]$PhoneNumberId,
        [Parameter(Mandatory = $false)][Bool]$Obfuscate = $false
    )

    $number = $PhoneNumberId -replace '^tel:', ''

    if ($Obfuscate -and $number.Length -ge 4) {
        $number = $number.Substring(0, $number.Length - 4) + '****'
    }

    return $number
}
