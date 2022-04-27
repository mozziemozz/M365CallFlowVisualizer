<#
    .SYNOPSIS
    Reads Business Hours from an auto Attendant in a human readable format.
    
    .DESCRIPTION
    Author:             Luca Sain (https://github.com/ChocoMilkWithoutSugar)
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function FixDisplayName {
    param (
        $String
    )
    return $String = $String -replace "[()]",""
}
