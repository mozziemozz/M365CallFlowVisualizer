<#
    .SYNOPSIS
    Reads Business Hours from an auto Attendant in a human readable format.
    
    .DESCRIPTION
    Author:             Luca Sain (https://github.com/ChocoMilkWithoutSugar)
    Conributors:        Martin Heusser
    Version:            1.0.2
    Changelog:          .\Changelog.md

#>

function Optimize-DisplayName {
    param (
        $String
    )
    return $String = ($String `
        -Replace "\(","" `
        -Replace "\)","" `
        -Replace "\[","" `
        -Replace "\]","" `
        -Replace "\|","" `
        -Replace "\~","" `
        -Replace "  "," " `
        -Replace "@"," at "`
        -Replace "call", "Call"
        ).Trim()
}
