<#
    .SYNOPSIS
    Sanitizes display names for safe use in Mermaid diagram syntax.

    .DESCRIPTION
    Removes or replaces characters that would break Mermaid diagram rendering,
    including parentheses, brackets, pipes, tildes, and special quotes.

    Author:             Luca Sain (https://github.com/ChocoMilkWithoutSugar)
    Contributors:       Martin Heusser, Patrik Lleshaj (https://github.com/realgarit)
    Version:            1.1.0
    Changelog:          .\Changelog.md

#>

function Optimize-DisplayName {
    param (
        [Parameter(Mandatory = $true)][String]$String
    )

    $result = $String -replace '[(\)\[\]|~]', '' `
                       -replace '  ', ' ' `
                       -replace '@', ' at ' `
                       -replace 'call', 'Call' `
                       -replace [char]0x2019, "'" `
                       -replace "`n", "'"

    return $result.Trim()
}
