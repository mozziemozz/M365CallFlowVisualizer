<#
    .SYNOPSIS
    Truncates a string to a specified length, preserving file extension if present.

    .DESCRIPTION
    If the string has a file extension (e.g. .wav, .mp3), it will be preserved
    after truncation. Used for shortening audio filenames and TTS greetings
    in Mermaid diagram nodes.

    Author:             Patrik Lleshaj (https://github.com/realgarit)
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function Get-TruncatedString {
    param (
        [Parameter(Mandatory = $true)][String]$InputString,
        [Parameter(Mandatory = $true)][Int]$MaxLength,
        [Parameter(Mandatory = $false)][Switch]$PreserveExtension
    )

    if ($InputString.Length -le $MaxLength) {
        return $InputString
    }

    if ($PreserveExtension) {
        $lastDot = $InputString.LastIndexOf('.')
        if ($lastDot -ge 0 -and ($InputString.Length - $lastDot) -le 5) {
            $extension = $InputString.Substring($lastDot)
            $truncated = $InputString.Substring(0, $MaxLength) + "... " + $extension
            return $truncated
        }
    }

    return $InputString.Substring(0, $MaxLength).TrimEnd() + "..."
}
