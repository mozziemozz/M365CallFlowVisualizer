# M365CallFlowVisualizer Module
# Loads all function files from the Functions directory

$FunctionPath = Join-Path -Path $PSScriptRoot -ChildPath 'Functions'

Get-ChildItem -Path $FunctionPath -Filter '*.ps1' -Recurse | ForEach-Object {
    try {
        . $_.FullName
    }
    catch {
        Write-Warning "Failed to load function: $($_.Name). Error: $_"
    }
}
