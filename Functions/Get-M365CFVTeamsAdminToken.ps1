<#
    .SYNOPSIS
    # Fork of: https://gist.github.com/ChrFrohn/99317724b35d37c11b0f149c1f123dad#file-connecttoteams-serviceprincipal-ps1
    
    .DESCRIPTION
    Author:             Christian Frohn
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function Get-M365CFVTeamsAdminToken {
    param (
        [Parameter(Mandatory = $true)][String]$TenantId,
        [Parameter(Mandatory = $true)][String]$AppId,
        [Parameter(Mandatory = $true)][String]$AppSecret
    )

    $graphtokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $AppId
        Client_Secret = $AppSecret
    }

    $graphToken = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token" -Method POST -Body $graphtokenBody | Select-Object -ExpandProperty Access_Token

    $teamstokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "48ac35b8-9aa8-4d74-927d-1f4a14a0b239/.default"
        Client_Id     = $AppId
        Client_Secret = $AppSecret
    }

    $teamsToken = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token" -Method POST -Body $teamstokenBody | Select-Object -ExpandProperty Access_Token
  
}