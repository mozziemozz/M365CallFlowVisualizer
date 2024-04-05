<#
    .SYNOPSIS
    Connects to Teams and Graph if not already connected. Supports interactive sign in via Microsoft Graph Command Line Tools (14d82eec-204b-4c2f-b7e8-296a70dab67e)
    or your own dedicated Entra ID App Registration / Service Principal.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.1.2
    Changelog:          .\Changelog.md

#>

function Connect-M365CFV {
    param (
    )

    if ($ShowSharedVoicemailGroupSubscribers -eq $true) {
   
        $exoConnection = Get-ConnectionInformation

        if (!$exoConnection) {

            Connect-ExchangeOnline

            $exoConnection = Get-ConnectionInformation

        }

        if ($exoConnection.Count -gt 1) {

            $exoConnection = ($exoConnection | Where-Object { $_.State -eq "Connected" -and $_.TokenStatus -eq "Active" } | Sort-Object Id -Descending)[0]

        }

        $exoTenantId = $exoConnection.TenantID

    }

    try {
        $msTeamsTenant = Get-CsTenant -ErrorAction Stop > $null
        $msTeamsTenant = Get-CsTenant
    }

    catch {

        Connect-MicrosoftTeams -ErrorAction SilentlyContinue
        $msTeamsTenant = Get-CsTenant

    }

    finally {
        
        if ($msTeamsTenant -and $? -eq $true) {

            Write-Host "Connected Teams Tenant: $($msTeamsTenant.DisplayName)" -ForegroundColor Green
            
        }

        if ($msTeamsTenant.TenantId -is [System.ValueType]) {

            $msTeamsTenantId = $msTeamsTenant.TenantId.Guid

        }

        else {

            $msTeamsTenantId = $msTeamsTenant.TenantId

        }
    }

    try {

        $testMsGraphConnection = Get-MgUser -Top 1 -ErrorAction Stop > $null
        $msGraphContext = (Get-MgContext).TenantId

        if ($msGraphContext -ne $msTeamsTenantId) {

            do {

                if ($msGraphContext) {
                    Write-Warning -Message "Connected Graph TenantId does not match connected Teams TenantId... Signing out of Graph... "
                    Disconnect-MgGraph
                }

                Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -TenantId $msTeamsTenantId

                $msGraphContext = (Get-MgContext).TenantId

            } until ($msGraphContext -eq $msTeamsTenantId)

        }
        
    }

    catch {

        if (Test-Path -Path "$env:USERPROFILE\.graph") {

            Remove-Item "$env:USERPROFILE\.graph" -Recurse -Force
            Write-Host "Microsoft Graph cache has been cleared." -ForegroundColor Yellow

        }

        if ($ConnectWithServicePrincipal) {

            Connect-MgGraph -AccessToken $graphTokenSecureString > $null

        }

        else {

            Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -TenantId $msTeamsTenantId

        }

        $msGraphContext = (Get-MgContext).TenantId
    }

    finally {

        if ($msGraphContext -eq $msTeamsTenantId -and $msTeamsTenant) {

            if ($ConnectWithServicePrincipal) {

                $EntraIdAppRegistrationName = (Get-MgContext).AppName
                Write-Host "Connected Graph Tenant matches connected Teams Tenant: $($msTeamsTenant.DisplayName)" -ForegroundColor Green
                Write-Host "Connected to Graph with App: $EntraIdAppRegistrationName" -ForegroundColor Green

            }

            else {

                $msGraphTenantName = (Get-MgContext).Account.Split("@")[-1]
                Write-Host "Connected Graph Tenant matches connected Teams Tenant: $msGraphTenantName" -ForegroundColor Green

            }

        }

        else {

            Write-Host "Not connected to Microsoft Teams. Please try again. Exiting..." -ForegroundColor Red
            Start-Sleep 3
            exit

        }
    }

    if ($ShowSharedVoicemailGroupSubscribers -eq $true) {

        if ($exoTenantId -eq $msTeamsTenantId -and $exoTenantId -eq $msGraphContext) {

            Write-Host "Connected Exchange Online Tenant matches connected Teams and Graph Tenant: $($msTeamsTenant.DisplayName)" -ForegroundColor Green

        }

        else {

            Write-Host "Connected Exchange Online Tenant does not match connected Teams and Graph Tenant. Please try again. Exiting..." -ForegroundColor Red

        }

    }
    
}