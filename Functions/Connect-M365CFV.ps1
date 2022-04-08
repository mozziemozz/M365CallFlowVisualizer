function Connect-M365CFV {
    param (
    )

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
        $msGraphContext = (Get-MgContext).TenantId

        if ($msGraphContext -ne $msTeamsTenantId) {

            do {

                if ($msGraphContext) {
                    Write-Warning -Message "Connected Graph TenantId does not match connected Teams TenantId... Signing out of Graph... "
                    Disconnect-MgGraph
                }

                Connect-MgGraph -Scopes "User.Read.All","Group.Read.All" -TenantId $msTeamsTenantId

                $msGraphContext = (Get-MgContext).TenantId

            } until ($msGraphContext -eq $msTeamsTenantId)

        }
        
    }
    catch {
        Connect-MgGraph -Scopes "User.Read.All","Group.Read.All" -TenantId $msTeamsTenantId
        $msGraphContext = (Get-MgContext).TenantId
    }
    finally {
        if ($msGraphContext -eq $msTeamsTenantId -and $msTeamsTenant) {

            $msGraphTenantName = (Get-MgContext).Account.Split("@")[-1]
            Write-Host "Connected Graph Tenant matches connected Teams Tenant: $msGraphTenantName" -ForegroundColor Green

        }

        else {

            Write-Host "Not connected to Microsoft Teams. Please try again. Exiting..." -ForegroundColor Red
            Start-Sleep 3
            exit

        }
    }
    
}