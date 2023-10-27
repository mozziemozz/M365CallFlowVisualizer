<#
    .SYNOPSIS
    Reads all voice apps and resource account into variables.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.1
    Changelog:          .\Changelog.md

#>
function Get-AllVoiceAppsAndResourceAccounts {
    param (
    )

    $queryResultSize = 100

    if ([String]::IsNullOrEmpty($Global:allAutoAttendants) -or $CacheResults -eq $false) {

        Write-Host "Retrieving all Auto Attendants... this can take a while..." -ForegroundColor Magenta

        $Global:allAutoAttendants = Get-CsAutoAttendant -First $queryResultSize

        if ($Global:allAutoAttendants.Count -ge $queryResultSize) {

            Write-Host "This tenant has at least $queryResultSize or more Auto Attendants. Querrying additional AAs..." -ForegroundColor Cyan

            $skipCounter = $queryResultSize

            do {
        
                $querriedAAs = Get-CsAutoAttendant -Skip $skipCounter
        
                $Global:allAutoAttendants += $querriedAAs

                $skipCounter += $querriedAAs.Count

            } until (
                $querriedAAs.Count -eq 0
            )

        }

        Write-Host "Finished getting all Auto Attendants. Number of Auto Attendants found: $($Global:allAutoAttendants.Count)"

    }

    else {

        Write-Warning  "Auto Attendant config is read from memory. If you don't see recent changes reflected in the output, use the -CacheResults `$false parameter."

    }

    if ([String]::IsNullOrEmpty($Global:allCallQueues) -or $CacheResults -eq $false) {

        Write-Host "Retrieving all Call Queues... this can take a while..." -ForegroundColor Magenta

        $Global:allCallQueues = Get-CsCallQueue -WarningAction SilentlyContinue -First $queryResultSize

        if ($Global:allCallQueues.Count -ge $queryResultSize) {

            Write-Host "This tenant has at least $queryResultSize or more Call Queues. Querrying additional CQs..." -ForegroundColor Cyan

            $skipCounter = $queryResultSize

            do {
        
                $querriedCQs = Get-CsCallQueue -WarningAction SilentlyContinue -Skip $skipCounter
        
                $Global:allCallQueues += $querriedCQs

                $skipCounter += $querriedCQs.Count

            } until (
                $querriedCQs.Count -eq 0
            )

        }

        Write-Host "Finished getting all Call Queues. Number of Call Queues found: $($Global:allCallQueues.Count)"

    }

    else {

        Write-Warning "Call Queue config is read from memory. If you don't see recent changes reflected in the output, use the -CacheResults `$false parameter."

    }

    if ([String]::IsNullOrEmpty($Global:allResourceAccounts) -or $CacheResults -eq $false) {

        Write-Host "Retrieving all Resource Accounts... this can take a while..." -ForegroundColor Magenta

        $Global:allResourceAccounts = Get-CsOnlineApplicationInstance -ResultSize 9999

        # -ResultSize not working in MicrosoftTeams PowerShell 5.0.0

        # $Global:allResourceAccounts = Get-CsOnlineApplicationInstance -ResultSize $queryResultSize

        # if ($allResourceAccounts.Count -ge $queryResultSize) {

        #     Write-Host "This tenant has at least $queryResultSize or more Resource Accounts. Querrying additional RAs..." -ForegroundColor Cyan

        #     $skipCounter = $queryResultSize

        #     do {
        
        #         $querriedRAs = Get-CsOnlineApplicationInstance -Skip $skipCounter
        
        #         $allResourceAccounts += $querriedRAs

        #         $skipCounter += $querriedRAs.Count

        #     } until (
        #         $querriedRAs.Count -eq 0
        #     )

        # }

        Write-Host "Finished getting all Resource Accounts. Number of Resource Accounts found: $($allResourceAccounts.Count)"


    }

    else {

        Write-Warning "Resource Accounts are read from memory. If you don't see recent changes reflected in the output, use the -CacheResults `$false parameter."

    }

}