<#
    .SYNOPSIS
    Finds Auto Attendant, Call Queues and Resource Accounts which are not in use anywhere in your tenant. (Can't be called externally.)

    Author:             Martin Heusser
    Version:            1.0.0
    Revision:
        28.01.2023      1.0.0: Creation

#>

#Requires -Modules @{ ModuleName = "MicrosoftTeams"; ModuleVersion = "4.9.1" }, "Microsoft.Graph.Users", "Microsoft.Graph.Groups"

. .\Functions\Connect-M365CFV.ps1

. Connect-M365CFV

Write-Host "Retrieving all Auto Attendants (max. 1000)... this can take a while..." -ForegroundColor Magenta
$global:allAutoAttendants = Get-CsAutoAttendant -First 1000

Write-Host "Retrieving all Call Queues (max. 1000)... this can take a while..." -ForegroundColor Magenta
$global:allCallQueues = Get-CsCallQueue -WarningAction SilentlyContinue -First 1000

Write-Host "Retrieving all Resource Accounts (max. 1000)... this can take a while..." -ForegroundColor Magenta
$global:allResourceAccounts = Get-CsOnlineApplicationInstance -ResultSize 1000

$report = @()

$autoAttendantsWithoutResourceAccounts = $allAutoAttendants | Where-Object {!$_.ApplicationInstances}

# Add Auto Attendants without resource accounts to report
foreach ($autoAttendant in $autoAttendantsWithoutResourceAccounts) {

    $autoAttendantDetails = New-Object -TypeName psobject

    $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Name" -Value $autoAttendant.Name
    $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Identity" -Value $autoAttendant.Identity
    $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Category" -Value "Voice App"
    $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Type" -Value "Auto Attendant"
    $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Association" -Value $false
    $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Phone Number Assigned" -Value $false
    $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Nested in Voice Apps" -Value $false


    $report += $autoAttendantDetails

}

$callQueuesWithoutResourceAccounts = $allCallQueues | Where-Object {!$_.ApplicationInstances}

# Add Call Queues without resource accounts to report
foreach ($callQueue in $callQueuesWithoutResourceAccounts) {

    $callQueueDetails = New-Object -TypeName psobject

    $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Name" -Value $callQueue.Name
    $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Identity" -Value $callQueue.Identity
    $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Category" -Value "Voice App"
    $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Type" -Value "Call Queue"
    $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Association" -Value $false
    $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Phone Number Assigned" -Value $false
    $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Nested in Voice Apps" -Value $false


    $report += $callQueueDetails

}

$resourceAccountIdsWithPhoneNumber = ($allResourceAccounts | Where-Object {$_.PhoneNumber}).ObjectId

$autoAttendantsWithResourceAccounts = $allAutoAttendants | Where-Object {$_.ApplicationInstances}

$callQueuesWithResourceAccounts = $allCallQueues | Where-Object {$_.ApplicationInstances}

$autoAttendantsWithPhoneNumber = @()
$callQueuesWithPhoneNumber = @()

foreach ($resourceAccount in $resourceAccountIdsWithPhoneNumber) {

    foreach ($autoAttendant in $autoAttendantsWithResourceAccounts) {

        if ($autoAttendant.ApplicationInstances -contains $resourceAccount) {

            $autoAttendantsWithPhoneNumber += $autoAttendant

        }

    }

    foreach ($callQueue in $callQueuesWithResourceAccounts) {

        if ($callQueue.ApplicationInstances -contains $resourceAccount) {

            $callQueuesWithPhoneNumber += $callQueue

        }

    }

}

$autoAttendantsWithPhoneNumber = $autoAttendantsWithPhoneNumber | Sort-Object Name -Unique
$callQueuesWithPhoneNumber = $callQueueWithPhoneNumber | Sort-Object Name -Unique

$nonTopLevelAutoAttendants = $autoAttendantsWithResourceAccounts | Where-Object {$autoAttendantsWithPhoneNumber -notcontains $_}
$nonTopLevelCallQueues = $callQueuesWithResourceAccounts | Where-Object {$callQueuesWithPhoneNumber -notcontains $_}

$allNestedVoiceAppIds = @()

foreach ($autoAttendantWithPhoneNumber in $autoAttendantsWithPhoneNumber) {

    . .\M365CallFlowVisualizerV2.ps1 -Identity $autoAttendantWithPhoneNumber.Identity -SaveToFile $false -SetClipBoard $false -ExportHtml $false -ShowNestedCallFlows $true -ShowUserCallingSettings $false -NoCache $false

    $allNestedVoiceAppIds += $nestedVoiceApps

}


foreach ($callQueueWithPhoneNumber in $callQueuesWithPhoneNumber) {

    . .\M365CallFlowVisualizerV2.ps1 -Identity $callQueueWithPhoneNumber.Identity -SaveToFile $false -SetClipBoard $false -ExportHtml $false -ShowNestedCallFlows $true -ShowUserCallingSettings $false -NoCache $false

    $allNestedVoiceAppIds += $nestedVoiceApps

}

$allNestedVoiceAppIds = $allNestedVoiceAppIds | Sort-Object -Unique

foreach ($autoAttendant in $nonTopLevelAutoAttendants) {

    if ($allNestedVoiceAppIds -notcontains $autoAttendant.Identity) {

        $autoAttendantDetails = New-Object -TypeName psobject

        $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Name" -Value $autoAttendant.Name
        $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Identity" -Value $autoAttendant.Identity
        $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Category" -Value "Voice App"
        $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Type" -Value "Auto Attendant"
        $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Association" -Value $true
        $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Phone Number Assigned" -Value $false
        $autoAttendantDetails | Add-Member -MemberType NoteProperty -Name "Nested in Voice Apps" -Value $false
    
        $report += $autoAttendantDetails

    }

}

foreach ($callQueue in $nonTopLevelcallQueues) {

    if ($allNestedVoiceAppIds -notcontains $callQueue.Identity) {

        $callQueueDetails = New-Object -TypeName psobject

        $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Name" -Value $callQueue.Name
        $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Identity" -Value $callQueue.Identity
        $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Category" -Value "Voice App"
        $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Type" -Value "Call Queue"
        $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Association" -Value $true
        $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Phone Number Assigned" -Value $false
        $callQueueDetails | Add-Member -MemberType NoteProperty -Name "Nested in Voice Apps" -Value $false
    
        $report += $callQueueDetails

    }

}

$allAssociatedVoiceApps = $allAutoAttendants.ApplicationInstances + $allCallQueues.ApplicationInstances

$unassignedResourceAccounts = $allResourceAccounts.ObjectId | Where-Object {$allAssociatedVoiceApps -notcontains $_}

foreach ($resourceAccount in $unassignedResourceAccounts) {

    $resourceAccountDetails = New-Object -TypeName psobject

    $applicationInstance = $allResourceAccounts | Where-Object {$_.ObjectId -eq $resourceAccount}

    $resourceAccountDetails | Add-Member -MemberType NoteProperty -Name "Name" -Value "$($applicationInstance.DisplayName) (UPN: $($applicationInstance.UserPrincipalName))"
    $resourceAccountDetails | Add-Member -MemberType NoteProperty -Name "Identity" -Value $applicationInstance.ObjectId
    $resourceAccountDetails | Add-Member -MemberType NoteProperty -Name "Category" -Value "Resource Account"

    switch ($applicationInstance.ApplicationId) {

        # Auto Attendants
        "ce933385-9390-45d1-9512-c8d228074e07" {

            $resourceAccountDetails | Add-Member -MemberType NoteProperty -Name "Type" -Value "Auto Attendant"

        }
        # Call Queues
        "11cd3e2e-fccb-42ad-ad00-878b93575e07" {

            $resourceAccountDetails | Add-Member -MemberType NoteProperty -Name "Type" -Value "Call Queue"

        }
        Default {

            $resourceAccountDetails | Add-Member -MemberType NoteProperty -Name "Type" -Value "Other"

        }

    }

    $resourceAccountDetails | Add-Member -MemberType NoteProperty -Name "Association" -Value $false

    if ($applicationInstance.PhoneNumber) {

        $resourceAccountDetails | Add-Member -MemberType NoteProperty -Name "Phone Number Assigned" -Value $true

    }

    else {

        $resourceAccountDetails | Add-Member -MemberType NoteProperty -Name "Phone Number Assigned" -Value $false

    }

    $resourceAccountDetails | Add-Member -MemberType NoteProperty -Name "Nested in Voice Apps" -Value $false


    $report += $resourceAccountDetails

}

$report | Export-Csv -Path ".\UnusedVoiceAppsAndResourceAccountsReport.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force