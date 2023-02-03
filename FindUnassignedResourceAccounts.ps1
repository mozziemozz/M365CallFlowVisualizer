<#
    .SYNOPSIS
    Finds and exports unassigned Teams resource accounts.

    Author:             Martin Heusser
    Version:            1.0.0
    Revision:
        08.12.2022      1.0.0: Creation

#>

#Requires -Modules @{ ModuleName = "MicrosoftTeams"; ModuleVersion = "4.9.3" }, "Microsoft.Graph.Users", "Microsoft.Graph.Groups"

. .\Functions\Connect-M365CFV.ps1

. Connect-M365CFV

Write-Host "Retrieving all Auto Attendants (max. 1000)... this can take a while..." -ForegroundColor Magenta
$allAutoAttendants = Get-CsAutoAttendant -First 1000
Write-Host "Retrieving all Call Queues (max. 1000)... this can take a while..." -ForegroundColor Magenta
$allCallQueues = Get-CsCallQueue -WarningAction SilentlyContinue -First 1000

Write-Host "Retrieving all Resource Accounts (max. 1000)... this can take a while..." -ForegroundColor Magenta
$allResourceAccounts = Get-CsOnlineApplicationInstance -ResultSize 1000

$allAssociatedVoiceApps = $allAutoAttendants.ApplicationInstances + $allCallQueues.ApplicationInstances

$unassignedResourceAccounts = $allResourceAccounts.ObjectId | Where-Object {$allAssociatedVoiceApps -notcontains $_}

$unassignedResourceAccountsExport = @()

foreach ($resourceAccount in $unassignedResourceAccounts) {

    $resourceAccountProperties = New-Object -TypeName psobject

    $applicationInstance = $allResourceAccounts | Where-Object {$_.ObjectId -eq $resourceAccount}

    $resourceAccountProperties | Add-Member -MemberType NoteProperty -Name "Display Name" -Value $applicationInstance.DisplayName
    $resourceAccountProperties | Add-Member -MemberType NoteProperty -Name "User Principal Name" -Value $applicationInstance.UserPrincipalName
    $resourceAccountProperties | Add-Member -MemberType NoteProperty -Name "Object Id" -Value $applicationInstance.ObjectId
    $resourceAccountProperties | Add-Member -MemberType NoteProperty -Name "Phone Number" -Value $applicationInstance.PhoneNumber

    switch ($applicationInstance.ApplicationId) {

        # Auto Attendants
        "ce933385-9390-45d1-9512-c8d228074e07" {

            $resourceAccountProperties | Add-Member -MemberType NoteProperty -Name "Type" -Value "Auto Attendant"

        }
        # Call Queues
        "11cd3e2e-fccb-42ad-ad00-878b93575e07" {

            $resourceAccountProperties | Add-Member -MemberType NoteProperty -Name "Type" -Value "Call Queue"

        }
        Default {

            $resourceAccountProperties | Add-Member -MemberType NoteProperty -Name "Type" -Value "Other"

        }

    }

    $unassignedResourceAccountsExport += $resourceAccountProperties

}

$unassignedResourceAccountsExport | Export-Csv -Path ".\UnassignedResourceAccounts.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force