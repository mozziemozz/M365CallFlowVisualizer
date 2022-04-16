# import required functions from this repository
. .\Functions\Connect-M365CFV.ps1
. .\Functions\Get-TeamsUserCallFlow.ps1

$allTeamsUsers = Get-CsOnlineUser -Filter {accountEnabled -eq $true}

$teamsUserCounter = 1

foreach ($teamsUser in $allTeamsUsers) {

    Write-Host "Visualizing Teams User Calling Settings for user '$($teamsUser.DisplayName)'. User $teamsUserCounter/$($allTeamsUsers.Count)"

    . Get-TeamsUserCallFlow -UserPrincipalName $teamsUser.UserPrincipalName -PreviewSvg $false -SetClipBoard $false

    $teamsUserCounter ++

}

