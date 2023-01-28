<#
    .SYNOPSIS
    Reads group members of an M365 Group.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.1
    Changelog:          .\Changelog.md

#>

function Get-SharedVoicemailGroupMembers {
    param (
        [Parameter(Mandatory=$false)][String]$SharedVoicemailGroupId
    )

    Write-Host "Shared Voicemail Group Id: $SharedVoicemailGroupId" -ForegroundColor Magenta

    $sharedVoicemailGroupMembers = Get-MgGroupMember -GroupId $SharedVoicemailGroupId

    $mdSharedVoicemailGroupMembers = "<br><br>Members"

    foreach ($sharedVoicemailGroupMember in $sharedVoicemailGroupMembers) {

        $currentMember = (Get-MgUser -UserId $sharedVoicemailGroupMember.Id).Mail

        if ($ObfuscatePhoneNumbers -eq $true) {

            $domainLength = $currentMember.Split("@")[-1].Length
            $topLevelDomainLength = $currentMember.Split("@")[-1].Split(".")[-1].Length

            $obfuscatedTopLevelDomain = "*"

            do {
                $obfuscatedTopLevelDomain += "*"
            } until (
                $obfuscatedTopLevelDomain.Length -eq $topLevelDomainLength
            )

            $emailLength = $currentMember.Split("@")[0].Length

            $currentMember = "$($currentMember.Split("@")[0].Remove($currentMember.Split("@")[0].Length - ($emailLength -2)))*****@$($currentMember.Split("@")[-1].Remove($currentMember.Split("@")[-1].Length - ($domainLength -2)))*****.$obfuscatedTopLevelDomain"

        }

        $currentMember = $currentMember.Replace("@"," at ")
        
        $mdSharedVoicemailGroupMembers += "<br>$currentMember"

    }

    #return $mdSharedVoicemailGroupMembers

}