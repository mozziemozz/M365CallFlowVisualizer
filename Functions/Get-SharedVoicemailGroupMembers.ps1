<#
    .SYNOPSIS
    Reads group members of an M365 Group.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.2
    Changelog:          .\Changelog.md

#>

function Get-SharedVoicemailGroupMembers {
    param (
        [Parameter(Mandatory=$false)][String]$SharedVoicemailGroupId
    )

    Write-Host "Shared Voicemail Group Id: $SharedVoicemailGroupId" -ForegroundColor Magenta

    $sharedVoicemailGroupMembers = Get-MgGroupMember -GroupId $SharedVoicemailGroupId

    if ($ShowSharedVoicemailGroupSubscribers -eq $true) {

        $sharedVoicemailGroupSubscribers = Get-UnifiedGroupLinks -Identity $SharedVoicemailGroupId -LinkType Subscribers -ResultSize Unlimited

    }

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

        if ($ShowSharedVoicemailGroupSubscribers -eq $true) {

            if ($sharedVoicemailGroupMember.Id -in $sharedVoicemailGroupSubscribers.ExternalDirectoryObjectId) {

                $currentMember = "$currentMember, Follow In Inbox: TRUE"

            }

            else {

                $currentMember = "$currentMember, Follow In Inbox: FALSE"

            }

        }
        
        $mdSharedVoicemailGroupMembers += "<br>$currentMember"

    }

    #return $mdSharedVoicemailGroupMembers

}