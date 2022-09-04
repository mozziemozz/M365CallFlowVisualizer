<#
    .SYNOPSIS
    Reads group members of an M365 Group.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function Get-SharedVoicemailGroupMembers {
    param (
        [Parameter(Mandatory=$false)][String]$SharedVoicemailGroupId
    )

    $sharedVoicemailGroupMembers = Get-MgGroupMember -GroupId $SharedVoicemailGroupId

    $mdSharedVoicemailGroupMembers = "<br><br>Members"

    foreach ($sharedVoicemailGroupMember in $sharedVoicemailGroupMembers) {

        $currentMember = (Get-MgUser -UserId $sharedVoicemailGroupMember.Id).Mail.Replace("@"," at ")
        
        $mdSharedVoicemailGroupMembers += "<br>$currentMember"

    }

    #return $mdSharedVoicemailGroupMembers

}