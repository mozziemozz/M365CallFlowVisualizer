<#
    .SYNOPSIS
    Determines if a User Id is a Resource Account or a normal user.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function Get-AccountType {
    param (
        [Parameter(Mandatory=$false)][String]$Id
    )

    if ($allAutoAttendantIds -contains $Id) {

        $accountType = "VoiceApp"

    }

    elseif ($allCallQueueIds -contains $Id) {

        $accountType = "VoiceApp"

    }

    else {

        $accountType = "UserAccount"

    }
    
}