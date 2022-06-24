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