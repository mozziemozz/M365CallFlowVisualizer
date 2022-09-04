function Get-SharedVoicemailGroupMembers {
    param (
        [Parameter(Mandatory=$false)][String]$Id
    )

    $sharedVoicemailGroupMembers = Get-MgGroupMember -GroupId $Id

    foreach ($sharedVoicemailGroupMember in $sharedVoicemailGroupMembers) {

        $currentMember = (Get-MgUser -UserId $sharedVoicemailGroupMember.Id).Mail.Replace("@","'AT'")
        
        $currentMember

    }

}