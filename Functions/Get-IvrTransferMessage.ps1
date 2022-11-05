# Please wait while your Call is being transferred.
# 1 = Transfer Message
# 2 = Transfer Message, System Message
# 3 = Transfer Message
# 4 = Transfer Message
# 5 = Let me transfer you to the operator

<#
    .SYNOPSIS
    Sets the IVR Transfer Message / Greeting based on an auto attendants languageId.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function Get-IvrTransferMessage {
    param (
    )
    
    switch ($languageId) {

        # English (United States)
        en-US {
            $transferGreetingText = "Please wait while your Call is being transferred."
            $transferGreetingFriendly = "Please wait while your Call is being transferred."
        }


        Default {
            $transferGreetingText = "Please wait while your Call is being transferred. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $transferGreetingFriendly = "Please wait while your Call is being transferred. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

    }

    if ($defaultCallFlowAction -eq "TransferCallToOperator" -or $afterHoursCallFlowAction -eq "TransferCallToOperator") {

        switch ($languageId) {

            # English (United States)
            en-US {
                $transferGreetingOperatorText = "Let me transfer you to the operator."
                $transferGreetingOperatorFriendly = "Let me transfer you to the operator."
            }
    
    
            Default {
                $transferGreetingOperatorText = "Let me transfer you to the operator."
                $transferGreetingOperatorFriendly = "Let me transfer you to the operator. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            }
    
        }    

    }

    else {

        $transferGreetingOperatorText = $null
        $transferGreetingOperatorFriendly = $null

    }

    return $transferGreetingText, $transferGreetingFriendly, $transferGreetingOperatorText, $transferGreetingOperatorFriendly
}