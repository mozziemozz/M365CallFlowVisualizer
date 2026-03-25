<#
    .SYNOPSIS
    Gets the IVR Transfer Message / Greeting based on an auto attendant's languageId.

    .DESCRIPTION
    Author:             Martin Heusser
    Contributors:       Patrik Lleshaj (https://github.com/realgarit)
    Version:            1.1.0
    Changelog:          .\Changelog.md

#>

function Get-IvrTransferMessage {
    param (
    )

    $knownTransferGreetings = @{
        'en-US' = 'Please wait while your Call is being transferred.'
        'en-GB' = 'Please wait while your Call is being transferred.'
    }

    $knownOperatorGreetings = @{
        'en-US' = 'Let me transfer you to the operator.'
    }

    $fallbackTransfer = "Please wait while your Call is being transferred. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
    $fallbackOperator = "Let me transfer you to the operator. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."

    if ($knownTransferGreetings.ContainsKey($languageId)) {
        $transferGreetingText = $knownTransferGreetings[$languageId]
        $transferGreetingFriendly = $knownTransferGreetings[$languageId]
    }
    else {
        $transferGreetingText = $fallbackTransfer
        $transferGreetingFriendly = $fallbackTransfer
    }

    if ($defaultCallFlowAction -eq "TransferCallToOperator" -or $afterHoursCallFlowAction -eq "TransferCallToOperator") {
        if ($knownOperatorGreetings.ContainsKey($languageId)) {
            $transferGreetingOperatorText = $knownOperatorGreetings[$languageId]
            $transferGreetingOperatorFriendly = $knownOperatorGreetings[$languageId]
        }
        else {
            $transferGreetingOperatorText = "Let me transfer you to the operator."
            $transferGreetingOperatorFriendly = $fallbackOperator
        }
    }
    else {
        $transferGreetingOperatorText = $null
        $transferGreetingOperatorFriendly = $null
    }

    return $transferGreetingText, $transferGreetingFriendly, $transferGreetingOperatorText, $transferGreetingOperatorFriendly
}
