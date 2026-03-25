<#
    .SYNOPSIS
    Gets the MS System Message / Greeting based on an auto attendant's or call queue's languageId.

    .DESCRIPTION
    Author:             Martin Heusser
    Contributors:       Patrik Lleshaj (https://github.com/realgarit)
    Version:            2.0.0
    Changelog:          .\Changelog.md

#>

function Get-MsSystemMessage {
    param (
    )

    # Languages with known translations
    $knownGreetings = @{
        'de-DE' = 'Bitte hinterlassen Sie eine Nachricht nach dem Ton. Wenn Sie fertig sind, legen Sie bitte auf.'
        'en-AU' = 'Please leave a message after the tone. When you have finished, please hang up.'
        'en-CA' = 'Please leave a message after the tone. When you have finished, please hang up.'
        'en-GB' = 'Please leave a message after the tone. When you have finished, please hang up.'
        'en-IN' = 'Please leave a message after the tone. When you have finished, please hang up.'
        'en-US' = 'Please leave a message after the tone. When you have finished, please hang up.'
        'sv-SE' = 'Lämna ett meddelande efter tonen. När du är klar lägger du på.'
        'cs-CZ' = 'Po zaznění tónu prosím zanechte vzkaz, na závěr zavěste.'
    }

    # Friendly versions (without diacritics where needed)
    $knownGreetingsFriendly = @{
        'de-DE' = 'Bitte hinterlassen Sie eine Nachricht nach dem Ton. Wenn Sie fertig sind, legen Sie bitte auf.'
        'en-AU' = 'Please leave a message after the tone. When you have finished, please hang up.'
        'en-CA' = 'Please leave a message after the tone. When you have finished, please hang up.'
        'en-GB' = 'Please leave a message after the tone. When you have finished, please hang up.'
        'en-IN' = 'Please leave a message after the tone. When you have finished, please hang up.'
        'en-US' = 'Please leave a message after the tone. When you have finished, please hang up.'
        'sv-SE' = 'Lämna ett meddelande efter tonen. När du är klar lägger du pa.'
        'cs-CZ' = 'Po zazneni tonu prosim zanechte vzkaz, na zaver zaveste.'
    }

    $fallbackText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."

    if ($knownGreetings.ContainsKey($languageId)) {
        $systemGreetingText = $knownGreetings[$languageId]
    }
    else {
        $systemGreetingText = $fallbackText
    }

    if ($knownGreetingsFriendly.ContainsKey($languageId)) {
        $systemGreetingTextFriendly = $knownGreetingsFriendly[$languageId]
    }
    else {
        $systemGreetingTextFriendly = $fallbackText
    }

    return $systemGreetingText, $systemGreetingTextFriendly
}
