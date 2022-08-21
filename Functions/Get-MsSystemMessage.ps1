<#
    .SYNOPSIS
    Sets the MS System Message / Greeting based on an auto attendants or call queues languageId.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.2
    Changelog:          .\Changelog.md

#>

function Get-MsSystemMessage {
    param (
    )
    
    switch ($languageId) {

        # Arabic (Egypt)
        ar-EG {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Catalan (Catalan)
        ca-ES {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Danish (Denmark)
        da-DK {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # German (Germany)
        de-DE {
            $systemGreetingText = "Bitte hinterlassen Sie eine Nachricht nach dem Ton. Wenn Sie fertig sind, legen Sie bitte auf."
            $systemGreetingTextFriendly = "Bitte hinterlassen Sie eine Nachricht nach dem Ton. Wenn Sie fertig sind, legen Sie bitte auf."
        }

        # English (Australia)
        en-AU {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up."
        }

        # English (Canada)
        en-CA {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up."
        }

        # English (United Kingdom)
        en-GB {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up."
        }

        # English (India)
        en-IN {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up."
        }

        # English (United States)
        en-US {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up."
        }

        # Spanish (Spain)
        es-ES {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Spanish (Mexico)
        es-MX {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Finnish (Finland)
        fi-FI {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # French (Canada)
        fr-CA {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # French (France)
        fr-FR {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Italian (Italy)
        it-IT {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Japanese (Japan)
        ja-JP {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Korean (Korea)
        ko-KR {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Norwegian, Bokmål (Norway)
        nb-NO {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Dutch (Netherlands)
        nl-NL {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Polish (Poland)
        pl-PL {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Portuguese (Portugal)
        pt-PT {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Portuguese (Brazil)
        pt-BR {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Russian (Russia)
        ru-RU {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Swedish (Sweden)
        sv-SE {
            $systemGreetingText = "Lämna ett meddelande efter tonen. När du är klar lägger du på."
            $systemGreetingTextFriendly = "Lämna ett meddelande efter tonen. När du är klar lägger du pa."
        }

        # Chinese (Simplified, PRC)
        zh-CN {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Chinese (Traditional, Hong Kong S.A.R.)
        zh-HK {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Chinese (Traditional, Taiwan)
        zh-TW {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Turkish (Turkey)
        tr-TR {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Czech (Czech Republic)
        cs-CZ {
            $systemGreetingText = "Po zaznění tónu prosím zanechte vzkaz, na závěr zavěste."
            $systemGreetingTextFriendly = "Po zazneni tonu prosim zanechte vzkaz, na zaver zaveste."
        }

        # Thai (Thai)
        th-TH {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Greek (Greek)
        el-GR {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Hungarian (Hungary)
        hu-HU {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Slovak (Slovakia)
        sk-SK {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Croatian (Croatia)
        hr-HR {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Slovenian (Slovenia)
        sl-SI {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Indonesian (Indonesia)
        id-ID {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Romanian (Romania)
        ro-RO {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        # Vietnamese (Vietnam)
        vi-VN {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

        Default {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
            $systemGreetingTextFriendly = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }

    }
    return $systemGreetingText, $systemGreetingTextFriendly
}