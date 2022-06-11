<#
    .SYNOPSIS
    Sets the MS System Message / Greeting based on an auto attendants or call queues languageId.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function Get-MsSystemMessage {
    param (
    )
    
    switch ($languageId) {
        en-US {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up."
        }
        en-AU {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up."
        }
        en-GB {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up."
        }
        en-CA {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up."
        }
        en-IN {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up."
        }
        de-DE {
            $systemGreetingText = "Bitte hinterlassen Sie eine Nachricht nach dem Ton. Wenn Sie fertig sind, legen Sie bitte auf."
        }
        Default {
            $systemGreetingText = "Please leave a message after the tone. When you have finished, please hang up. Hint: This greeting will be synthesized in '$languageId'. However, this language is not supported by M365 Call Flow Visualizer yet. If you would like to help and provide a transcript of this message in your language, please reach out to me."
        }
    }
    return $systemGreetingText
}