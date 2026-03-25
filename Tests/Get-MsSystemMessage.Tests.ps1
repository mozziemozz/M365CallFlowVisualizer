BeforeAll {
    . "$PSScriptRoot/../Functions/Get-MsSystemMessage.ps1"
}

Describe 'Get-MsSystemMessage' {

    Context 'known languages with translations' {

        It 'returns German greeting for de-DE' {
            $languageId = 'de-DE'
            $result = Get-MsSystemMessage
            $result[0] | Should -Be 'Bitte hinterlassen Sie eine Nachricht nach dem Ton. Wenn Sie fertig sind, legen Sie bitte auf.'
            $result[1] | Should -Be 'Bitte hinterlassen Sie eine Nachricht nach dem Ton. Wenn Sie fertig sind, legen Sie bitte auf.'
        }

        It 'returns English greeting for en-US' {
            $languageId = 'en-US'
            $result = Get-MsSystemMessage
            $result[0] | Should -Be 'Please leave a message after the tone. When you have finished, please hang up.'
            $result[1] | Should -Be 'Please leave a message after the tone. When you have finished, please hang up.'
        }

        It 'returns English greeting for en-GB' {
            $languageId = 'en-GB'
            $result = Get-MsSystemMessage
            $result[0] | Should -Be 'Please leave a message after the tone. When you have finished, please hang up.'
        }

        It 'returns Swedish greeting for sv-SE' {
            $languageId = 'sv-SE'
            $result = Get-MsSystemMessage
            $result[0] | Should -BeLike '*meddelande efter tonen*'
        }

        It 'returns Czech greeting for cs-CZ' {
            $languageId = 'cs-CZ'
            $result = Get-MsSystemMessage
            $result[0] | Should -BeLike '*zanechte vzkaz*'
        }

        It 'returns friendly version without diacritics for cs-CZ' {
            $languageId = 'cs-CZ'
            $result = Get-MsSystemMessage
            $result[1] | Should -Be 'Po zazneni tonu prosim zanechte vzkaz, na zaver zaveste.'
        }
    }

    Context 'unsupported languages fall back to hint message' {

        It 'returns fallback with language hint for fr-FR' {
            $languageId = 'fr-FR'
            $result = Get-MsSystemMessage
            $result[0] | Should -BeLike "*not supported by M365 Call Flow Visualizer*"
            $result[0] | Should -BeLike "*fr-FR*"
        }

        It 'returns fallback for ja-JP' {
            $languageId = 'ja-JP'
            $result = Get-MsSystemMessage
            $result[0] | Should -BeLike "*not supported*"
        }

        It 'returns fallback for completely unknown language' {
            $languageId = 'xx-XX'
            $result = Get-MsSystemMessage
            $result[0] | Should -BeLike "*not supported*"
            $result[0] | Should -BeLike "*xx-XX*"
        }
    }

    Context 'return value structure' {

        It 'returns exactly two values' {
            $languageId = 'en-US'
            $result = Get-MsSystemMessage
            $result.Count | Should -Be 2
        }
    }
}
