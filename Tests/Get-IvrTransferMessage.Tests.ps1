BeforeAll {
    . "$PSScriptRoot/../Functions/Get-IvrTransferMessage.ps1"
}

Describe 'Get-IvrTransferMessage' {

    Context 'known languages' {

        It 'returns English transfer greeting for en-US' {
            $languageId = 'en-US'
            $defaultCallFlowAction = 'SomeAction'
            $afterHoursCallFlowAction = 'SomeAction'
            $result = Get-IvrTransferMessage
            $result[0] | Should -Be 'Please wait while your Call is being transferred.'
            $result[1] | Should -Be 'Please wait while your Call is being transferred.'
        }

        It 'returns English transfer greeting for en-GB' {
            $languageId = 'en-GB'
            $defaultCallFlowAction = 'SomeAction'
            $afterHoursCallFlowAction = 'SomeAction'
            $result = Get-IvrTransferMessage
            $result[0] | Should -Be 'Please wait while your Call is being transferred.'
        }
    }

    Context 'unsupported languages' {

        It 'returns fallback with hint for unknown language' {
            $languageId = 'fr-FR'
            $defaultCallFlowAction = 'SomeAction'
            $afterHoursCallFlowAction = 'SomeAction'
            $result = Get-IvrTransferMessage
            $result[0] | Should -BeLike "*not supported*"
            $result[0] | Should -BeLike "*fr-FR*"
        }
    }

    Context 'operator transfer handling' {

        It 'returns operator greeting when default action is TransferCallToOperator' {
            $languageId = 'en-US'
            $defaultCallFlowAction = 'TransferCallToOperator'
            $afterHoursCallFlowAction = 'SomeAction'
            $result = Get-IvrTransferMessage
            $result[2] | Should -Be 'Let me transfer you to the operator.'
            $result[3] | Should -Be 'Let me transfer you to the operator.'
        }

        It 'returns operator greeting when afterHours action is TransferCallToOperator' {
            $languageId = 'en-US'
            $defaultCallFlowAction = 'SomeAction'
            $afterHoursCallFlowAction = 'TransferCallToOperator'
            $result = Get-IvrTransferMessage
            $result[2] | Should -Be 'Let me transfer you to the operator.'
        }

        It 'returns null operator greetings when no operator transfer' {
            $languageId = 'en-US'
            $defaultCallFlowAction = 'SomeAction'
            $afterHoursCallFlowAction = 'SomeAction'
            $result = Get-IvrTransferMessage
            $result[2] | Should -BeNullOrEmpty
            $result[3] | Should -BeNullOrEmpty
        }
    }

    Context 'return value structure' {

        It 'returns exactly four values' {
            $languageId = 'en-US'
            $defaultCallFlowAction = 'TransferCallToOperator'
            $afterHoursCallFlowAction = 'SomeAction'
            $result = Get-IvrTransferMessage
            $result.Count | Should -Be 4
        }
    }
}
