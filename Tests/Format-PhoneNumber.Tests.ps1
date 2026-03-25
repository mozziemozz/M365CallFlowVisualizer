BeforeAll {
    . "$PSScriptRoot/../Functions/Format-PhoneNumber.ps1"
}

Describe 'Format-PhoneNumber' {

    It 'strips tel: prefix' {
        Format-PhoneNumber -PhoneNumberId 'tel:+15551234567' | Should -Be '+15551234567'
    }

    It 'returns number as-is without tel: prefix' {
        Format-PhoneNumber -PhoneNumberId '+15551234567' | Should -Be '+15551234567'
    }

    It 'obfuscates last 4 digits when requested' {
        Format-PhoneNumber -PhoneNumberId 'tel:+15551234567' -Obfuscate $true | Should -Be '+1555123****'
    }

    It 'does not obfuscate by default' {
        Format-PhoneNumber -PhoneNumberId 'tel:+15551234567' | Should -Be '+15551234567'
    }

    It 'handles short numbers with obfuscation' {
        Format-PhoneNumber -PhoneNumberId 'tel:1234' -Obfuscate $true | Should -Be '****'
    }

    It 'handles numbers shorter than 4 digits with obfuscation gracefully' {
        Format-PhoneNumber -PhoneNumberId 'tel:12' -Obfuscate $true | Should -Be '12'
    }

    It 'handles empty tel: URI' {
        Format-PhoneNumber -PhoneNumberId 'tel:' | Should -Be ''
    }
}
