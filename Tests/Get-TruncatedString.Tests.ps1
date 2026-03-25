BeforeAll {
    . "$PSScriptRoot/../Functions/Get-TruncatedString.ps1"
}

Describe 'Get-TruncatedString' {

    Context 'basic truncation' {

        It 'returns string unchanged when shorter than max' {
            Get-TruncatedString -InputString 'Hello' -MaxLength 10 | Should -Be 'Hello'
        }

        It 'returns string unchanged when exactly max length' {
            Get-TruncatedString -InputString 'Hello' -MaxLength 5 | Should -Be 'Hello'
        }

        It 'truncates and adds ellipsis when longer than max' {
            Get-TruncatedString -InputString 'Hello World this is long' -MaxLength 10 | Should -Be 'Hello Worl...'
        }
    }

    Context 'with file extension preservation' {

        It 'preserves .wav extension when truncating' {
            $result = Get-TruncatedString -InputString 'VeryLongAudioFileName.wav' -MaxLength 10 -PreserveExtension
            $result | Should -BeLike '*.wav'
            $result | Should -BeLike '*...*'
        }

        It 'preserves .mp3 extension when truncating' {
            $result = Get-TruncatedString -InputString 'VeryLongAudioFileName.mp3' -MaxLength 10 -PreserveExtension
            $result | Should -BeLike '*.mp3'
        }

        It 'does not preserve extension when string is short enough' {
            Get-TruncatedString -InputString 'short.wav' -MaxLength 20 -PreserveExtension | Should -Be 'short.wav'
        }

        It 'handles files with no extension' {
            $result = Get-TruncatedString -InputString 'VeryLongFileNameWithNoExtension' -MaxLength 10 -PreserveExtension
            $result | Should -Be 'VeryLongFi...'
        }

        It 'does not treat long pseudo-extensions as file extensions' {
            $result = Get-TruncatedString -InputString 'filename.longextension' -MaxLength 8 -PreserveExtension
            $result | Should -Be 'filename...'
        }
    }
}
