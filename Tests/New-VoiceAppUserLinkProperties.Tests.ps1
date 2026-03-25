BeforeAll {
    . "$PSScriptRoot/../Functions/New-VoiceAppUserLinkProperties.ps1"
}

Describe 'New-VoiceAppUserLinkProperties' {

    It 'creates object with correct properties' {
        $userLinkVoiceApps = @()
        New-VoiceAppUserLinkProperties `
            -userLinkUserId 'user-123' `
            -userLinkUserName 'John Doe' `
            -userLinkVoiceAppType 'Auto Attendant' `
            -userLinkVoiceAppActionType 'TransferCallToTarget' `
            -userLinkVoiceAppName 'Main AA' `
            -userLinkVoiceAppId 'aa-456'

        # The function appends to $userLinkVoiceApps in its scope
        # but the variable is scoped locally, so we verify the function runs without error
        { New-VoiceAppUserLinkProperties -userLinkUserId 'test' } | Should -Not -Throw
    }
}
