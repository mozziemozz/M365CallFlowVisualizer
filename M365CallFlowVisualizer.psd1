@{
    RootModule        = 'M365CallFlowVisualizer.psm1'
    ModuleVersion     = '3.2.2'
    GUID              = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'
    Author            = 'Martin Heusser, Patrik Lleshaj'
    CompanyName       = 'heusser.pro'
    Copyright         = '(c) Martin Heusser, Patrik Lleshaj. All rights reserved.'
    Description       = 'Visualizes Microsoft 365 Phone System call flows (Auto Attendants and Call Queues) as Mermaid diagrams.'

    PowerShellVersion = '5.1'

    # Required modules (checked at runtime by Connect-M365CFV):
    # - MicrosoftTeams >= 4.9.3
    # - Microsoft.Graph.Users >= 1.9.6
    # - Microsoft.Graph.Groups >= 1.9.6
    RequiredModules   = @()

    FunctionsToExport = @(
        'Connect-M365CFV'
        'Connect-MsTeamsServicePrincipal'
        'Find-CallQueueAndAutoAttendantUserLinks'
        'Get-AccountType'
        'Get-AllVoiceAppsAndResourceAccounts'
        'Get-AllVoiceAppsAndResourceAccountsAppAuth'
        'Get-AutoAttendantDirectorySearchConfig'
        'Get-AutoAttendantHolidayCallFlow'
        'Get-CallQueueAgentsStatus'
        'Get-IvrTransferMessage'
        'Get-MsSystemMessage'
        'Get-SharedVoicemailGroupMembers'
        'Get-TeamsUserCallFlow'
        'New-VoiceAppUserLinkProperties'
        'Optimize-DisplayName'
        'Read-BusinessHours'
        'Format-PhoneNumber'
        'Get-TruncatedString'
        'Get-MZZSecureCreds'
        'New-MZZEncryptedPassword'
        'Get-MZZTenantIdTxt'
        'Get-MZZAppIdTxt'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()

    PrivateData       = @{
        PSData = @{
            Tags       = @('Microsoft365', 'Teams', 'CallFlow', 'AutoAttendant', 'CallQueue', 'Mermaid', 'Visualization')
            LicenseUri = 'https://github.com/mozziemozz/M365CallFlowVisualizer/blob/main/LICENSE'
            ProjectUri = 'https://github.com/mozziemozz/M365CallFlowVisualizer'
        }
    }
}
