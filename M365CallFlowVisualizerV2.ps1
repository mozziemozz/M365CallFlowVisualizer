<#
    .SYNOPSIS
    Reads the configuration from a Microsoft 365 Phone System auto attendant or call queue and visualizes the call flow using mermaid-js.

    .DESCRIPTION
    Reads the configuration from a Microsoft 365 Phone System auto attendant or call queue by either specifying the voice app name and type, unique identity of the voice app or presents a selection of available auto attendants or call queues if none of the identifiers are supplied.
    The call flow is then written into either a mermaid (*.mmd) or a markdown (*.md) file containing the mermaid syntax.

    Author:             Martin Heusser
    Version:            3.2.2
    Changelog:          Moved to repository at .\Changelog.md
    Repository:         https://github.com/mozziemozz/M365CallFlowVisualizer
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro

    .PARAMETER Name
    -Identity
        Specifies the identity of the first / top-level voice app
        Required:           false
        Type:               string
        Accepted values:    unique identifier of an auto attendant or call queue (not resource account) run Get-CsAutoAttendant or Get-CsCallQueue in order to retrieve an identity.
        Default value:      none

    -SetClipBoard
        Specifies if the mermaid code should be copied to the clipboard after the script has finished.
        Required:           false
        Type:               boolean
        Default value:      false
    
    -SaveToFile
        Specifies if the mermaid code should be saved into either a mermaid or markdown file.
        Required:           false
        Type:               boolean
        Default value:      true

    -ExportHtml
        Specifies if, in addition to the Markdown or Mermaid file, also a *.htm file should be exported
        Required:           false
        Type:               boolean
        Default value:      true

    -ExportPng
        Specifies if, the mermaid or markdown file should be converted and saved as a PNG file. This requires Node.JS and @mermaid-js/mermaid-cli + mermaid packages.
        Required:           false
        Type:               boolean
        Default value:      false

    -ExportPDF
        Specifies if, the mermaid or markdown file should be converted and saved as a PDF file. This requires Node.JS and @mermaid-js/mermaid-cli + mermaid packages.
        Required:           false
        Type:               boolean
        Default value:      false

    -PreviewHtml
        Specifies if the exported html file should be opened in default / last active browser (only works on Windows systems)
        Required:           false
        Type:               switch
        Default value:      false
    
    -DocFxMode
        This switch adds a subfolder in the relative file path of click actions for exported audio file or TTS greetings. It can only be used with a script from another repository. Do not use this manually.
        Required:           false
        Type:               switch
        Default value:      false

    -CacheResults
        Specifies if the script should read all auto attendants and call queues in each run. If set to true, the script will read the data from the global variables, if they are not null or empty. If set to false, the script will read the data with each run.
        Required:           false
        Type:               boolean
        Default value:      false

    -CustomFilePath
        Specifies the file path for the output file. The directory must already exist.
        Required:           false
        Type:               string
        Accepted values:    file paths e.g. "C:\Temp"
        Default value:      ".\" (current folder)

    -ShowNestedCallFlows
        Specifies whether or not to also display the call flows of nested call queues or auto attendants. If set to false, only the name of nested voice apps will be rendered. Nested call flows won't be expanded.
        Required:           false
        Type:               boolean
        Default value:      true

    -ShowNestedHolidayCallFlows
        Specifies whether or not to also display the call flows of nested call queues or auto attendants of holiday call handlings. Call flows will be expanded and linked to the holiday subgraph. To use this parameter, -ShowNestedCallFlows must be true.
        Required:           false
        Type:               boolean
        Default value:      false

    -ShowNestedHolidayIVRs
        Specifies whether or not to also display nested IVRs on Holiday Call Handlings. They will render in the main diagram and not in the Holiday subgraph.
        Required:           false
        Type:               boolean
        Default value:      false

    -ShowUserCallingSettings
        Specifies whether or not to also display the user calling settings of a Teams user. If set to false, only the name of a user will be rendered.
        Required:           false
        Type:               boolean
        Default value:      true

    -ShowNestedUserCallGroups
        Specifies if call groups of users should be expanded and included into the diagram.
        Required:           false
        Type:               boolean
        Default value:      false

    -ShowNestedUserDelegates
        Specifies if delegates s of users should be expanded and included into the diagram.
        Required:           false
        Type:               boolean
        Default value:      false

    -ShowTransferCallToTargetType
        Specifies if TransferCallToTarget nodes that redirect to an auto attendant or a call queue should include the type of the target (Resource Account (ApplicationEndpoint) or Voice App (ConfigurationEndpoint)).
        Required:           false
        Type:               boolean
        Default value:      true

    -CombineDisconnectCallNodes
        Specifies whether or not to only display one mermaid node for all disconnect actions. If this is enabled, diagrams are most likely less readable. This does not apply to "DisconnectCall" actions in holiday call handlings.
        Required:           false
        Type:               boolean
        Default value:      false

    -CombineCallConnectedNodes
        Specifies whether or not to only display one mermaid node for all call connected noed. If this is enabled, diagrams are most likely less readable.
        Required:           false
        Type:               boolean
        Default value:      false

    -ShowCqAgentPhoneNumbers
        Specifies whether or not the agent subgraphs of call queues should include a users direct number.
        Required:           false
        Type:               switch
        Default value:      false   

    -ShowCqAgentOptInStatus
        Specifies whether or not the current opt in status of agents should be displayed.
        Required:           false
        Type:               switch
        Default value:      false

    -ShowPhoneNumberType
        Specifies whether or not the phone number type of phone numbers should be displayed. (CallingPlan, OperatorConnect, DirectRouting)
        Required:           false
        Type:               switch
        Default value:      false   

    -ShowTTSGreetingText
        Specifies whether or not the text of TTS greetings should be included in greeting nodes. Note: this can create wide diagrams. Use parameter -TruncateGreetings to shorten the text.
        Required:           false
        Type:               switch
        Default value:      false
        
    -ShowAudioFileName
        Specifies whether or not the filename of audio file greetings should be included in greeting nodes. Note: this can create wide diagrams. Use parameter -TruncateGreetings to shorten the filename
        Required:           false
        Type:               switch
        Default value:      false

    -TruncateGreetings
        Specifies how many characters of the file name or the greeting text should be included. The default value is 20. This will shorten all greetings and filenames to 20 characters, excluding the file name extension.
        Required:           false
        Type:               single
        Default value:      20

    -ExportAudioFiles
        Specifies if the audio files of greetings, announcements and music on hold should be exported to the specified directory. If this is enabled, Markdown and HTML output will have clickable links on the greeting nodes which open an audio file in the browser. For this to work, you must also use the -ShowAudioFileName switch.
        Required:           false
        Type:               switch
        Default value:      false

    -ExportTTSGreetings
        Specifies if the value of TTS greetings and announcements should be exported to the specified directory. If this is enabled, Markdown and HTML output will have clickable links on the greeting nodes with which open a text file in the browser. For this to work, you must also use the -ShowTTSGreetingText switch.
        Required:           false
        Type:               switch
        Default value:      false

    -OverrideVoiceIdToFemale
        If set to true, all instances of 'Male' voice Id of auto attendants will be overwritten to 'Female'. It's likely that an auto attendants properties say Male even though the aa is using a female voice. This is a bug from MS.
        Required:           false
        Type:               switch
        Default value:      false

    -FindUserLinks
        This paramter can only be used by the external function Find-CallQueueAndAutoAttendantUserLinks. When this parameter is true, the script will write every user which is part of a call flow into an external variable.
        Required:           false
        Type:               switch
        Default value:      false

    -CheckCallFlowRouting
        This paramter will check if any Auto attendant (top-level or nested) is currently open or closed based on holiday and business hours schedule. Output is written to the console. By default, your current system date and time and time zone is used to check in respective Auto Attendant time zone.
        Required:           false
        Type:               switch
        Default value:      false

    -CheckCallFlowRoutingSpecificDateTime
        This paramter can only be used when -CheckCallFlowRouting is $true. Specify any date and time in the future or past to check if an any Auto attendant (top-level or nested) is open or closed based on holiday and business hours schedule on that specific date. Enter the date as string in your local date time format. Example: "19.11.2023 23:59:59"
        Required:           false
        Type:               string
        Default value:      none

    -ObfuscatePhoneNumbers
        Specifies if phone numbers in call flows should be obfuscated for sharing / example reasons. This will replace the last 4 digits in numbers with an asterisk (*) character. Warning: This will only obfuscate phone numbers in node descriptions. The complete phone number will still be included in Markdown, Mermaid and HTML output!
        Required:           false
        Type:               bool
        Default value:      false

    -ShowSharedVoicemailGroupMembers
        Specifies if group members (email) should be shown on a shared voicemail node.
        Required:           false
        Type:               bool
        Default value:      false

    -ShowSharedVoicemailGroupSubscribers
        Specifies if the info if a group member is also following the group mailbox in their personal inbox should be included in the diagram. Requires -ShowSharedVoicemailGroupMembers to be $true
        Required:           false
        Type:               bool
        Default value:      false

    -ShowCqOutboundCallingIds
        Specifies if outbound calling Ids of call queues should be shown
        Required:           false
        Type:               bool
        Default value:      false

    -ShowCqAuthorizedUsers
        Specifies if authorized users of call queues should be shown
        Required:           false
        Type:               bool
        Default value:      false

    -ShowAaAuthorizedUsers
        Specifies if authorized users of auto attendants should be shown
        Required:           false
        Type:               bool
        Default value:      false

    -ShowUserOutboundCallingIds
        Specifies if outbound calling Ids of individual call queue agents should be shown. This only works if -ShowCqAgentPhoneNumbers is $true
        Required:           false
        Type:               bool
        Default value:      false

    -DocType
        Specifies the document type.
        Required:           false
        Type:               string
        Accepted values:    Markdown, Mermaid
        Default value:      Markdown

    -DateFormat
        Specifies the format of the holiday dates. EU = "dd.MM.yyyy HH:mm" US = "MM.dd.yyyy HH:mm"
        Required:           false
        Type:               string
        Accepted values:    EU, US
        Default value:      EU

    -Theme
        Specifies the mermaid theme in Markdown. Custom will use the default hex color codes below if not specified otherwise. Themes are currently only supported for Markdown output.
        Required:           false
        Type:               string
        Accepted values:    default, dark, neutral, forest, custom
        Default value:      default

    -NodeColor
        Specifies a custom color for the node fill
        Required:           false
        Type:               string
        Accepted values:    Hexadecimal color codes starting with # enclosed by ""
        Default value:      "#505AC9"

    -NodeBorderColor
        Specifies a custom color for the node border
        Required:           false
        Type:               string
        Accepted values:    Hexadecimal color codes starting with # enclosed by ""
        Default value:      "#464EB8"

    -FontColor
        Specifies a custom color for the node border
        Required:           false
        Type:               string
        Accepted values:    Hexadecimal color codes starting with # enclosed by ""
        Default value:      "#FFFFFF"

    -LinkColor
        Specifies a custom color for links
        Required:           false
        Type:               string
        Accepted values:    Hexadecimal color codes starting with # enclosed by ""
        Default value:      "#505AC9"

    -LinkTextColor
        Specifies a custom color for text on links
        Required:           false
        Type:               string
        Accepted values:    Hexadecimal color codes starting with # enclosed by ""
        Default value:      "#000000"

    -SubgraphColor
        Specifies a custom color for subgraph backgrounds
        Required:           false
        Type:               string
        Accepted values:    Hexadecimal color codes starting with # enclosed by ""
        Default value:      "#7B83EB"
    
    -VoiceAppName
        If provided, you won't be provided with a selection of available voice apps. The script will search for a voice app with the specified name. This is the display name of a voice app, not a resource account. If you specify the VoiceAppName, VoiceAppType will become mandatory.
        Required:           false
        Type:               string
        Accepted values:    Voice App Name
        Default value:      none

    -VoiceAppType
        This becomes mandatory if VoiceAppName is specified. Because an auto attendant and a call queue could have the same arbitrary name, it is neccessary to also specify the type of the voice app, if no unique identity is specified.
        Required:           true, if VoiceAppName is specified
        Type:               string
        Accepted values:    Auto Attendant, Call Queue
        Default value:      none

    -HardcoreMode
        When this switch is enabled, all parameters which will gather and display additional data are enabled. This will overwrite any other parameters you might have set. Use it to easily explore the scripts capabilities.
        Required:           false
        Type:               switch
        Default value:      false

    -ConnectWithServicePrincipal
        Connect to Teams and Graph using your own Entra ID App Registration
        Required:           false
        Type:               switch
        Default value:      false

    -EntraTenantIdFileName
        Specifies the name of the file to store/read the tenant ID in an encrypted format at .\.local\SecureCreds. If -ConnectWithServicePrincipal is specified/$true this parameter must be specified as well.
        Required:           false
        Type:               string
        Default value:      "m365-cfv-tenant-id"

    -EntraApplicationIdFileName
        Specifies the name of the file to store/read the app ID in an encrypted format at .\.local\SecureCreds. If -ConnectWithServicePrincipal is specified/$true this parameter must be specified as well.
        Required:           false
        Type:               string
        Default value:      "m365-cfv-app-id"

    -EntraClientSecretFileName
        Specifies the name of the file to store/read the client secret in an encrypted format at .\.local\SecureCreds. If -ConnectWithServicePrincipal is specified/$true this parameter must be specified as well.
        Required:           false
        Type:               string
        Default value:      "m365-cfv-client-secret"

    .INPUTS
        None.

    .OUTPUTS
        Files:
            - *.md
            - *.mmd

    .EXAMPLE
        .\M365CallFlowVisualizerV2.ps1

    .LINK
    https://github.com/mozziemozz/M365CallFlowVisualizer
    
#>

# Requires -Modules @{ ModuleName = "MicrosoftTeams"; ModuleVersion = "6.1.0" }, "Microsoft.Graph.Users", "Microsoft.Graph.Groups"

[CmdletBinding(DefaultParametersetName="None")]
param(
    [Parameter(Mandatory = $false)][String]$Identity,
    [Parameter(Mandatory = $false)][Bool]$SetClipBoard = $false,
    [Parameter(Mandatory = $false)][Bool]$SaveToFile = $true,
    [Parameter(Mandatory = $false)][Bool]$ExportHtml = $true,
    [Parameter(Mandatory = $false)][Bool]$ExportPng = $false,
    [Parameter(Mandatory=$false)][Bool]$ExportPDF = $false,
    [Parameter(Mandatory = $false)][Switch]$PreviewHtml,
    [Parameter(Mandatory = $false)][Switch]$DocFxMode,
    [Parameter(Mandatory = $false)][Bool]$CacheResults = $false,
    [Parameter(Mandatory = $false)][String]$CustomFilePath = ".\Output\$(Get-Date -Format "yyyy-MM-dd")",
    [Parameter(Mandatory = $false)][Bool]$ShowNestedCallFlows = $true,
    [Parameter(Mandatory = $false)][Bool]$ShowUserCallingSettings = $true,
    [Parameter(Mandatory = $false)][Bool]$ShowNestedUserCallGroups = $false,
    [Parameter(Mandatory = $false)][Bool]$ShowNestedUserDelegates = $false,
    [Parameter(Mandatory = $false)][Bool]$ShowNestedHolidayCallFlows = $false,
    [Parameter(Mandatory = $false)][Bool]$ShowNestedHolidayIVRs = $false,
    [Parameter(Mandatory = $false)][Bool]$ShowTransferCallToTargetType = $true,
    [Parameter(Mandatory = $false)][Bool]$CombineDisconnectCallNodes = $false,
    [Parameter(Mandatory = $false)][Bool]$CombineCallConnectedNodes = $false,
    [Parameter(Mandatory = $false)][Switch]$ShowCqAgentPhoneNumbers,
    [Parameter(Mandatory = $false)][Switch]$ShowCqAgentOptInStatus,
    [Parameter(Mandatory = $false)][Switch]$ShowPhoneNumberType,
    [Parameter(Mandatory = $false)][Switch]$ShowTTSGreetingText,
    [Parameter(Mandatory = $false)][Switch]$ShowAudioFileName,
    [Parameter(Mandatory = $false)][Single]$TruncateGreetings = 20,
    [Parameter(Mandatory = $false)][Switch]$ExportAudioFiles,
    [Parameter(Mandatory = $false)][Switch]$ExportTTSGreetings,
    [Parameter(Mandatory = $false)][Switch]$OverrideVoiceIdToFemale,
    [Parameter(Mandatory = $false)][Switch]$FindUserLinks,
    [Parameter(Mandatory = $false)][Switch]$CheckCallFlowRouting,
    [Parameter(Mandatory = $false)][String]$CheckCallFlowRoutingSpecificDateTime,
    [Parameter(Mandatory = $false)][Bool]$ObfuscatePhoneNumbers = $false,
    [Parameter(Mandatory = $false)][Bool]$ShowSharedVoicemailGroupMembers = $false,
    [Parameter(Mandatory = $false)][Bool]$ShowSharedVoicemailGroupSubscribers = $false,
    [Parameter(Mandatory = $false)][Bool]$ShowCqOutboundCallingIds = $false,
    [Parameter(Mandatory = $false)][Bool]$ShowUserOutboundCallingIds = $false,
    [Parameter(Mandatory = $false)][Bool]$ShowCqAuthorizedUsers = $false,
    [Parameter(Mandatory = $false)][Bool]$ShowAaAuthorizedUsers = $false,
    [Parameter(Mandatory = $false)][ValidateSet("Markdown","Mermaid")][String]$DocType = "Markdown",
    [Parameter(Mandatory = $false)][ValidateSet("EU","US")][String]$DateFormat = "EU",
    [Parameter(Mandatory = $false)][ValidateSet("default","forest","dark","neutral","custom")][String]$Theme = "default",
    [Parameter(Mandatory = $false)][String]$NodeColor = "#505AC9",
    [Parameter(Mandatory = $false)][String]$NodeBorderColor = "#464EB8",
    [Parameter(Mandatory = $false)][String]$FontColor = "#FFFFFF",
    [Parameter(Mandatory = $false)][String]$LinkColor = "#505AC9",
    [Parameter(Mandatory = $false)][String]$LinkTextColor = "#000000",
    [Parameter(Mandatory = $false)][String]$SubgraphColor = "#7B83EB",
    [Parameter(ParameterSetName="VoiceAppProperties",Mandatory = $false)][String]$VoiceAppName,
    [Parameter(ParameterSetName="VoiceAppProperties",Mandatory = $true)][ValidateSet("Auto Attendant","Call Queue")][String]$VoiceAppType,
    [Parameter(Mandatory = $false)][Switch]$HardcoreMode,
    [Parameter(Mandatory = $false)][Switch]$ConnectWithServicePrincipal,
    [Parameter(Mandatory = $false)][String]$EntraTenantIdFileName = "m365-cfv-tenant-id",
    [Parameter(Mandatory = $false)][String]$EntraApplicationIdFileName = "m365-cfv-app-id",
    [Parameter(Mandatory = $false)][String]$EntraClientSecretFileName = "m365-cfv-client-secret"
)

$ErrorActionPreference = "Continue"

# Load Functions

. .\Functions\Connect-M365CFV.ps1
. .\Functions\Read-BusinessHours.ps1
. .\Functions\Optimize-DisplayName.ps1
. .\Functions\Get-TeamsUserCallFlow.ps1
. .\Functions\Get-MsSystemMessage.ps1
. .\Functions\Get-AccountType.ps1
. .\Functions\New-VoiceAppUserLinkProperties.ps1
. .\Functions\Get-SharedVoicemailGroupMembers.ps1
. .\Functions\Get-IvrTransferMessage.ps1
. .\Functions\Get-AutoAttendantDirectorySearchConfig.ps1
. .\Functions\Get-AllVoiceAppsAndResourceAccounts.ps1
. .\Functions\Get-AllVoiceAppsAndResourceAccountsAppAuth.ps1
. .\Functions\SecureCredsMgmt.ps1
. .\Functions\Connect-MsTeamsServicePrincipal.ps1

# Connect to MicrosoftTeams and Microsoft.Graph

if ($ConnectWithServicePrincipal) {

    if ($ShowSharedVoicemailGroupSubscribers -eq $true -or $HardcoreMode -eq $true) {

        Write-Warning -Message "Shared voicemail group subscribers are not supported with ConnectWithServicePrincipal. Please disable this switch."

        Read-Host -Prompt "Press Enter to exit..."

        exit

    }

    Write-Warning -Message "Connecting to Microsoft Teams and Microsoft Graph using your own Entra ID App Registration is not supported yet because it doesn't support [Get|Set|New|Sync]-CsOnlineApplicationInstance yet. Please don't use this parameter yet. https://learn.microsoft.com/en-us/MicrosoftTeams/teams-powershell-application-authentication#cmdlets-supported"

    $checkGetCsOnlineUser = Get-CsOnlineUser -ResultSize 1 -ErrorAction SilentlyContinue

    if (!$checkGetCsOnlineUser) {

        . Get-MZZTenantIdTxt -FileName $EntraTenantIdFileName
        . Get-MZZAppIdTxt -FileName $EntraApplicationIdFileName
        . Get-MZZSecureCreds -FileName $EntraClientSecretFileName -NoClipboard > $null
        $AppSecret = $passwordDecrypted

        . Connect-MsTeamsServicePrincipal -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

        $graphTokenSecureString = $graphToken | ConvertTo-SecureString -AsPlainText -Force

        $graphTokenSecureString = $graphToken | ConvertTo-SecureString -AsPlainText -Force

        Connect-MgGraph -AccessToken $graphTokenSecureString


    }

}

else {

    . Connect-M365CFV

}

if ($HardcoreMode -eq $true) {

    if ($ConnectWithServicePrincipal -eq $true) {

        Write-Warning -Message "Hardcore Mode is not supported with ConnectWithServicePrincipal. Please disable this switch."

        Read-Host -Prompt "Press Enter to exit..."

        exit

    }

    Write-Host "Hardcore Mode is enabled. This means that all options will be enabled and included in the output. This may overwrite individual values you set in other parameters!" -ForegroundColor Magenta

    $ShowNestedHolidayCallFlows = $true
    $ShowNestedHolidayIVRs = $true
    $ShowCqAgentPhoneNumbers = $true
    $ShowCqAgentOptInStatus = $true
    $ShowPhoneNumberType = $true
    $ShowTTSGreetingText = $true
    $ShowAudioFileName = $true
    $TruncateGreetings = 999
    $ExportAudioFiles = $true
    $ExportTTSGreetings = $true
    $ShowSharedVoicemailGroupMembers = $true
    $ShowSharedVoicemailGroupSubscribers = $true
    $ShowCqOutboundCallingIds = $true
    $ShowUserOutboundCallingIds = $true
    $ShowCqAuthorizedUsers = $true
    $ShowAaAuthorizedUsers = $true
    $PreviewHtml = $true
    $ExportPng = $true
    $OverrideVoiceIdToFemale = $true
    $Theme = "dark"

}

Write-Warning -Message "There is currently a bug in MicrosoftTeams PowerShell. Auto attendants are likely to output 'Male' as VoiceId property when queried via PowerShell. `nPlease call your auto attendant's phone number to confirm the voice Id it's using. Use the -OverrideVoiceIdToFemale switch param to change all 'Male' values to 'Female' in diagram output."

if ($SaveToFile -eq $false -and $CustomFilePath -ne ".\Output") {

    Write-Warning -Message "Custom file path is specified but SaveToFile is set to false. The call flow won't be saved!"

}

if ($ObfuscatePhoneNumbers -eq $true) {

    Write-Warning -Message "Obfuscate phone numbers is True. This will only obfuscate phone numbers in node descriptions. The complete phone number will still be included in Markdown, Mermaid and HTML output!"

}

if (($ExportPng -eq $true) -or ($ExportPDF -eq $true)) {

    try {
        $ErrorActionPreference = "SilentlyContinue"
        $checkNPM = npm list -g
        $ErrorActionPreference = "Continue"
    }
    catch {
        Write-Warning -Message "Node.JS is not installed. Please install Node.JS for PNG output.`nwinget install --id=OpenJS.NodeJS  -e"
    }

    try {
        $ErrorActionPreference = "SilentlyContinue"
        $checkMmdcPackage = mmdc --version
        $ErrorActionPreference = "Continue"
    }
    catch {
        Write-Warning -Message "mermaid npm packages is not installed. Please install mermaid npm packages for PNG output. `nnpm install -g @mermaid-js/mermaid-cli"
        $ExportPng = $false
        $ExportPDF = $false
    }

}

# Common arrays and variables
$nestedVoiceApps = @()
$processedVoiceApps = @()
$allMermaidNodes = @()
$allSubgraphs = @()
$audioFileNames = @()
$ttsGreetings = @()
$allCallFlowRoutingChecks = @()

# Get all voice apps and resource accounts from external function

if ($ConnectWithServicePrincipal) {

    . Get-AllVoiceAppsAndResourceAccountsAppAuth

}

else {

    . Get-AllVoiceAppsAndResourceAccounts

}

$allAutoAttendantIds = $allAutoAttendants.Identity
$allCallQueueIds = $allCallQueues.Identity

$applicationIdAa = "ce933385-9390-45d1-9512-c8d228074e07"
#$applicationIdCq = "11cd3e2e-fccb-42ad-ad00-878b93575e07" # not in use at the moment

switch ($DateFormat) {
    EU {
        $dateFormatString = "dd.MM.yyyy HH:mm"
    }
    US {
        $dateFormatString = "MM.dd.yyyy HH:mm"
    }
    Default {
        $dateFormatString = "dd.MM.yyyy HH:mm"
    }
}

if ($CustomFilePath) {

    $FilePath = $CustomFilePath

    if (!(Test-Path -Path $FilePath)) {

        New-Item -Path $FilePath -ItemType Directory

    }

}

else {

    $FilePath = "."

}

function Set-Mermaid {
    param (
        [Parameter(Mandatory = $true)][String]$DocType
        )

    if ($Theme -eq "custom") {

        $MarkdownTheme =@"
%%{init: {'maxTextSize': 99000, 'flowchart' : { 'curve' : 'basis' } } }%%

"@ 

    }

    else {

        $Theme = $Theme.ToLower()

        $MarkdownTheme =@"
%%{init: {'theme': '$($Theme)', 'maxTextSize': 99000, 'flowchart' : { 'curve' : 'basis' } } }%%

"@ 

    }


    if ($DocType -eq "Markdown") {

        $mdStart =@"
## CallFlowNamePlaceHolder

``````mermaid
$MarkdownTheme
flowchart TB
"@

        $mdEnd =@"

``````
"@

        $fileExtension = ".md"
    }

    else {
        $mdStart =@"
$MarkdownTheme
flowchart TB
"@

        $mdEnd =@"

"@

        $fileExtension = ".mmd"
    }

    $mermaidCode = @()

    $mermaidCode += $mdStart
    
}

function Find-Holidays {
    param (
        [Parameter(Mandatory = $true)][String]$VoiceAppId

    )

    $aa = $allAutoAttendants | Where-Object {$_.Identity -eq $VoiceAppId}

    if ($aa.CallHandlingAssociations.Type -contains "Holiday") {
        $aaHasHolidays = $true
    }

    else {
        $aaHasHolidays = $false
    }

    if ($aa.VoiceResponseEnabled) {

        $aaIsVoiceResponseEnabled = $true

    }

    else {
        
        $aaIsVoiceResponseEnabled = $false

    }

    # Check if auto attendant has an operator
    $Operator = $aa.Operator

    if ($Operator) {

        switch ($Operator.Type) {

            User { 
                $OperatorTypeFriendly = "User"
                $OperatorUser = (Get-MgUser -UserId $($Operator.Id))
                $OperatorName = Optimize-DisplayName -String $OperatorUser.DisplayName
                $OperatorIdentity = $OperatorUser.Id
                $AddOperatorToNestedVoiceApps = $true
            }
            ExternalPstn { 
                $OperatorTypeFriendly = "External Number"
                $OperatorName = ($Operator.Id).Replace("tel:","")
                $OperatorIdentity = $OperatorName
                $AddOperatorToNestedVoiceApps = $false

                if ($ObfuscatePhoneNumbers -eq $true) {

                    $OperatorName = $OperatorName.Remove(($OperatorName.Length -4)) + "****"

                }

            }
            ApplicationEndpoint {

                # Check if application endpoint is auto attendant or call queue

                $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $Operator.Id -and $_.ApplicationId -eq $applicationIdAa}

                if ($matchingApplicationInstanceCheckAa) {

                    $MatchingOperatorAa = $allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $Operator.Id -or $_.Identity -eq $Operator.Id}

                    $OperatorTypeFriendly = "[Auto Attendant"
                    $OperatorName = "$($MatchingOperatorAa.Name)]"
                    $OperatorIdentity = $MatchingOperatorAa.Identity

                }

                else {

                    $MatchingOperatorCq = $allCallQueues | Where-Object {$_.ApplicationInstances -contains $Operator.Id -or $_.Identity -eq $Operator.Id}

                    $OperatorTypeFriendly = "[Call Queue"
                    $OperatorName = "$($MatchingOperatorCq.Name)]"
                    $OperatorIdentity = $MatchingOperatorCq.Identity

                }

                $AddOperatorToNestedVoiceApps = $true

            }
            ConfigurationEndpoint {

                # Check if application endpoint is auto attendant or call queue

                if ($allAutoAttendantIds -contains $Operator.Id) {

                    $MatchingOperatorAa = $allAutoAttendants | Where-Object { $_.Identity -eq $Operator.Id }

                    $OperatorTypeFriendly = "[Auto Attendant"
                    $OperatorName = "$($MatchingOperatorAa.Name)]"
                    $OperatorIdentity = $MatchingOperatorAa.Identity

                }

                else {

                    $MatchingOperatorCq = $allCallQueues | Where-Object { $_.Identity -eq $Operator.Id }

                    $OperatorTypeFriendly = "[Call Queue"
                    $OperatorName = "$($MatchingOperatorCq.Name)]"
                    $OperatorIdentity = $MatchingOperatorCq.Identity

                }

                $AddOperatorToNestedVoiceApps = $true

            }

        }

        

    }

    else {

        $AddOperatorToNestedVoiceApps = $false

    }
    
}

function Find-AfterHours {
    param (
        [Parameter(Mandatory = $true)][String]$VoiceAppId

    )

    $aa = $allAutoAttendants | Where-Object {$_.Identity -eq $VoiceAppId}

    Write-Host "Reading call flow for: $($aa.Name)" -ForegroundColor Magenta
    Write-Host "##################################################" -ForegroundColor Magenta
    Write-Host "Voice App Id: $($aa.Identity)" -ForegroundColor Magenta

    if ($($aa.Name -ne (Optimize-DisplayName -String $aa.Name))) {

        Write-Warning -Message "The Auto Attendant '$($aa.Name)' contains special characters which will be removed in the diagram. New Name: '$(Optimize-DisplayName -String $aa.Name)'"

    }

    # Create ps object which has no business hours, needed to check if it matches an auto attendants after hours schedule
    $aaDefaultScheduleProperties = New-Object -TypeName psobject

    #$aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "ComplementEnabled" -Value $true
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "DisplayMondayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "DisplayTuesdayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "DisplayWednesdayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "DisplayThursdayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "DisplayFridayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "DisplaySaturdayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "DisplaySundayHours" -Value "00:00:00-1.00:00:00"

    # Convert to string for comparison
    $aaDefaultScheduleProperties = $aaDefaultScheduleProperties | Out-String
    
    # Get the current auto attendants after hours schedule and convert to string

    # Check if the auto attendant has business hours by comparing the ps object to the actual config of the current auto attendant
    # Additional check for auto attendants which somehow have no schedules at all
    if ($aa.Schedules.Type -contains "WeeklyRecurrence") {

        $aaAfterHoursScheduleId = ($aa.CallHandlingAssociations | Where-Object {$_.Type -eq "AfterHours"}).ScheduleId
        $aaAfterHoursScheduleProperties = ($aa.Schedules | Where-Object {$_.Id -eq $aaAfterHoursScheduleId}).WeeklyRecurrentSchedule

        . Read-BusinessHours

        if ($aaAfterHoursScheduleProperties.ComplementEnabled -eq $false) {

            Write-Warning -Message "Complement is disabled. This can only be set through PowerShell. Any time you change business hours in TAC, complement will be enabled again."

            $mdComplementNo = "Yes"
            $mdComplementYes = "No"

        }

        else {
                
            $mdComplementNo = "No"
            $mdComplementYes = "Yes"

        }
    
        # Check if the auto attendant has business hours by comparing the ps object to the actual config of the current auto attendant
        if ($aaDefaultScheduleProperties -eq ($aaEffectiveScheduleProperties | Out-String)) {
            
            $aaHasAfterHours = $false

        }

        else {

            $aaHasAfterHours = $true

            if ($CheckCallFlowRouting -eq $true) {

                Write-Host "Checking if Auto Attendant '$($aa.Name)' is currently in business hours or after hours schedule..." -ForegroundColor Magenta

                # Local time zone and date time
                $localTimeZone = (Get-TimeZone).Id

                if ($CheckCallFlowRoutingSpecificDateTime) {

                    $localDateTime = Get-Date $CheckCallFlowRoutingSpecificDateTime

                }

                else {
                        
                    $localDateTime = Get-Date

                }
                
                # Time zone configured on Auto Attendant
                $toTimeZone = $aa.TimeZoneId

                # Convert local time to time zone configured on Auto Attendant
                $convertedDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($localDateTime, $localTimeZone, $toTimeZone)

                $currentTimeAtAaTimeZone = $convertedDateTime.ToLongTimeString()
                $currentDayOfWeekAtAaTimeZone = $convertedDateTime.DayOfWeek

                $dayOfWeekAaScheduleToCheck = $aaAfterHoursScheduleProperties."$($currentDayOfWeekAtAaTimeZone)Hours"

                $openTimeRange = @()

                foreach ($businessHoursTimeRange in $dayOfWeekAaScheduleToCheck) {

                    # Check if end time is end of day
                    if ($businessHoursTimeRange.End.TotalHours -eq 24) {

                        $businessHoursTimeRangeString = "23:59:59"
                    }

                    else {

                        $businessHoursTimeRangeString = $businessHoursTimeRange.End.ToString()

                    }

                    if ($currentTimeAtAaTimeZone -ge ($businessHoursTimeRange.Start).ToString() -and $currentTimeAtAaTimeZone -le $businessHoursTimeRangeString) {

                        $openTimeRange += $businessHoursTimeRange

                    }


                    if ($aaAfterHoursScheduleProperties.ComplementEnabled -eq $true) {

                        Write-Host "Complement: Enabled" -ForegroundColor Yellow

                        $complementEnabled = $true

                    }

                    else {

                        Write-Host "Complement: Disabled" -ForegroundColor Yellow

                        $complementEnabled = $false

                    }

                }

                Write-Host "Local Time Zone: $localTimeZone" -ForegroundColor Yellow
                Write-Host "Local Date Time: $($localDateTime.ToString('dddd, yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Yellow
                Write-Host "Auto Attendant Time Zone: $toTimeZone" -ForegroundColor Yellow
                Write-Host "Time in Auto Attendant Time Zone: $($convertedDateTime.ToString('dddd, yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Yellow

                $businessHoursTimeRangesStarts = ($dayOfWeekAaScheduleToCheck.Start -join ",").Split(",")
                $businessHoursTimeRangesEnds = ($dayOfWeekAaScheduleToCheck.End -join ",").Split(",")

                $businessHoursTimeRangesString = ""

                $businessHoursTimeRangesIndex = 0

                foreach ($time in $businessHoursTimeRangesStarts) {

                    $businessHoursTimeRangesString += "$($businessHoursTimeRangesStarts[$businessHoursTimeRangesIndex]) - $($businessHoursTimeRangesEnds[$businessHoursTimeRangesIndex]), "

                    $businessHoursTimeRangesIndex ++

                }

                $businessHoursTimeRangesString = $businessHoursTimeRangesString.TrimEnd(", ")

                $businessHoursTimeRangesString = $businessHoursTimeRangesString.Replace("1.00:00:00","23:59:59")

                if ($VoiceAppFileName -eq $aa.Name) {

                    $topLevelAaInfo = "(initially queried Auto Attendant)"

                }

                else {

                    $topLevelAaInfo = "(nested Auto Attendant of initially queried Auto Attendant '$($VoiceAppFileName)')"

                }

                if ($openTimeRange -and $complementEnabled -eq $true) {

                    Write-Host "Business Hours for $($currentDayOfWeekAtAaTimeZone) in AA Time Zone: $businessHoursTimeRangesString" -ForegroundColor Yellow

                    Write-Host "Active Business Hours Time Range for $($currentDayOfWeekAtAaTimeZone): $($openTimeRange.Start) - $(($openTimeRange.End).ToString().Replace("1.00:00:00","23:59:59"))" -ForegroundColor Green
                    Write-Host "Auto Attendant: '$($aa.Name)' $topLevelAaInfo is currently in business hours schedule (open)" -ForegroundColor Green

                    Write-Host "Check Call Flow Diagram to see where calls are routed to during business hours..." -ForegroundColor Cyan

                    $autoAttendantIsOpen = $true

                }

                elseif ($openTimeRange -and $complementEnabled -eq $false) {

                    Write-Host "After Hours for $($currentDayOfWeekAtAaTimeZone) in AA Time Zone: $businessHoursTimeRangesString" -ForegroundColor Yellow

                    Write-Host "Active After Hours Time Range for $($currentDayOfWeekAtAaTimeZone): $($openTimeRange.Start) - $(($openTimeRange.End).ToString().Replace("1.00:00:00","23:59:59"))" -ForegroundColor Green
                    Write-Host "Auto Attendant: '$($aa.Name)' $topLevelAaInfo is currently in afther hours schedule (closed) because Complement is disabled and schedule is not inverted" -ForegroundColor Green

                    Write-Host "Check Call Flow Diagram to see where calls are routed to during after hours..." -ForegroundColor Cyan

                    $autoAttendantIsOpen = $false

                }

                elseif (!$openTimeRange -and $complementEnabled -eq $false) {

                    Write-Host "Business Hours for $($currentDayOfWeekAtAaTimeZone) in AA Time Zone: $businessHoursTimeRangesString" -ForegroundColor Yellow

                    Write-Host "Auto Attendant: '$($aa.Name)' $topLevelAaInfo is currently in business hours schedule (open) because Complement is disabled and schedule is not inverted" -ForegroundColor Red

                    Write-Host "Check Call Flow Diagram to see where calls are routed to during business hours..." -ForegroundColor Cyan

                    $autoAttendantIsOpen = $true

                }

                else {

                    Write-Host "Business Hours for $($currentDayOfWeekAtAaTimeZone) in AA Time Zone: $businessHoursTimeRangesString" -ForegroundColor Yellow

                    Write-Host "Auto Attendant: '$($aa.Name)' $topLevelAaInfo is currently in after hours schedule (closed)" -ForegroundColor Red

                    Write-Host "Check Call Flow Diagram to see where calls are routed to during after hours..." -ForegroundColor Cyan

                    $autoAttendantIsOpen = $false

                }

            }

        }

    }

    else {

        $aaHasAfterHours = $false

    }

    if ($aaHasAfterHours -eq $false -and $CheckCallFlowRouting -eq $true) {

        if ($VoiceAppFileName -eq $aa.Name) {

            $topLevelAaInfo = "(initially queried Auto Attendant)"

            $autoAttendantNestedType = "Top-Level"

        }

        else {

            $topLevelAaInfo = "(nested Auto Attendant of initially queried Auto Attendant '$($VoiceAppFileName)')"

            $autoAttendantNestedType = "Nested"

        }

        Write-Host "Checking if Auto Attendant '$($aa.Name)' is currently in business hours or after hours schedule..." -ForegroundColor Magenta

        Write-Host "Auto Attendant: '$($aa.Name)' $topLevelAaInfo has no business hours/after hours schedule configured. It's permanently in business hours schedule (open)" -ForegroundColor Green
     
        Write-Host "Check Call Flow Diagram to see where calls are permanently routed to (during business hours / default call flow)..." -ForegroundColor Cyan

        $autoAttendantIsOpen = $true

    }

    if ($CheckCallFlowRouting -eq $true -and $aaHasHolidays -eq $false) {

        if ($VoiceAppFileName -eq $aa.Name) {

            $topLevelAaInfo = "(initially queried Auto Attendant)"

            $autoAttendantNestedType = "Top-Level"

        }

        else {

            $topLevelAaInfo = "(nested Auto Attendant of initially queried Auto Attendant '$($VoiceAppFileName)')"

            $autoAttendantNestedType = "Nested"

        }

        Write-Host "Checking if Auto Attendant '$($aa.Name)' is currently in holiday schedule..." -ForegroundColor Magenta

        Write-Host "Auto Attendant: '$($aa.Name)' $topLevelAaInfo does not have any holidays configured." -ForegroundColor Green

        Write-Host "Check Call Flow Diagram to see where calls are routed to during normal operation (when it's not a holiday)..." -ForegroundColor Cyan

        # $autoAttendantIsOpen = $true

    }

    $allCallFlowRoutingChecks += [PSCustomObject]@{
        AutoAttendantName = "$($aa.Name)"
        CheckedDateTime = $localDateTime
        AutoAttendantIsOpen = $autoAttendantIsOpen
        AutoAttendantNestedType    = $autoAttendantNestedType
        AutoAttendantScheduleType = "BusinessHours"
        AutoAttendantScheduleName = "After Hours Call Flow"
        AutoAttendantScheduleRange        = "$($businessHoursTimeRangesString)"
        # AutoAttendantActiveBusinessHours = "$($openTimeRange.Start) - $(($openTimeRange.End).ToString().Replace("1.00:00:00","23:59:59"))"
        AutoAttendantTimeZone = "$($toTimeZone)"
        LocalTimeZone = "$($localTimeZone)"
        AutoAttendantId            = "$($aa.Identity)"
    }
   
}

function Get-AutoAttendantHolidaysAndAfterHours {
    param (
        [Parameter(Mandatory = $true)][String]$VoiceAppId
    )

    $aaObjectId = $aa.Identity

    $languageId = $aa.LanguageId

    $aa.Name = Optimize-DisplayName -String $aa.Name

    if ($aaHasHolidays -eq $true) {

        $holidaySubgraphName = "subgraphHolidays$($aa.Identity)[Holidays $($aa.Name)]"

        $allSubgraphs += "subgraphHolidays$($aa.Identity)"

        # The counter is here so that each element is unique in Mermaid
        $HolidayCounter = 1

        # Create empty mermaid subgraph for holidays
        $mdSubGraphHolidays =@"
subgraph $holidaySubgraphName
    direction LR
"@

        $aaHolidays = $aa.CallHandlingAssociations | Where-Object {$_.Type -match "Holiday" -and $_.Enabled -eq $true}

        $mdHolidayNestedCallFlowLinks = ""

        if ($CheckCallFlowRouting -eq $true) {

            Write-Host "Checking if Auto Attendant '$($aa.Name)' is currently in holiday schedule..." -ForegroundColor Magenta

            $holidayExceptionList = @()

        }

        foreach ($HolidayCallHandling in $aaHolidays) {

            $holidayCallFlow = $aa.CallFlows | Where-Object {$_.Id -eq $HolidayCallHandling.CallFlowId}
            $holidaySchedule = $aa.Schedules | Where-Object {$_.Id -eq $HolidayCallHandling.ScheduleId}

            if ($CheckCallFlowRouting -eq $true) {

                # Local time zone and date time
                $localTimeZone = (Get-TimeZone).Id

                if ($CheckCallFlowRoutingSpecificDateTime) {

                    $localDateTime = Get-Date $CheckCallFlowRoutingSpecificDateTime

                }

                else {
                        
                    $localDateTime = Get-Date

                }
                
                # Time zone configured on Auto Attendant
                $toTimeZone = $aa.TimeZoneId

                # Convert local time to time zone configured on Auto Attendant
                $convertedDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($localDateTime, $localTimeZone, $toTimeZone)

                foreach ($dateTimeRange in $holidaySchedule.FixedSchedule.DateTimeRanges) {
                    
                    if ($VoiceAppFileName -eq $aa.Name) {

                        $topLevelAaInfo = "(initially queried Auto Attendant)"

                        $autoAttendantNestedType = "Top-Level"

                    }

                    else {

                        $topLevelAaInfo = "(nested Auto Attendant of initially queried Auto Attendant '$($VoiceAppFileName)')"

                        $autoAttendantNestedType = "Nested"

                    }

                    if ($convertedDateTime -ge $dateTimeRange.Start -and $convertedDateTime -le $dateTimeRange.End) {

                        $holidayExceptionDetails = New-Object -TypeName psobject

                        $holidayExceptionDetails | Add-Member -MemberType NoteProperty -Name "HolidayCallFlowName" -Value $holidayCallFlow.Name
                        $holidayExceptionDetails | Add-Member -MemberType NoteProperty -Name "HolidayScheduleName" -Value $holidaySchedule.Name
                        $holidayExceptionDetails | Add-Member -MemberType NoteProperty -Name "HolidayScheduleStart" -Value $dateTimeRange.Start
                        $holidayExceptionDetails | Add-Member -MemberType NoteProperty -Name "HolidayScheduleEnd" -Value $dateTimeRange.End

                        $holidayExceptionList += $holidayExceptionDetails

                    }

                }

            }

            if (!$holidayCallFlow.Greetings) {

                $holidayGreeting = "Greeting <br> None"  

            }

            else {

                $holidayGreeting = "Greeting <br> $($holidayCallFlow.Greetings.ActiveType)"

                if ($($holidayCallFlow.Greetings.ActiveType) -eq "TextToSpeech" -and $ShowTTSGreetingText) {

                    $audioFileName = $null

                    $holidayTTSGreetingValueExport = $holidayCallFlow.Greetings.TextToSpeechPrompt
                    $holidayTTSGreetingValue = Optimize-DisplayName -String $holidayCallFlow.Greetings.TextToSpeechPrompt

                    if ($ExportTTSGreetings) {

                        $holidayTTSGreetingValueExport | Out-File "$FilePath\$($holidayCallFlow.Name)_$($aaObjectId)-$($HolidayCounter)_HolidayGreeting.txt"

                        $ttsGreetings += ("click elementAAHolidayGreeting$($aaObjectId)-$($HolidayCounter) " + '"' + "$FilePath\$($holidayCallFlow.Name)_$($aaObjectId)-$($HolidayCounter)_HolidayGreeting.txt" + '"')

                    }

                    if ($holidayTTSGreetingValue.Length -gt $truncateGreetings) {

                        $holidayTTSGreetingValue = $holidayTTSGreetingValue.Remove($holidayTTSGreetingValue.Length - ($holidayTTSGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                    
                    }

                    $holidayGreeting += " <br> ''$holidayTTSGreetingValue''"
                
                }

                elseif ($($holidayCallFlow.Greetings.ActiveType) -eq "AudioFile" -and $ShowAudioFileName) {

                    $holidayTTSGreetingValue = $null

                    # Audio File Greeting Name
                    $audioFileName = Optimize-DisplayName -String ($holidayCallFlow.Greetings.AudioFilePrompt.FileName)

                    if ($ExportAudioFiles) {

                        $content = Export-CsOnlineAudioFile -Identity $holidayCallFlow.Greetings.AudioFilePrompt.Id -ApplicationId OrgAutoAttendant
                        [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                        $audioFileNames += ("click elementAAHolidayGreeting$($aaObjectId)-$($HolidayCounter) " + '"' + "$FilePath\$audioFileName" + '"')
        
                    }


                    if ($audioFileName.Length -gt $truncateGreetings) {

                        $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                        $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"

                    }

                    $holidayGreeting += " <br> $audioFileName"


                }

            }

            $mermaidFriendlyHolidayName = Optimize-DisplayName -String $($holidayCallFlow.Name)

            $holidayAction = $holidayCallFlow.Menu.MenuOptions.Action

            if ($holidayCallFlow.Menu.Prompts.ActiveType) {

                $holidayCallHandlingName = Optimize-DisplayName -String $holidayCallFlow.Name

                $nodeElementHolidayAction = "elementAAHolidayAction$($aaObjectId)-$($HolidayCounter){IVR<br>$holidayCallHandlingName} $holidayVoicemailSystemGreeting"

                if ($ShowNestedHolidayIVRs) {

                    . .\Functions\Get-AutoAttendantHolidayCallFlow.ps1

                    . Get-AutoAttendantHolidayCallFlow

                    if ($holidayCallFlow.Menu.Prompts.ActiveType -and $holidayCallFlow.Menu.MenuOptions[0].DtmfResponse -ne "Automatic") {

                        # Display holiday name on link text
                        $mermaidCode += "$holidaySubgraphName -. Holiday: $mermaidFriendlyHolidayName .- elementAAHolidayIvr$($aaObjectId)-$($HolidayCounter)"
                        # No holiday name on link text
                        # $mermaidCode += "$holidaySubgraphName -.- elementAAHolidayIvr$($aaObjectId)-$($HolidayCounter)"

                    }

                    $mermaidCode += $mdAutoAttendantHolidayCallFlow

                    $allMermaidNodes += @("elementAAHolidayIvr$($aaObjectId)-$($HolidayCounter)")

                }

                $allMermaidNodes += @("elementAAHolidayAction$($aaObjectId)-$($HolidayCounter)")

            }

            else {

                # Check if holiday call handling is disconnect call
                if ($holidayAction -eq "DisconnectCall") {

                    $nodeElementHolidayAction = "elementAAHolidayAction$($aaObjectId)-$($HolidayCounter)(($holidayAction))"

                    $holidayVoicemailSystemGreeting = $null

                    $allMermaidNodes += "elementAAHolidayAction$($aaObjectId)-$($HolidayCounter)"

                }

                else {

                    $holidayActionTargetType = $holidayCallFlow.Menu.MenuOptions.CallTarget.Type

                    # Switch through different transfer call to target types
                    switch ($holidayActionTargetType) {
                        User { $holidayActionTargetTypeFriendly = "User" 
                        $holidayActionTargetName = Optimize-DisplayName -String (Get-MgUser -UserId $($holidayCallFlow.Menu.MenuOptions.CallTarget.Id)).DisplayName

                        $holidayVoicemailSystemGreeting = $null

                        if ($FindUserLinks -eq $true) {
            
                            . New-VoiceAppUserLinkProperties -userLinkUserId $($holidayCallFlow.Menu.MenuOptions.CallTarget.Id) -userLinkUserName $holidayActionTargetName -userLinkVoiceAppType "Auto Attendant" -userLinkVoiceAppActionType "HolidayActionTarget" -userLinkVoiceAppName $aa.Name -userLinkVoiceAppId $aa.Identity
                        
                        }

                        if ($ShowUserCallingSettings -eq $true -and $ShowNestedHolidayCallFlows -eq $true) {

                            # Display holiday name on link text
                            $mermaidCode += "$holidaySubgraphName -. Holiday: $mermaidFriendlyHolidayName .- $($holidayCallFlow.Menu.MenuOptions.CallTarget.Id)(User <br> $holidayActionTargetName)"
                            # No holiday name on link text
                            # $mermaidCode += "$holidaySubgraphName -.- $($holidayCallFlow.Menu.MenuOptions.CallTarget.Id)(User <br> $holidayActionTargetName)"
                            
                            if ($nestedVoiceApps -notcontains $($holidayCallFlow.Menu.MenuOptions.CallTarget.Id)) {

                                $nestedVoiceApps += $($holidayCallFlow.Menu.MenuOptions.CallTarget.Id)
    
                            }

                            $allMermaidNodes += $($holidayCallFlow.Menu.MenuOptions.CallTarget.Id)

                        }

                    }
                        SharedVoicemail { $holidayActionTargetTypeFriendly = "Shared Voicemail"
                        $holidayActionTargetName = Optimize-DisplayName -String (Get-MgGroup -GroupId $($holidayCallFlow.Menu.MenuOptions.CallTarget.Id)).DisplayName

                        if ($holidayCallFlow.Menu.MenuOptions.CallTarget.EnableSharedVoicemailSystemPromptSuppression -eq $false) {

                            $holidayVoicemailSystemGreeting = "--> elementAAHolidayVoicemailSystemGreeting$($aaObjectId)-$($HolidayCounter)>Greeting <br> MS System Message] "

                            $holidayVoicemailSystemGreetingValue = (. Get-MsSystemMessage)[-1]
                            $holidayVoicemailSystemGreetingValueExport = (. Get-MsSystemMessage)[0]

                            if ($ShowTTSGreetingText) {
            
                                if ($ExportTTSGreetings) {
            
                                    $holidayVoicemailSystemGreetingValueExport | Out-File "$FilePath\$($aaObjectId)_autoAttendantHolidayVoicemailMsSystemMessage.txt"
                    
                                    $ttsGreetings += ("click elementAAHolidayVoicemailSystemGreeting$($aaObjectId)-$($HolidayCounter) " + '"' + "$FilePath\$($aaObjectId)_autoAttendantHolidayVoicemailMsSystemMessage.txt" + '"')
                    
                                }
            
                                if ($holidayVoicemailSystemGreetingValue.Length -gt $truncateGreetings) {
            
                                    $holidayVoicemailSystemGreetingValue = $holidayVoicemailSystemGreetingValue.Remove($holidayVoicemailSystemGreetingValue.Length - ($holidayVoicemailSystemGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
            
                                }
            
                                $holidayVoicemailSystemGreeting = $holidayVoicemailSystemGreeting.Replace("] "," <br> ''$holidayVoicemailSystemGreetingValue''] ")
            
                            }        

                            $allMermaidNodes += "elementAAHolidayVoicemailSystemGreeting$($aaObjectId)-$($HolidayCounter)"

                        }

                        else {

                            $holidayVoicemailSystemGreeting = $null

                        }

                        if ($ShowSharedVoicemailGroupMembers -eq $true) {

                            . Get-SharedVoicemailGroupMembers -SharedVoicemailGroupId $($holidayCallFlow.Menu.MenuOptions.CallTarget.Id)

                            $holidayActionTargetName = "$holidayActionTargetName$mdSharedVoicemailGroupMembers"

                        }

                    }
                        ExternalPstn { $holidayActionTargetTypeFriendly = "External Number" 
                        $holidayActionTargetName =  ($holidayCallFlow.Menu.MenuOptions.CallTarget.Id).Replace("tel:","")

                        if ($ObfuscatePhoneNumbers -eq $true) {

                            $holidayActionTargetName = $holidayActionTargetName.Remove(($holidayActionTargetName.Length -4)) + "****"
        
                        }

                        $holidayVoicemailSystemGreeting = $null

                    }
                        # Check if the application endpoint is an auto attendant or a call queue
                        ApplicationEndpoint {

                            if ($ShowTransferCallToTargetType -eq $true) {

                                $holidayAction = "$holidayAction <br> Resource Account"

                            }
                            
                            $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $holidayCallFlow.Menu.MenuOptions.CallTarget.Id -and $_.ApplicationId -eq $applicationIdAa}

                            if ($matchingApplicationInstanceCheckAa) {

                                $MatchingAA = $allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $holidayCallFlow.Menu.MenuOptions.CallTarget.Id -or $_.Identity -eq $holidayCallFlow.Menu.MenuOptions.CallTarget.Id}

                                $holidayActionTargetTypeFriendly = "[Auto Attendant"
                                $holidayActionTargetName = (Optimize-DisplayName -String $($MatchingAA.Name)) + "]"

                                if ($ShowNestedHolidayCallFlows) {

                                    if ($nestedVoiceApps -notcontains $MatchingAA.Identity) {
        
                                        $nestedVoiceApps += $MatchingAA.Identity
            
                                    }

                                    $holidayActionTargetVoiceAppId = $MatchingAA.Identity
        
                                }

                            }

                            else {

                                $MatchingCQ = $allCallQueues | Where-Object {$_.ApplicationInstances -contains $holidayCallFlow.Menu.MenuOptions.CallTarget.Id -or $_.Identity -eq $holidayCallFlow.Menu.MenuOptions.CallTarget.Id}

                                $holidayActionTargetTypeFriendly = "[Call Queue"
                                $holidayActionTargetName = (Optimize-DisplayName -String $($MatchingCQ.Name)) + "]"

                                if ($ShowNestedHolidayCallFlows) {

                                    if ($nestedVoiceApps -notcontains $MatchingCQ.Identity) {
        
                                        $nestedVoiceApps += $MatchingCQ.Identity
            
                                    }

                                    $holidayActionTargetVoiceAppId = $MatchingCQ.Identity
        
                                }

                            }

                            $holidayVoicemailSystemGreeting = $null

                        }
                        ConfigurationEndpoint {

                            if ($ShowTransferCallToTargetType -eq $true) {

                                $holidayAction = "$holidayAction <br> Voice App"

                            }
                        
                            if ($allAutoAttendantIds -contains $holidayCallFlow.Menu.MenuOptions.CallTarget.Id) {

                                $MatchingAA = $allAutoAttendants | Where-Object { $_.Identity -eq $holidayCallFlow.Menu.MenuOptions.CallTarget.Id }

                                $holidayActionTargetTypeFriendly = "[Auto Attendant"
                                $holidayActionTargetName = (Optimize-DisplayName -String $($MatchingAA.Name)) + "]"

                                if ($ShowNestedHolidayCallFlows) {

                                    if ($nestedVoiceApps -notcontains $MatchingAA.Identity) {
        
                                        $nestedVoiceApps += $MatchingAA.Identity
            
                                    }

                                    $holidayActionTargetVoiceAppId = $MatchingAA.Identity
        
                                }

                            }

                            else {

                                $MatchingCQ = $allCallQueues | Where-Object { $_.Identity -eq $holidayCallFlow.Menu.MenuOptions.CallTarget.Id }

                                $holidayActionTargetTypeFriendly = "[Call Queue"
                                $holidayActionTargetName = (Optimize-DisplayName -String $($MatchingCQ.Name)) + "]"

                                if ($ShowNestedHolidayCallFlows) {

                                    if ($nestedVoiceApps -notcontains $MatchingCQ.Identity) {
        
                                        $nestedVoiceApps += $MatchingCQ.Identity
            
                                    }

                                    $holidayActionTargetVoiceAppId = $MatchingCQ.Identity
        
                                }

                            }

                            $holidayVoicemailSystemGreeting = $null

                        }

                    
                    }

                # Create mermaid code for the holiday action node based on the variables created in the switch statemenet
                $nodeElementHolidayAction = "elementAAHolidayAction$($aaObjectId)-$($HolidayCounter)($holidayAction) $holidayVoicemailSystemGreeting--> elementAAHolidayActionTargetType$($aaObjectId)-$($HolidayCounter)($holidayActionTargetTypeFriendly <br> $holidayActionTargetName)"

                $allMermaidNodes += @("elementAAHolidayAction$($aaObjectId)-$($HolidayCounter)","elementAAHolidayActionTargetType$($aaObjectId)-$($HolidayCounter)")

                }

            }

            # Create subgraph per holiday call handling inside the Holidays subgraph

            $holidayScheduleSorted = $holidaySchedule.FixedSchedule.DateTimeRanges | Sort-Object Start

            $holidayDates = ""
            $holidayScheduleCounter = 0
            
            foreach ($holidayDate in $holidayScheduleSorted) {
            
                $holidayDates += "$($holidayDate.Start.ToString($dateFormatString)) - $($holidayDate.End.ToString($dateFormatString))"
            
                $holidayScheduleCounter ++
            
                if ($holidayScheduleCounter -lt $holidayScheduleSorted.Count) {
            
                    $holidayDates += "<br>"
            
                }
            
            }

            $matchingHolidayScheduleName = Optimize-DisplayName -String (Get-CsOnlineSchedule -Id $holidaySchedule.Id).Name
            
            $nodeElementHolidayDetails =@"
            
            subgraph subgraphHolidayCallHandling$($HolidayCallHandling.CallFlowId)[Call Handling: $mermaidFriendlyHolidayName]
            subgraph subgraphHolidaySchedule$($HolidayCallHandling.CallFlowId)[Holiday Schedule: $matchingHolidayScheduleName]
            direction LR
            elementAAHoliday$($aaObjectId)-$($HolidayCounter)(Dates<br>$holidayDates) --> elementAAHolidayGreeting$($aaObjectId)-$($HolidayCounter)>$holidayGreeting] --> $nodeElementHolidayAction
                end
                end
"@
            
            $allMermaidNodes += @("elementAAHoliday$($aaObjectId)-$($HolidayCounter)","elementAAHolidayGreeting$($aaObjectId)-$($HolidayCounter)")


            if ($holidayGreeting -eq "Greeting <br> None") {

                $nodeElementHolidayDetails = $nodeElementHolidayDetails.Replace("elementAAHolidayGreeting$($aaObjectId)-$($HolidayCounter)>$holidayGreeting] --> ","")

            }

            if ($ShowNestedHolidayCallFlows -and !$holidayCallFlow.Menu.Prompts.ActiveType -and $holidayCallFlow.Menu.MenuOptions.CallTarget.Type -eq "ApplicationEndpoint" -or $holidayCallFlow.Menu.MenuOptions.CallTarget.Type -eq "ConfigurationEndpoint") {

                # Display holiday name on link text
                $mdHolidayNestedCallFlowLinks += "$holidaySubgraphName -. Holiday: $mermaidFriendlyHolidayName .- $holidayActionTargetVoiceAppId`n"
                # No holiday name on link name
                #$mdHolidayNestedCallFlowLinks += "$holidaySubgraphName -.- $holidayActionTargetVoiceAppId`n"

            }

            # Increase the counter by 1
            $HolidayCounter ++

            # Add holiday call handling subgraph to holiday subgraph
            $mdSubGraphHolidays += $nodeElementHolidayDetails

            $allSubgraphs += @("subgraphHolidayCallHandling$($HolidayCallHandling.CallFlowId)","subgraphHolidaySchedule$($HolidayCallHandling.CallFlowId)")

        } # End of for-each loop

        if ($CheckCallFlowRouting -eq $true -and $holidayExceptionList) {

            if ($holidayExceptionList.Count -gt 1) {

                $holidayExceptionListTimeSpans = @()

                foreach ($holidayException in $holidayExceptionList) {

                    $holidayExceptionListTimeSpans += ($holidayException.HolidayScheduleEnd - $holidayException.HolidayScheduleStart).TotalSeconds

                }

                $shortestTimeSpan = ($holidayExceptionListTimeSpans | Measure-Object -Minimum).Minimum
                $shortestTimeSpanIndex = $holidayExceptionListTimeSpans.IndexOf($shortestTimeSpan)

                $activeHolidayException = $holidayExceptionList[$shortestTimeSpanIndex]

                Write-Host "Auto Attendant: '$($aa.Name)' $topLevelAaInfo has multiple holiday schedules which match for $($localDateTime.ToString('dddd, yyyy-MM-dd HH:mm:ss')) (local time: $localTimeZone) / $($convertedDateTime.ToString('dddd, yyyy-MM-dd HH:mm:ss')) (Auto Attendant time: $toTimeZone)." -ForegroundColor Yellow
                Write-Host "In this case, the schedule which is most precise (smallest time range) will be active. All matching holidays:" -ForegroundColor Yellow
                Write-Host ($holidayExceptionList | Out-String) -ForegroundColor Yellow

            }

            else {

                $activeHolidayException = $holidayExceptionList

            }

            Write-Host "Local Time Zone: $localTimeZone" -ForegroundColor Yellow
            Write-Host "Local Date Time: $($localDateTime.ToString('dddd, yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Yellow
            Write-Host "Auto Attendant Time Zone: $toTimeZone" -ForegroundColor Yellow
            Write-Host "Time in Auto Attendant Time Zone: $($convertedDateTime.ToString('dddd, yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Yellow

            Write-Host "Holiday Schedule: $($activeHolidayException.HolidayScheduleStart.ToString('dddd, yyyy-MM-dd HH:mm:ss')) - $($activeHolidayException.HolidayScheduleEnd.ToString('dddd, yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Yellow
            Write-Host "Holiday Call Flow Name: $($activeHolidayException.HolidayCallFlowName)" -ForegroundColor Yellow
            Write-Host "Holiday Schedule Name: $($activeHolidayException.HolidayScheduleName)" -ForegroundColor Yellow
            Write-Host "Auto Attendant: '$($aa.Name)' $topLevelAaInfo is currently in holiday schedule (exception)" -ForegroundColor Red

            Write-Host "Check Call Flow Diagram to see where calls are routed to on Holiday '$($activeHolidayException.HolidayCallFlowName)'..." -ForegroundColor Cyan

            $allCallFlowRoutingChecks += [PSCustomObject]@{
                AutoAttendantName = "$($aa.Name)"
                CheckedDateTime            = $localDateTime
                AutoAttendantIsOpen        = $false
                AutoAttendantNestedType    = $autoAttendantNestedType
                AutoAttendantScheduleType = "Holiday"
                AutoAttendantScheduleName = "$($activeHolidayException.HolidayScheduleName)"
                AutoAttendantScheduleRange = "$($activeHolidayException.HolidayScheduleStart.ToString('dddd, yyyy-MM-dd HH:mm:ss')) - $($activeHolidayException.HolidayScheduleEnd.ToString('dddd, yyyy-MM-dd HH:mm:ss'))"
                AutoAttendantTimeZone = "$($toTimeZone)"
                LocalTimeZone = "$($localTimeZone)"
                AutoAttendantId            = "$($aa.Identity)"
            }

        }

        # Console output if there are no active holidays
        if ($CheckCallFlowRouting -eq $true -and !$holidayExceptionList) {

            Write-Host "Auto Attendant: '$($aa.Name)' $topLevelAaInfo is currently not in holiday schedule." -ForegroundColor Green

            Write-Host "Check Call Flow Diagram to see where calls are routed to during normal operation (when it's not a holiday)..." -ForegroundColor Cyan

        }

        # Create end for the holiday subgraph
        $mdSubGraphHolidaysEnd =@"

    end
"@

        if ($ShowNestedHolidayCallFlows -and $mdHolidayNestedCallFlowLinks -ne "") {

            $mdSubGraphHolidaysEnd += "`n$mdHolidayNestedCallFlowLinks"

        }
            
        # Add the end to the holiday subgraph mermaid code
        $mdSubGraphHolidays += $mdSubGraphHolidaysEnd

        # Mermaid node holiday check
        $nodeElementHolidayCheck = "elementHolidayCheck$($aaObjectId){During Holiday?}"
        $allMermaidNodes += "elementHolidayCheck$($aaObjectId)"
    } # End if aa has holidays

    # Check if auto attendant has after hours and holidays
    if ($aaHasAfterHours) {

        $aaTimeZone = $aa.TimeZoneId

        $aaBusinessHoursFriendly = $aaEffectiveScheduleProperties

        # Monday
        # Check if Monday has business hours which are open 24 Hours per day
        if ($aaBusinessHoursFriendly.DisplayMondayHours -eq "00:00:00-1.00:00:00") {
            $mondayHours = "Monday: Open 24 Hours"
        }
        # Check if Monday has business hours set different than 24 Hours open per day
        elseif ($aaBusinessHoursFriendly.DisplayMondayHours) {
            $mondayHours = "Monday: $($aaBusinessHoursFriendly.DisplayMondayHours)"

            if ($mondayHours -match ",") {

                $mondayHoursTimeRanges = $mondayHours.Split(",")

                $mondayHoursFirstTimeRange = "$($mondayHoursTimeRanges[0])"
                $MondayHoursFirstTimeRangeStart = $mondayHoursFirstTimeRange.Split("-")[0].Remove(($mondayHoursFirstTimeRange.Split("-")[0]).Length -3)
                $MondayHoursFirstTimeRangeEnd = $mondayHoursFirstTimeRange.Split("-")[1].Remove(($mondayHoursFirstTimeRange.Split("-")[1]).Length -3)
                $mondayHours = "$MondayHoursFirstTimeRangeStart - $MondayHoursFirstTimeRangeEnd"
    

                foreach ($TimeRange in $mondayHoursTimeRanges | Where-Object {$_ -notcontains $mondayHoursTimeRanges[0]} ) {

                    $MondayHoursStart = ($TimeRange.Split("-")[0].Remove(($TimeRange.Split("-")[0]).Length -3)).Replace("Monday: ","")
                    $MondayHoursEnd = ($TimeRange.Split("-")[1].Remove(($TimeRange.Split("-")[1]).Length -3))

                    $mondayHours += (", $MondayHoursStart - $MondayHoursEnd")
                }

            }

            else {

                $MondayHoursStart = $MondayHours.Split("-")[0].Remove(($MondayHours.Split("-")[0]).Length -3)
                $MondayHoursEnd = $MondayHours.Split("-")[1].Remove(($MondayHours.Split("-")[1]).Length -3)
                $MondayHours = "$MondayHoursStart - $MondayHoursEnd"    

            }

        }
        # Check if Monday has no business hours at all / is closed 24 Hours per day
        else {
            $mondayHours = "Monday: Closed"
        }

        # Tuesday
        if ($aaBusinessHoursFriendly.DisplayTuesdayHours -eq "00:00:00-1.00:00:00") {
            $TuesdayHours = "Tuesday: Open 24 Hours"
        }
        elseif ($aaBusinessHoursFriendly.DisplayTuesdayHours) {
            $TuesdayHours = "Tuesday: $($aaBusinessHoursFriendly.DisplayTuesdayHours)"

            if ($TuesdayHours -match ",") {

                $TuesdayHoursTimeRanges = $TuesdayHours.Split(",")

                $TuesdayHoursFirstTimeRange = "$($TuesdayHoursTimeRanges[0])"
                $TuesdayHoursFirstTimeRangeStart = $TuesdayHoursFirstTimeRange.Split("-")[0].Remove(($TuesdayHoursFirstTimeRange.Split("-")[0]).Length -3)
                $TuesdayHoursFirstTimeRangeEnd = $TuesdayHoursFirstTimeRange.Split("-")[1].Remove(($TuesdayHoursFirstTimeRange.Split("-")[1]).Length -3)
                $TuesdayHours = "$TuesdayHoursFirstTimeRangeStart - $TuesdayHoursFirstTimeRangeEnd"
    

                foreach ($TimeRange in $TuesdayHoursTimeRanges | Where-Object {$_ -notcontains $TuesdayHoursTimeRanges[0]} ) {

                    $TuesdayHoursStart = ($TimeRange.Split("-")[0].Remove(($TimeRange.Split("-")[0]).Length -3)).Replace("Tuesday: ","")
                    $TuesdayHoursEnd = ($TimeRange.Split("-")[1].Remove(($TimeRange.Split("-")[1]).Length -3))

                    $TuesdayHours += (", $TuesdayHoursStart - $TuesdayHoursEnd")
                }

            }

            else {

                $TuesdayHoursStart = $TuesdayHours.Split("-")[0].Remove(($TuesdayHours.Split("-")[0]).Length -3)
                $TuesdayHoursEnd = $TuesdayHours.Split("-")[1].Remove(($TuesdayHours.Split("-")[1]).Length -3)
                $TuesdayHours = "$TuesdayHoursStart - $TuesdayHoursEnd"    

            }

        } 
        else {
            $TuesdayHours = "Tuesday: Closed"
        }

        # Wednesday
        if ($aaBusinessHoursFriendly.DisplayWednesdayHours -eq "00:00:00-1.00:00:00") {
            $WednesdayHours = "Wednesday: Open 24 Hours"
        } 
        elseif ($aaBusinessHoursFriendly.DisplayWednesdayHours) {
            $WednesdayHours = "Wednesday: $($aaBusinessHoursFriendly.DisplayWednesdayHours)"

            if ($WednesdayHours -match ",") {

                $WednesdayHoursTimeRanges = $WednesdayHours.Split(",")

                $WednesdayHoursFirstTimeRange = "$($WednesdayHoursTimeRanges[0])"
                $WednesdayHoursFirstTimeRangeStart = $WednesdayHoursFirstTimeRange.Split("-")[0].Remove(($WednesdayHoursFirstTimeRange.Split("-")[0]).Length -3)
                $WednesdayHoursFirstTimeRangeEnd = $WednesdayHoursFirstTimeRange.Split("-")[1].Remove(($WednesdayHoursFirstTimeRange.Split("-")[1]).Length -3)
                $WednesdayHours = "$WednesdayHoursFirstTimeRangeStart - $WednesdayHoursFirstTimeRangeEnd"
    

                foreach ($TimeRange in $WednesdayHoursTimeRanges | Where-Object {$_ -notcontains $WednesdayHoursTimeRanges[0]} ) {

                    $WednesdayHoursStart = ($TimeRange.Split("-")[0].Remove(($TimeRange.Split("-")[0]).Length -3)).Replace("Wednesday: ","")
                    $WednesdayHoursEnd = ($TimeRange.Split("-")[1].Remove(($TimeRange.Split("-")[1]).Length -3))

                    $WednesdayHours += (", $WednesdayHoursStart - $WednesdayHoursEnd")
                }

            }

            else {

                $WednesdayHoursStart = $WednesdayHours.Split("-")[0].Remove(($WednesdayHours.Split("-")[0]).Length -3)
                $WednesdayHoursEnd = $WednesdayHours.Split("-")[1].Remove(($WednesdayHours.Split("-")[1]).Length -3)
                $WednesdayHours = "$WednesdayHoursStart - $WednesdayHoursEnd"    

            }

        }
        else {
            $WednesdayHours = "Wednesday: Closed"
        }

        # Thursday
        if ($aaBusinessHoursFriendly.DisplayThursdayHours -eq "00:00:00-1.00:00:00") {
            $ThursdayHours = "Thursday: Open 24 Hours"
        } 
        elseif ($aaBusinessHoursFriendly.DisplayThursdayHours) {
            $ThursdayHours = "Thursday: $($aaBusinessHoursFriendly.DisplayThursdayHours)"

            if ($ThursdayHours -match ",") {

                $ThursdayHoursTimeRanges = $ThursdayHours.Split(",")

                $ThursdayHoursFirstTimeRange = "$($ThursdayHoursTimeRanges[0])"
                $ThursdayHoursFirstTimeRangeStart = $ThursdayHoursFirstTimeRange.Split("-")[0].Remove(($ThursdayHoursFirstTimeRange.Split("-")[0]).Length -3)
                $ThursdayHoursFirstTimeRangeEnd = $ThursdayHoursFirstTimeRange.Split("-")[1].Remove(($ThursdayHoursFirstTimeRange.Split("-")[1]).Length -3)
                $ThursdayHours = "$ThursdayHoursFirstTimeRangeStart - $ThursdayHoursFirstTimeRangeEnd"
    

                foreach ($TimeRange in $ThursdayHoursTimeRanges | Where-Object {$_ -notcontains $ThursdayHoursTimeRanges[0]} ) {

                    $ThursdayHoursStart = ($TimeRange.Split("-")[0].Remove(($TimeRange.Split("-")[0]).Length -3)).Replace("Thursday: ","")
                    $ThursdayHoursEnd = ($TimeRange.Split("-")[1].Remove(($TimeRange.Split("-")[1]).Length -3))

                    $ThursdayHours += (", $ThursdayHoursStart - $ThursdayHoursEnd")
                }

            }

            else {

                $ThursdayHoursStart = $ThursdayHours.Split("-")[0].Remove(($ThursdayHours.Split("-")[0]).Length -3)
                $ThursdayHoursEnd = $ThursdayHours.Split("-")[1].Remove(($ThursdayHours.Split("-")[1]).Length -3)
                $ThursdayHours = "$ThursdayHoursStart - $ThursdayHoursEnd"    

            }

        }
        else {
            $ThursdayHours = "Thursday: Closed"
        }

        # Friday
        if ($aaBusinessHoursFriendly.DisplayFridayHours -eq "00:00:00-1.00:00:00") {
            $FridayHours = "Friday: Open 24 Hours"
        } 
        elseif ($aaBusinessHoursFriendly.DisplayFridayHours) {
            $FridayHours = "Friday: $($aaBusinessHoursFriendly.DisplayFridayHours)"

            if ($FridayHours -match ",") {

                $FridayHoursTimeRanges = $FridayHours.Split(",")

                $FridayHoursFirstTimeRange = "$($FridayHoursTimeRanges[0])"
                $FridayHoursFirstTimeRangeStart = $FridayHoursFirstTimeRange.Split("-")[0].Remove(($FridayHoursFirstTimeRange.Split("-")[0]).Length -3)
                $FridayHoursFirstTimeRangeEnd = $FridayHoursFirstTimeRange.Split("-")[1].Remove(($FridayHoursFirstTimeRange.Split("-")[1]).Length -3)
                $FridayHours = "$FridayHoursFirstTimeRangeStart - $FridayHoursFirstTimeRangeEnd"
    

                foreach ($TimeRange in $FridayHoursTimeRanges | Where-Object {$_ -notcontains $FridayHoursTimeRanges[0]} ) {

                    $FridayHoursStart = ($TimeRange.Split("-")[0].Remove(($TimeRange.Split("-")[0]).Length -3)).Replace("Friday: ","")
                    $FridayHoursEnd = ($TimeRange.Split("-")[1].Remove(($TimeRange.Split("-")[1]).Length -3))

                    $FridayHours += (", $FridayHoursStart - $FridayHoursEnd")
                }

            }

            else {

                $FridayHoursStart = $FridayHours.Split("-")[0].Remove(($FridayHours.Split("-")[0]).Length -3)
                $FridayHoursEnd = $FridayHours.Split("-")[1].Remove(($FridayHours.Split("-")[1]).Length -3)
                $FridayHours = "$FridayHoursStart - $FridayHoursEnd"    

            }

        }
        else {
            $FridayHours = "Friday: Closed"
        }

        # Saturday
        if ($aaBusinessHoursFriendly.DisplaySaturdayHours -eq "00:00:00-1.00:00:00") {
            $SaturdayHours = "Saturday: Open 24 Hours"
        } 

        elseif ($aaBusinessHoursFriendly.DisplaySaturdayHours) {
            $SaturdayHours = "Saturday: $($aaBusinessHoursFriendly.DisplaySaturdayHours)"

            if ($SaturdayHours -match ",") {

                $SaturdayHoursTimeRanges = $SaturdayHours.Split(",")

                $SaturdayHoursFirstTimeRange = "$($SaturdayHoursTimeRanges[0])"
                $SaturdayHoursFirstTimeRangeStart = $SaturdayHoursFirstTimeRange.Split("-")[0].Remove(($SaturdayHoursFirstTimeRange.Split("-")[0]).Length -3)
                $SaturdayHoursFirstTimeRangeEnd = $SaturdayHoursFirstTimeRange.Split("-")[1].Remove(($SaturdayHoursFirstTimeRange.Split("-")[1]).Length -3)
                $SaturdayHours = "$SaturdayHoursFirstTimeRangeStart - $SaturdayHoursFirstTimeRangeEnd"
    

                foreach ($TimeRange in $SaturdayHoursTimeRanges | Where-Object {$_ -notcontains $SaturdayHoursTimeRanges[0]} ) {

                    $SaturdayHoursStart = ($TimeRange.Split("-")[0].Remove(($TimeRange.Split("-")[0]).Length -3)).Replace("Saturday: ","")
                    $SaturdayHoursEnd = ($TimeRange.Split("-")[1].Remove(($TimeRange.Split("-")[1]).Length -3))

                    $SaturdayHours += (", $SaturdayHoursStart - $SaturdayHoursEnd")
                }

            }

            else {

                $SaturdayHoursStart = $SaturdayHours.Split("-")[0].Remove(($SaturdayHours.Split("-")[0]).Length -3)
                $SaturdayHoursEnd = $SaturdayHours.Split("-")[1].Remove(($SaturdayHours.Split("-")[1]).Length -3)
                $SaturdayHours = "$SaturdayHoursStart - $SaturdayHoursEnd"    

            }

        }

        else {
            $SaturdayHours = "Saturday: Closed"
        }

        # Sunday
        if ($aaBusinessHoursFriendly.DisplaySundayHours -eq "00:00:00-1.00:00:00") {
            $SundayHours = "Sunday: Open 24 Hours"
        }
        elseif ($aaBusinessHoursFriendly.DisplaySundayHours) {
            $SundayHours = "Sunday: $($aaBusinessHoursFriendly.DisplaySundayHours)"

            if ($SundayHours -match ",") {

                $SundayHoursTimeRanges = $SundayHours.Split(",")

                $SundayHoursFirstTimeRange = "$($SundayHoursTimeRanges[0])"
                $SundayHoursFirstTimeRangeStart = $SundayHoursFirstTimeRange.Split("-")[0].Remove(($SundayHoursFirstTimeRange.Split("-")[0]).Length -3)
                $SundayHoursFirstTimeRangeEnd = $SundayHoursFirstTimeRange.Split("-")[1].Remove(($SundayHoursFirstTimeRange.Split("-")[1]).Length -3)
                $SundayHours = "$SundayHoursFirstTimeRangeStart - $SundayHoursFirstTimeRangeEnd"
    

                foreach ($TimeRange in $SundayHoursTimeRanges | Where-Object {$_ -notcontains $SundayHoursTimeRanges[0]} ) {

                    $SundayHoursStart = ($TimeRange.Split("-")[0].Remove(($TimeRange.Split("-")[0]).Length -3)).Replace("Sunday: ","")
                    $SundayHoursEnd = ($TimeRange.Split("-")[1].Remove(($TimeRange.Split("-")[1]).Length -3))

                    $SundayHours += (", $SundayHoursStart - $SundayHoursEnd")
                }

            }

            else {

                $SundayHoursStart = $SundayHours.Split("-")[0].Remove(($SundayHours.Split("-")[0]).Length -3)
                $SundayHoursEnd = $SundayHours.Split("-")[1].Remove(($SundayHours.Split("-")[1]).Length -3)
                $SundayHours = "$SundayHoursStart - $SundayHoursEnd"    

            }

        }

        else {
            $SundayHours = "Sunday: Closed"
        }

        # Create the mermaid node for business hours check including the actual business hours
        $nodeElementAfterHoursCheck = "elementAfterHoursCheck$($aaObjectId){During Business Hours? <br> Time Zone: $aaTimeZone <br> $mondayHours <br> $tuesdayHours  <br> $wednesdayHours  <br> $thursdayHours <br> $fridayHours <br> $saturdayHours <br> $sundayHours}"

        $allMermaidNodes += "elementAfterHoursCheck$($aaObjectId)"

    } # End if aa has after hours

    $aa.Name = Optimize-DisplayName -String $aa.Name
    $nodeElementHolidayLink = "$($aa.Identity)([Auto Attendant <br> $($aa.Name)])"

    $allMermaidNodes += "$($aa.Identity)"

    if ($aaHasHolidays -eq $true) {

        if ($aaHasAfterHours) {

            $mdHolidayAndAfterHoursCheck =@"
$nodeElementHolidayLink --> $nodeElementHolidayCheck
$nodeElementHolidayCheck -->|Yes| $holidaySubgraphName
$nodeElementHolidayCheck -->|No| $nodeElementAfterHoursCheck
$nodeElementAfterHoursCheck -->|$mdComplementNo| $mdAutoAttendantAfterHoursCallFlow
$nodeElementAfterHoursCheck -->|$mdComplementYes| $mdAutoAttendantDefaultCallFlow

$mdSubGraphHolidays

"@
        }

        else {
            $mdHolidayAndAfterHoursCheck =@"
$nodeElementHolidayLink --> $nodeElementHolidayCheck
$nodeElementHolidayCheck -->|Yes| $holidaySubgraphName
$nodeElementHolidayCheck -->|No| $mdAutoAttendantDefaultCallFlow

$mdSubGraphHolidays

"@
        }

    }

    
    # Check if auto attendant has no Holidays but after hours
    else {
    
        if ($aaHasAfterHours -eq $true) {

            $mdHolidayAndAfterHoursCheck =@"
$nodeElementHolidayLink --> $nodeElementAfterHoursCheckCheck
$nodeElementAfterHoursCheck -->|$mdComplementNo| $mdAutoAttendantAfterHoursCallFlow
$nodeElementAfterHoursCheck -->|$mdComplementYes| $mdAutoAttendantDefaultCallFlow


"@      
        }

        # Check if auto attendant has no after hours and no holidays
        else {

            $mdHolidayAndAfterHoursCheck =@"
$nodeElementHolidayLink --> $mdAutoAttendantDefaultCallFlow

"@
        }

    
    }

    #Check if AA is not already present in mermaid code
    if ($mermaidCode -notcontains $mdHolidayAndAfterHoursCheck) {

        $mermaidCode += $mdHolidayAndAfterHoursCheck

    }

    if ($ShowAaAuthorizedUsers -eq $true -and $aa.AuthorizedUsers) {

        $mdAaAuthorizedUsers = "$($aa.Identity) -.- aaAuthorizedUsers$($aaObjectId)[(Authorized Users<br>"
    
        foreach ($aaAuthorizedUser in $aa.AuthorizedUsers.Guid) {
    
            $aaAuthorizedCsOnlineUser = Get-CsOnlineUser -Identity $aaAuthorizedUser
    
            if (!$aaAuthorizedCsOnlineUser.TeamsVoiceApplicationsPolicy.Name) {
    
                Write-Warning -Message "User $($aaAuthorizedCsOnlineUser.DisplayName) is an authorized user of CQ $($aa.Name) but doesn't have a Voice Application Policy assigned."
    
                $mdAaAuthorizedUserVoiceApplicationPolicy = ", Assigned Policy: None"
    
            }
    
            else {
    
                # $aaAuthorizedUserVoiceApplicationPolicy = Get-CsTeamsVoiceApplicationsPolicy -Identity $aaAuthorizedCsOnlineUser.TeamsVoiceApplicationsPolicy.Name
    
                $mdAaAuthorizedUserVoiceApplicationPolicy = ", Assigned Policy: $($aaAuthorizedCsOnlineUser.TeamsVoiceApplicationsPolicy.Name)"
    
            }
    
            $mdAaAuthorizedUsers += ($aaAuthorizedCsOnlineUser.DisplayName) + $mdAaAuthorizedUserVoiceApplicationPolicy + "<br>"
    
        }
    
        $mdAaAuthorizedUsers = $mdAaAuthorizedUsers.Remove(($mdAaAuthorizedUsers.Length -4),4)
        $mdAaAuthorizedUsers += ")]"
    
        $allMermaidNodes += "aaAuthorizedUsers$($aaObjectId)"

        if ($mermaidCode -notcontains $mdAaAuthorizedUsers) {

            $mermaidCode += $mdAaAuthorizedUsers
    
        }
    
    }
    
}

function Get-AutoAttendantDefaultCallFlow {
    param (
        [Parameter(Mandatory = $false)][String]$VoiceAppId
    )

    $aaDefaultCallFlowAaObjectId = $aa.Identity

    $languageId = $aa.LanguageId

    # Get the current auto attendants default call flow and default call flow action
    $defaultCallFlow = $aa.DefaultCallFlow
    $defaultCallFlowAction = $aa.DefaultCallFlow.Menu.MenuOptions.Action

    . Get-AutoAttendantDirectorySearchConfig -CallFlowType "defaultCallFlow"

    # Get the current auto attentans default call flow greeting
    if (!$defaultCallFlow.Greetings.ActiveType){
        $defaultCallFlowGreeting = "Greeting <br> None"
    }

    else {

        $defaultCallFlowGreeting = "Greeting <br> $($defaultCallFlow.Greetings.ActiveType)"

        if ($($defaultCallFlow.Greetings.ActiveType) -eq "TextToSpeech" -and $ShowTTSGreetingText) {

            $audioFileName = $null

            $defaultTTSGreetingValueExport = $defaultCallFlow.Greetings.TextToSpeechPrompt
            $defaultTTSGreetingValue = Optimize-DisplayName -String $defaultCallFlow.Greetings.TextToSpeechPrompt

            if ($ExportTTSGreetings) {

                $defaultTTSGreetingValueExport | Out-File "$FilePath\$($aaDefaultCallFlowAaObjectId)_defaultCallFlowGreeting.txt"

                $ttsGreetings += ("click defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId) " + '"' + "$FilePath\$($aaDefaultCallFlowAaObjectId)_defaultCallFlowGreeting.txt" + '"')

            }

            if ($defaultTTSGreetingValue.Length -gt $truncateGreetings) {

                $defaultTTSGreetingValue = $defaultTTSGreetingValue.Remove($defaultTTSGreetingValue.Length - ($defaultTTSGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
            
            }

            $defaultCallFlowGreeting += " <br> ''$defaultTTSGreetingValue''"
        
        }

        elseif ($($defaultCallFlow.Greetings.ActiveType) -eq "AudioFile" -and $ShowAudioFileName) {

            $defaultTTSGreetingValue = $null

            # Audio File Greeting Name
            $audioFileName = Optimize-DisplayName -String ($defaultCallFlow.Greetings.AudioFilePrompt.FileName)

            if ($ExportAudioFiles) {

                $content = Export-CsOnlineAudioFile -Identity $defaultCallFlow.Greetings.AudioFilePrompt.Id -ApplicationId OrgAutoAttendant
                [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)

                $audioFileNames += ("click defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId) " + '"' + "$FilePath\$audioFileName" + '"')

            }


            if ($audioFileName.Length -gt $truncateGreetings) {

                $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"

            }

            $defaultCallFlowGreeting += " <br> $audioFileName"


        }

    }

    # Check if default callflow action is disconnect call
    if ($defaultCallFlowAction -eq "DisconnectCall") {

        if ($CombineDisconnectCallNodes -eq $true) {

            $mdAutoAttendantDefaultCallFlow = "defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowGreeting] --> disconnectCall(($defaultCallFlowAction))`n"

            $allMermaidNodes += @("defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)","disconnectCall")

        }

        else {

            $mdAutoAttendantDefaultCallFlow = "defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowGreeting] --> defaultCallFlow$($aaDefaultCallFlowAaObjectId)(($defaultCallFlowAction))`n"

            $allMermaidNodes += @("defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)","defaultCallFlow$($aaDefaultCallFlowAaObjectId)")

        }

    }

    # Check if the default callflow action is transfer call to target
    else {

        $defaultCallFlowMenuOptions = $aa.DefaultCallFlow.Menu.MenuOptions

        if ($defaultCallFlowMenuOptions.Count -lt 2 -and !$defaultCallFlow.Menu.Prompts.ActiveType) {

            $mdDefaultCallFlowGreeting = "defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowGreeting] --> "

            $allMermaidNodes += @("defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)")

            $defaultCallFlowMenuOptionsKeyPress = $null

            $mdAutoAttendantDefaultCallFlowMenuOptions = $null

        }

        # Auto Attendant has multiple options / voice menu
        else {

            $defaultCallFlowMenuOptionsGreeting = "IVR Greeting <br> $($defaultCallFlow.Menu.Prompts.ActiveType)"

            if ($($defaultCallFlow.Menu.Prompts.ActiveType) -eq "TextToSpeech" -and $ShowTTSGreetingText) {

                $audioFileName = $null
    
                $defaultCallFlowMenuOptionsTTSGreetingValueExport = $defaultCallFlow.Menu.Prompts.TextToSpeechPrompt
                $defaultCallFlowMenuOptionsTTSGreetingValue = Optimize-DisplayName -String $defaultCallFlow.Menu.Prompts.TextToSpeechPrompt

                if ($ExportTTSGreetings) {

                    $defaultCallFlowMenuOptionsTTSGreetingValueExport | Out-File "$FilePath\$($aaDefaultCallFlowAaObjectId)_defaultCallFlowMenuOptionsGreeting.txt"
    
                    $ttsGreetings += ("click defaultCallFlowMenuOptionsGreeting$($aaDefaultCallFlowAaObjectId) " + '"' + "$FilePath\$($aaDefaultCallFlowAaObjectId)_defaultCallFlowMenuOptionsGreeting.txt" + '"')
    
                }    
    
                if ($defaultCallFlowMenuOptionsTTSGreetingValue.Length -gt $truncateGreetings) {
    
                    $defaultCallFlowMenuOptionsTTSGreetingValue = $defaultCallFlowMenuOptionsTTSGreetingValue.Remove($defaultCallFlowMenuOptionsTTSGreetingValue.Length - ($defaultCallFlowMenuOptionsTTSGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                
                }
    
                $defaultCallFlowMenuOptionsGreeting += " <br> ''$defaultCallFlowMenuOptionsTTSGreetingValue''"
            
            }
    
            elseif ($($defaultCallFlow.Menu.Prompts.ActiveType) -eq "AudioFile" -and $ShowAudioFileName) {
    
                $defaultCallFlowMenuOptionsTTSGreetingValue = $null
    
                # Audio File Greeting Name
                $audioFileName = Optimize-DisplayName -String ($defaultCallFlow.Menu.Prompts.AudioFilePrompt.FileName)

                if ($ExportAudioFiles) {

                    $content = Export-CsOnlineAudioFile -Identity $defaultCallFlow.Menu.Prompts.AudioFilePrompt.Id -ApplicationId OrgAutoAttendant
                    [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
    
                    $audioFileNames += ("click defaultCallFlowMenuOptionsGreeting$($aaDefaultCallFlowAaObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
    
                }

    
                if ($audioFileName.Length -gt $truncateGreetings) {
    
                    $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                    $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
    
                }
    
                $defaultCallFlowMenuOptionsGreeting += " <br> $audioFileName"
    
    
            }

            if ($defaultCallFlow.ForceListenMenuEnabled -eq $true) {

                $defaultCallFlowForceListen = "<br> Force Listen: True"

            }

            else {

                $defaultCallFlowForceListen = "<br> Force Listen: False"

            }

            if ($aaIsVoiceResponseEnabled) {

                $defaultCallFlowVoiceResponse = " or <br> Voice Response <br> Language: $($aa.LanguageId)<br>Voice Style: $($aa.VoiceId)$defaultCallFlowForceListen"

            }

            else {

                $defaultCallFlowVoiceResponse = "$defaultCallFlowForceListen"

            }

            $defaultCallFlowMenuOptionsKeyPress = @"

defaultCallFlowMenuOptions$($aaDefaultCallFlowAaObjectId){Key Press$defaultCallFlowVoiceResponse}
"@

            $mdDefaultCallFlowGreeting =@"
defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowGreeting] --> defaultCallFlowMenuOptionsGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowMenuOptionsGreeting] --> $defaultCallFlowMenuOptionsKeyPress

"@

            $mdAutoAttendantDefaultCallFlowMenuOptions =@"

"@

            $allMermaidNodes += @("defaultCallFlowMenuOptions$($aaDefaultCallFlowAaObjectId)","defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)","defaultCallFlowMenuOptionsGreeting$($aaDefaultCallFlowAaObjectId)")

        }

        $defaultCallFlowSharedVoicemailCounter = 1
        $defaultCallFlowUserOptionCounter = 1
        $defaultCallFlowPSTNOptionCounter = 1

        foreach ($MenuOption in $defaultCallFlowMenuOptions) {

            if ($defaultCallFlowMenuOptions.Count -lt 2 -and !$defaultCallFlow.Menu.Prompts.ActiveType) {

                $mdDtmfLink = $null
                $DtmfKey = $null
                $voiceResponse = $null

            }

            else {

                if ($aaIsVoiceResponseEnabled) {

                    if ($MenuOption.VoiceResponses) {

                        $voiceResponse = "/ Voice Response: ''$($MenuOption.VoiceResponses)''"

                    }

                    else {

                        $voiceResponse = "/ No Voice Response Configured"

                    }

                }

                else {

                    $voiceResponse = $null

                }

                [String]$DtmfKey = ($MenuOption.DtmfResponse)

                $DtmfKey = $DtmfKey.Replace("Tone","")

                $mdDtmfLink = "$defaultCallFlowMenuOptionsKeyPress --> |$DtmfKey $voiceResponse|"

            }

            # Get transfer target type
            $defaultCallFlowTargetType = $MenuOption.CallTarget.Type
            $defaultCallFlowAction = $MenuOption.Action

            if ($defaultCallFlowAction -eq "TransferCallToOperator") {

                switch ($aa.Operator.Type) {
                    ApplicationEndpoint {

                        if ($ShowTransferCallToTargetType -eq $true) {

                            $defaultCallFlowAction = "$defaultCallFlowAction <br> Resource Account"
                          
                        }

                    }
                    ConfigurationEndpoint {

                        if ($ShowTransferCallToTargetType -eq $true) {

                            $defaultCallFlowAction = "$defaultCallFlowAction <br> Voice App"       
                          
                        }

                    }
                    Default {}
                }

                if ($aaIsVoiceResponseEnabled) {

                    $defaultCallFlowOperatorVoiceResponse = "/ Voice Response: ''Operator''"

                    $mdDtmfLink = $mdDtmfLink.Replace($voiceResponse,$defaultCallFlowOperatorVoiceResponse)

                }

                else {

                    $defaultCallFlowOperatorVoiceResponse = $null

                }

                if ($ShowTTSGreetingText) {

                    $defaultCallFlowTransferGreetingOperatorValue = (. Get-IvrTransferMessage)[2]
                    $defaultCallFlowTransferGreetingOperatorValueExport = (. Get-IvrTransferMessage)[3]

                    if ($ExportTTSGreetings) {

                        $defaultCallFlowTransferGreetingOperatorValueExport | Out-File "$FilePath\$($aadefaultCallFlowAaObjectId)_autoAttendantdefaultCallFlowTransferGreetingOperator.txt"
                        $ttsGreetings += ("click defaultCallFlowTransferGreetingOperator$($aadefaultCallFlowAaObjectId) " + '"' + "$FilePath\$($aadefaultCallFlowAaObjectId)_autoAttendantdefaultCallFlowTransferGreetingOperator.txt" + '"')

                    }

                    if ($defaultCallFlowTransferGreetingOperatorValue.Length -gt $truncateGreetings) {

                        $defaultCallFlowTransferGreetingOperatorValue = $defaultCallFlowTransferGreetingOperatorValue.Remove($defaultCallFlowTransferGreetingOperatorValue.Length - ($defaultCallFlowTransferGreetingOperatorValue.Length -$truncateGreetings)).TrimEnd() + "..."
            
                    }

                    $defaultCallFlowTransferGreetingOperatorValue = "defaultCallFlowTransferGreetingOperator$($aaDefaultCallFlowAaObjectId)>Greeting<br>Transfer Message<br>''$defaultCallFlowTransferGreetingOperatorValue''] -->"
    
                }

                else {

                    $defaultCallFlowTransferGreetingOperatorValueExport = $null
                    $defaultCallFlowTransferGreetingOperatorValue = "defaultCallFlowTransferGreetingOperator$($aaDefaultCallFlowAaObjectId)>Greeting<br>Transfer Message] -->"

                }

                $allMermaidNodes += "defaultCallFlowTransferGreetingOperator$($aaDefaultCallFlowAaObjectId)"

                $mdAutoAttendantdefaultCallFlow = "$mdDtmfLink $defaultCallFlowTransferGreetingOperatorValue defaultCallFlow$($aadefaultCallFlowAaObjectId)$DtmfKey($defaultCallFlowAction) --> $($OperatorIdentity)($OperatorTypeFriendly <br> $OperatorName)`n"

                $allMermaidNodes += @("defaultCallFlow$($aadefaultCallFlowAaObjectId)$DtmfKey","$($OperatorIdentity)")

                $defaultCallFlowVoicemailSystemGreeting = $null

                if ($nestedVoiceApps -notcontains $OperatorIdentity -and $AddOperatorToNestedVoiceApps -eq $true) {

                    $nestedVoiceApps += $OperatorIdentity

                }

                if ($FindUserLinks -eq $true) {
         
                    . New-VoiceAppUserLinkProperties -userLinkUserId $OperatorIdentity -userLinkUserName $OperatorName -userLinkVoiceAppType "Auto Attendant" -userLinkVoiceAppActionType "TransferCallToOperator (DTMF Option: $DtmfKey)" -userLinkVoiceAppName $aa.Name -userLinkVoiceAppId $aa.Identity
                
                }

            }

            elseif ($defaultCallFlowAction -eq "Announcement") {
                
                $voiceMenuOptionAnnouncementType = $MenuOption.Prompt.ActiveType

                $defaultCallFlowMenuOptionsAnnouncement = "$voiceMenuOptionAnnouncementType"

                if ($voiceMenuOptionAnnouncementType -eq "TextToSpeech" -and $ShowTTSGreetingText) {
    
                    $audioFileName = $null
        
                    $defaultCallFlowMenuOptionsTTSAnnouncementValueExport = $MenuOption.Prompt.TextToSpeechPrompt
                    $defaultCallFlowMenuOptionsTTSAnnouncementValue = Optimize-DisplayName -String $MenuOption.Prompt.TextToSpeechPrompt

                    if ($ExportTTSGreetings) {

                        $defaultCallFlowMenuOptionsTTSAnnouncementValueExport | Out-File "$FilePath\$($aaDefaultCallFlowAaObjectId)_defaultCallFlowMenuOptionsAnnouncement$DtmfKey.txt"
        
                        $ttsGreetings += ("click defaultCallFlow$($aadefaultCallFlowAaObjectId)$DtmfKey " + '"' + "$FilePath\$($aaDefaultCallFlowAaObjectId)_defaultCallFlowMenuOptionsAnnouncement$DtmfKey.txt" + '"')
        
                    }    
    
                    if ($defaultCallFlowMenuOptionsTTSAnnouncementValue.Length -gt $truncateGreetings) {
        
                        $defaultCallFlowMenuOptionsTTSAnnouncementValue = $defaultCallFlowMenuOptionsTTSAnnouncementValue.Remove($defaultCallFlowMenuOptionsTTSAnnouncementValue.Length - ($defaultCallFlowMenuOptionsTTSAnnouncementValue.Length -$truncateGreetings)).TrimEnd() + "..."
                    
                    }
        
                    $defaultCallFlowMenuOptionsAnnouncement += " <br> ''$defaultCallFlowMenuOptionsTTSAnnouncementValue''"
                
                }
        
                elseif ($voiceMenuOptionAnnouncementType -eq "AudioFile" -and $ShowAudioFileName) {
        
                    $defaultCallFlowMenuOptionsTTSAnnouncementValue = $null
        
                    # Audio File Announcement Name
                    $audioFileName = Optimize-DisplayName -String ($MenuOption.Prompt.AudioFilePrompt.FileName)

                    if ($ExportAudioFiles) {

                        $content = Export-CsOnlineAudioFile -Identity $MenuOption.Prompt.AudioFilePrompt.Id -ApplicationId OrgAutoAttendant
                        [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                        $audioFileNames += ("click defaultCallFlow$($aadefaultCallFlowAaObjectId)$DtmfKey " + '"' + "$FilePath\$audioFileName" + '"')
        
                    }
    
        
                    if ($audioFileName.Length -gt $truncateGreetings) {
        
                        $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                        $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                    }
        
                    $defaultCallFlowMenuOptionsAnnouncement += " <br> $audioFileName"
        
        
                }

                $mdAutoAttendantdefaultCallFlow = "$mdDtmfLink defaultCallFlow$($aadefaultCallFlowAaObjectId)$DtmfKey>$defaultCallFlowAction <br> $defaultCallFlowMenuOptionsAnnouncement] ---> defaultCallFlowMenuOptionsGreeting$($aaDefaultCallFlowAaObjectId)`n"

                $allMermaidNodes += @("defaultCallFlow$($aadefaultCallFlowAaObjectId)$DtmfKey","defaultCallFlowMenuOptionsGreeting$($aaDefaultCallFlowAaObjectId)")

                $defaultCallFlowVoicemailSystemGreeting = $null

            }

            else {

                # Switch through transfer target type and set variables accordingly
                switch ($defaultCallFlowTargetType) {
                    User { 
                        $defaultCallFlowTargetTypeFriendly = "User"
                        $defaultCallFlowTargetUser = (Get-MgUser -UserId $($MenuOption.CallTarget.Id))
                        $defaultCallFlowTargetName = Optimize-DisplayName -String $defaultCallFlowTargetUser.DisplayName
                        $defaultCallFlowTargetIdentity = $defaultCallFlowTargetUser.Id

                        if ($FindUserLinks -eq $true) {

                            if ($DtmfKey) {

                                $DtmfOption = " (DTMF Option: $DtmfKey)"

                            }

                            else {

                                $DtmfOption =$null

                            }
         
                            . New-VoiceAppUserLinkProperties -userLinkUserId $($MenuOption.CallTarget.Id) -userLinkUserName $defaultCallFlowTargetUser.DisplayName -userLinkVoiceAppType "Auto Attendant" -userLinkVoiceAppActionType "DefaultCallFlowTransferCallToTarget$DtmfOption" -userLinkVoiceAppName $aa.Name -userLinkVoiceAppId $aa.Identity
                        
                        }                

                        if ($nestedVoiceApps -notcontains $defaultCallFlowTargetUser.Id) {

                            $nestedVoiceApps += $defaultCallFlowTargetUser.Id

                        }

                        $defaultCallFlowVoicemailSystemGreeting = $null

                        if ($MenuOption.DtmfResponse -ne "Automatic") {

                            if ($ShowTTSGreetingText) {

                                $defaultCallFlowTransferGreetingValue = (. Get-IvrTransferMessage)[0]
                                $defaultCallFlowTransferGreetingValueExport = (. Get-IvrTransferMessage)[1]
            
                                if ($ExportTTSGreetings) {
            
                                    $defaultCallFlowTransferGreetingValueExport | Out-File "$FilePath\$($aadefaultCallFlowAaObjectId)_autoAttendantdefaultCallFlowTransferUserGreeting.txt"
                                    $ttsGreetings += ("click defaultCallFlowTransferUserGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowUserOptionCounter " + '"' + "$FilePath\$($aadefaultCallFlowAaObjectId)_autoAttendantdefaultCallFlowTransferUserGreeting.txt" + '"')
            
                                }
            
                                if ($defaultCallFlowTransferGreetingValue.Length -gt $truncateGreetings) {
            
                                    $defaultCallFlowTransferGreetingValue = $defaultCallFlowTransferGreetingValue.Remove($defaultCallFlowTransferGreetingValue.Length - ($defaultCallFlowTransferGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                        
                                }
            
                                $defaultCallFlowTransferGreetingValue = "defaultCallFlowTransferUserGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowUserOptionCounter>Greeting<br>Transfer Message<br>''$defaultCallFlowTransferGreetingValue''] -->"
                
                            }
        
                            else {
            
                                $defaultCallFlowTransferGreetingValueExport = $null
                                $defaultCallFlowTransferGreetingValue = "defaultCallFlowTransferUserGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowUserOptionCounter>Greeting<br>Transfer Message] -->"
            
                            }

                        }

                        else {

                            $defaultCallFlowTransferGreetingValueExport = $null
                            $defaultCallFlowTransferGreetingValue = $null

                        }
        
                        $allMermaidNodes += "defaultCallFlowTransferUserGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowUserOptionCounter"

                        $defaultCallFlowUserOptionCounter ++
        

                    }
                    ExternalPstn { 
                        $defaultCallFlowTargetTypeFriendly = "External Number"
                        $defaultCallFlowTargetName = ($MenuOption.CallTarget.Id).Replace("tel:","")
                        $defaultCallFlowTargetIdentity = $defaultCallFlowTargetName

                        if ($ObfuscatePhoneNumbers -eq $true) {

                            $defaultCallFlowTargetName = $defaultCallFlowTargetName.Remove(($defaultCallFlowTargetName.Length -4)) + "****"
        
                        }

                        $defaultCallFlowVoicemailSystemGreeting = $null

                        if ($MenuOption.DtmfResponse -ne "Automatic") {

                            if ($ShowTTSGreetingText) {

                                $defaultCallFlowTransferGreetingValue = (. Get-IvrTransferMessage)[0]
                                $defaultCallFlowTransferGreetingValueExport = (. Get-IvrTransferMessage)[1]
            
                                if ($ExportTTSGreetings) {
            
                                    $defaultCallFlowTransferGreetingValueExport | Out-File "$FilePath\$($aadefaultCallFlowAaObjectId)_autoAttendantdefaultCallFlowTransferPSTNGreeting.txt"
                                    $ttsGreetings += ("click defaultCallFlowTransferPSTNGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowPSTNOptionCounter " + '"' + "$FilePath\$($aadefaultCallFlowAaObjectId)_autoAttendantdefaultCallFlowTransferPSTNGreeting.txt" + '"')
            
                                }
            
                                if ($defaultCallFlowTransferGreetingValue.Length -gt $truncateGreetings) {
            
                                    $defaultCallFlowTransferGreetingValue = $defaultCallFlowTransferGreetingValue.Remove($defaultCallFlowTransferGreetingValue.Length - ($defaultCallFlowTransferGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                        
                                }
            
                                $defaultCallFlowTransferGreetingValue = "defaultCallFlowTransferPSTNGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowPSTNOptionCounter>Greeting<br>Transfer Message<br>''$defaultCallFlowTransferGreetingValue''] -->"
                
                            }
            
                            else {
            
                                $defaultCallFlowTransferGreetingValueExport = $null
                                $defaultCallFlowTransferGreetingValue = "defaultCallFlowTransferPSTNGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowPSTNOptionCounter>Greeting<br>Transfer Message] -->"
            
                            }

                        }

                        else {
            
                            $defaultCallFlowTransferGreetingValueExport = $null
                            $defaultCallFlowTransferGreetingValue = $null
        
                        }
        
                        $allMermaidNodes += "defaultCallFlowTransferPSTNGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowPSTNOptionCounter"

                        $defaultCallFlowPSTNOptionCounter ++

                    }
                    ApplicationEndpoint {

                        # Check if application endpoint is auto attendant or call queue

                        if ($ShowTransferCallToTargetType -eq $true) {

                            $defaultCallFlowAction = "$defaultCallFlowAction <br> Resource Account"       
                          
                        }

                        $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $MenuOption.CallTarget.Id -and $_.ApplicationId -eq $applicationIdAa}

                        if ($matchingApplicationInstanceCheckAa) {

                            $MatchingAaDefaultCallFlowAa = $allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id -or $_.Identity -eq $MenuOption.CallTarget.Id}

                            $defaultCallFlowTargetTypeFriendly = "[Auto Attendant"
                            $defaultCallFlowTargetName = (Optimize-DisplayName -String $($MatchingAaDefaultCallFlowAa.Name)) + "]"

                        }

                        else {

                            $MatchingCqAaDefaultCallFlow = $allCallQueues | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id -or $_.Identity -eq $MenuOption.CallTarget.Id}

                            $defaultCallFlowTargetTypeFriendly = "[Call Queue"
                            $defaultCallFlowTargetName = (Optimize-DisplayName -String $($MatchingCqAaDefaultCallFlow.Name)) + "]"

                        }

                        $defaultCallFlowVoicemailSystemGreeting = $null

                    }
                    ConfigurationEndpoint {

                        # Check if application endpoint is auto attendant or call queue

                        if ($ShowTransferCallToTargetType -eq $true) {

                            $defaultCallFlowAction = "$defaultCallFlowAction <br> Voice App"       
                          
                        }

                        if ($allAutoAttendantIds -contains $MenuOption.CallTarget.Id) {

                            $MatchingAaDefaultCallFlowAa = $allAutoAttendants | Where-Object { $_.Identity -eq $MenuOption.CallTarget.Id }

                            $defaultCallFlowTargetTypeFriendly = "[Auto Attendant"
                            $defaultCallFlowTargetName = (Optimize-DisplayName -String $($MatchingAaDefaultCallFlowAa.Name)) + "]"

                        }

                        else {

                            $MatchingCqAaDefaultCallFlow = $allCallQueues | Where-Object { $_.Identity -eq $MenuOption.CallTarget.Id }

                            $defaultCallFlowTargetTypeFriendly = "[Call Queue"
                            $defaultCallFlowTargetName = (Optimize-DisplayName -String $($MatchingCqAaDefaultCallFlow.Name)) + "]"

                        }

                        $defaultCallFlowVoicemailSystemGreeting = $null

                    }
                    SharedVoicemail {

                        $defaultCallFlowTargetTypeFriendly = "Shared Voicemail"
                        $defaultCallFlowTargetGroup = (Get-MgGroup -GroupId $MenuOption.CallTarget.Id)
                        $defaultCallFlowTargetName = Optimize-DisplayName -String $defaultCallFlowTargetGroup.DisplayName
                        $defaultCallFlowTargetIdentity = $defaultCallFlowTargetGroup.Id

                        if ($MenuOption.CallTarget.EnableSharedVoicemailSystemPromptSuppression -eq $false) {
                            
                            $defaultCallFlowVoicemailSystemGreeting = "defaultCallFlowSystemGreeting$($aaDefaultCallFlowAaObjectId)-$defaultCallFlowSharedVoicemailCounter>Greeting <br> MS System Message] -->"

                            $defaultCallFlowVoicemailSystemGreetingValue = (. Get-MsSystemMessage)[-1]
                            $defaultCallFlowVoicemailSystemGreetingValueExport = (. Get-MsSystemMessage)[0]

                            if ($ShowTTSGreetingText) {
            
                                if ($ExportTTSGreetings) {
            
                                    $defaultCallFlowVoicemailSystemGreetingValueExport | Out-File "$FilePath\$($aadefaultCallFlowAaObjectId)_autoAttendantdefaultCallFlowMsSystemMessage.txt"
                    
                                    $ttsGreetings += ("click defaultCallFlowSystemGreeting$($aaDefaultCallFlowAaObjectId)-$defaultCallFlowSharedVoicemailCounter " + '"' + "$FilePath\$($aadefaultCallFlowAaObjectId)_autoAttendantdefaultCallFlowMsSystemMessage.txt" + '"')
                    
                                }    
                
            
                                if ($defaultCallFlowVoicemailSystemGreetingValue.Length -gt $truncateGreetings) {
            
                                    $defaultCallFlowVoicemailSystemGreetingValue = $defaultCallFlowVoicemailSystemGreetingValue.Remove($defaultCallFlowVoicemailSystemGreetingValue.Length - ($defaultCallFlowVoicemailSystemGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
            
                                }
            
                                $defaultCallFlowVoicemailSystemGreeting = $defaultCallFlowVoicemailSystemGreeting.Replace("]"," <br> ''$defaultCallFlowVoicemailSystemGreetingValue'']")
            
                            }
            
                            $defaultCallFlowTransferGreetingValue = $null

                        }

                        else {
                            
                            $defaultCallFlowVoicemailSystemGreeting = $null

                        }

                        if ($ShowSharedVoicemailGroupMembers -eq $true) {

                            . Get-SharedVoicemailGroupMembers -SharedVoicemailGroupId $MenuOption.CallTarget.Id

                            $defaultCallFlowTargetName = "$defaultCallFlowTargetName$mdSharedVoicemailGroupMembers"

                        }

                        if ($MenuOption.DtmfResponse -ne "Automatic") {

                            if ($ShowTTSGreetingText) {

                                $defaultCallFlowTransferGreetingValue = (. Get-IvrTransferMessage)[0]
                                $defaultCallFlowTransferGreetingValueExport = (. Get-IvrTransferMessage)[1]
            
                                if ($ExportTTSGreetings) {
            
                                    $defaultCallFlowTransferGreetingValueExport | Out-File "$FilePath\$($aadefaultCallFlowAaObjectId)_autoAttendantdefaultCallFlowTransferSharedVoicemailGreeting.txt"
                                    $ttsGreetings += ("click defaultCallFlowTransferSharedVoicemailGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowSharedVoicemailCounter " + '"' + "$FilePath\$($aadefaultCallFlowAaObjectId)_autoAttendantdefaultCallFlowTransferSharedVoicemailGreeting.txt" + '"')
            
                                }
            
                                if ($defaultCallFlowTransferGreetingValue.Length -gt $truncateGreetings) {
            
                                    $defaultCallFlowTransferGreetingValue = $defaultCallFlowTransferGreetingValue.Remove($defaultCallFlowTransferGreetingValue.Length - ($defaultCallFlowTransferGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                        
                                }
            
                                $defaultCallFlowTransferGreetingValue = "defaultCallFlowTransferSharedVoicemailGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowSharedVoicemailCounter>Greeting<br>Transfer Message<br>''$defaultCallFlowTransferGreetingValue''] -->"
                
                            }
            
                            else {
            
                                $defaultCallFlowTransferGreetingValueExport = $null
                                $defaultCallFlowTransferGreetingValue = "defaultCallFlowTransferSharedVoicemailGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowSharedVoicemailCounter>Greeting<br>Transfer Message] -->"
            
                            }

                        }

                        else {
            
                            $defaultCallFlowTransferGreetingValueExport = $null
                            $defaultCallFlowTransferGreetingValue = $null
        
                        }
        
                        $allMermaidNodes += "defaultCallFlowTransferSharedVoicemailGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowSharedVoicemailCounter"

                        $allMermaidNodes += "defaultCallFlowSystemGreeting$($aaDefaultCallFlowAaObjectId)-$defaultCallFlowSharedVoicemailCounter"

                        $defaultCallFlowSharedVoicemailCounter ++

                    }
                }

                # Check if transfer target type is call queue
                if ($defaultCallFlowTargetTypeFriendly -eq "[Call Queue") {

                    $MatchingCQIdentity = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id -or $_.Identity -eq $MenuOption.CallTarget.Id}).Identity

                    $mdAutoAttendantDefaultCallFlow = "$mdDtmfLink defaultCallFlow$($aaDefaultCallFlowAaObjectId)$DtmfKey($defaultCallFlowAction) --> $($MatchingCQIdentity)($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName)`n"
                    
                    if ($nestedVoiceApps -notcontains $MatchingCQIdentity) {

                        $nestedVoiceApps += $MatchingCQIdentity

                    }

                    $allMermaidNodes += @("defaultCallFlow$($aadefaultCallFlowAaObjectId)$DtmfKey","$($MatchingCQIdentity)")

                
                } # End if transfer target type is call queue

                elseif ($defaultCallFlowTargetTypeFriendly -eq "[Auto Attendant") {

                    $mdAutoAttendantDefaultCallFlow = "$mdDtmfLink defaultCallFlow$($aaDefaultCallFlowAaObjectId)$DtmfKey($defaultCallFlowAction) --> $($MatchingAaDefaultCallFlowAa.Identity)($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName)`n"

                    if ($nestedVoiceApps -notcontains $MatchingAaDefaultCallFlowAa.Identity) {

                        $nestedVoiceApps += $MatchingAaDefaultCallFlowAa.Identity

                    }

                    $allMermaidNodes += @("defaultCallFlow$($aadefaultCallFlowAaObjectId)$DtmfKey","$($MatchingAaDefaultCallFlowAa.Identity)")

                }

                # Check if default callflow action target is trasnfer call to target but something other than call queue
                else {

                    $mdAutoAttendantDefaultCallFlow = $mdDtmfLink + "$defaultCallFlowTransferGreetingValue defaultCallFlow$($aaDefaultCallFlowAaObjectId)$DtmfKey($defaultCallFlowAction) --> $defaultCallFlowVoicemailSystemGreeting $($defaultCallFlowTargetIdentity)($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName)`n"

                    $allMermaidNodes += @("defaultCallFlow$($aadefaultCallFlowAaObjectId)$DtmfKey","$($defaultCallFlowTargetIdentity)")

                }

            }

            $mdAutoAttendantDefaultCallFlowMenuOptions += $mdAutoAttendantDefaultCallFlow

        }

        # Greeting can only exist once. Add greeting before call flow and set mdAutoAttendantDefaultCallFlow to the new variable.

            $mdDefaultCallFlowGreeting += $mdAutoAttendantDefaultCallFlowMenuOptions
            $mdAutoAttendantDefaultCallFlow = $mdDefaultCallFlowGreeting
    
    }

    # Remove Greeting node, if none is configured
    if ($defaultCallFlowGreeting -eq "Greeting <br> None") {

        $mdAutoAttendantDefaultCallFlow = ($mdAutoAttendantDefaultCallFlow.Replace("defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowGreeting] --> ","")).TrimStart()

    }
    
}

function Get-AutoAttendantAfterHoursCallFlow {
    param (
        [Parameter(Mandatory = $false)][String]$VoiceAppId
    )

    $aaAfterHoursCallFlowAaObjectId = $aa.Identity

    $languageId = $aa.LanguageId

    # Get after hours call flow
    $afterHoursAssociatedCallFlowId = ($aa.CallHandlingAssociations | Where-Object {$_.Type -eq "AfterHours"}).CallFlowId
    $afterHoursCallFlow = ($aa.CallFlows | Where-Object {$_.Id -eq $afterHoursAssociatedCallFlowId})
    $afterHoursCallFlowAction = ($aa.CallFlows | Where-Object {$_.Id -eq $afterHoursAssociatedCallFlowId}).Menu.MenuOptions.Action

    . Get-AutoAttendantDirectorySearchConfig -CallFlowType "afterHoursCallFlow"

    # Get after hours greeting
    if (!$afterHoursCallFlow.Greetings.ActiveType){
        $afterHoursCallFlowGreeting = "Greeting <br> None"
    }

    else {
        $afterHoursCallFlowGreeting = "Greeting <br> $($afterHoursCallFlow.Greetings.ActiveType)"

        if ($($afterHoursCallFlow.Greetings.ActiveType) -eq "TextToSpeech" -and $ShowTTSGreetingText) {

            $audioFileName = $null

            $afterHoursTTSGreetingValueExport = $afterHoursCallFlow.Greetings.TextToSpeechPrompt
            $afterHoursTTSGreetingValue = Optimize-DisplayName -String $afterHoursCallFlow.Greetings.TextToSpeechPrompt

            if ($ExportTTSGreetings) {

                $afterHoursTTSGreetingValueExport | Out-File "$FilePath\$($aaDefaultCallFlowAaObjectId)_afterHoursCallFlowGreeting.txt"

                $ttsGreetings += ("click afterHoursCallFlowGreeting$($aaDefaultCallFlowAaObjectId) " + '"' + "$FilePath\$($aaDefaultCallFlowAaObjectId)_afterHoursCallFlowGreeting.txt" + '"')

            }


            if ($afterHoursTTSGreetingValue.Length -gt $truncateGreetings) {

                $afterHoursTTSGreetingValue = $afterHoursTTSGreetingValue.Remove($afterHoursTTSGreetingValue.Length - ($afterHoursTTSGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
            
            }

            $afterHoursCallFlowGreeting += " <br> ''$afterHoursTTSGreetingValue''"
        
        }

        elseif ($($afterHoursCallFlow.Greetings.ActiveType) -eq "AudioFile" -and $ShowAudioFileName) {

            $afterHoursTTSGreetingValue = $null

            # Audio File Greeting Name
            $audioFileName = Optimize-DisplayName -String ($afterHoursCallFlow.Greetings.AudioFilePrompt.FileName)

            if ($ExportAudioFiles) {

                $content = Export-CsOnlineAudioFile -Identity $afterHoursCallFlow.Greetings.AudioFilePrompt.Id -ApplicationId OrgAutoAttendant
                [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)

                $audioFileNames += ("click afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId) " + '"' + "$FilePath\$audioFileName" + '"')

            }


            if ($audioFileName.Length -gt $truncateGreetings) {

                $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"

            }

            $afterHoursCallFlowGreeting += " <br> $audioFileName"


        }

    }

    # Check if the after hours callflow action is disconnect call
    if ($afterHoursCallFlowAction -eq "DisconnectCall") {

        if ($CombineDisconnectCallNodes -eq $true) {

            $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId)>$AfterHoursCallFlowGreeting] --> disconnectCall(($afterHoursCallFlowAction))`n"

            $allMermaidNodes += @("afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId)","disconnectCall")

        }

        else {

            $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId)>$AfterHoursCallFlowGreeting] --> afterHoursCallFlow$($aaAfterHoursCallFlowAaObjectId)(($afterHoursCallFlowAction))`n"

            $allMermaidNodes += @("afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId)","afterHoursCallFlow$($aaAfterHoursCallFlowAaObjectId)")
    
        }

    }
    
    # if after hours action is not disconnect call
    else  {

        $afterHoursCallFlowMenuOptions = $afterHoursCallFlow.Menu.MenuOptions

        # Check if IVR is disabled
        if ($afterHoursCallFlowMenuOptions.Count -lt 2 -and !$afterHoursCallFlow.Menu.Prompts.ActiveType) {

            $mdafterHoursCallFlowGreeting = "afterHoursCallFlowGreeting$($aaafterHoursCallFlowAaObjectId)>$afterHoursCallFlowGreeting] --> "

            $allMermaidNodes += @("afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId)")

            $afterHoursCallFlowMenuOptionsKeyPress = $null

            $mdAutoAttendantafterHoursCallFlowMenuOptions = $null

        }

        else {

            $afterHoursCallFlowMenuOptionsGreeting = "IVR Greeting <br> $($afterHoursCallFlow.Menu.Prompts.ActiveType)"

            if ($($afterHoursCallFlow.Menu.Prompts.ActiveType) -eq "TextToSpeech" -and $ShowTTSGreetingText) {

                $audioFileName = $null
    
                $afterHoursCallFlowMenuOptionsTTSGreetingValueExport = $afterHoursCallFlow.Menu.Prompts.TextToSpeechPrompt
                $afterHoursCallFlowMenuOptionsTTSGreetingValue = Optimize-DisplayName -String $afterHoursCallFlow.Menu.Prompts.TextToSpeechPrompt

                if ($ExportTTSGreetings) {

                    $afterHoursCallFlowMenuOptionsTTSGreetingValueExport | Out-File "$FilePath\$($aaafterHoursCallFlowAaObjectId)_afterHoursCallFlowMenuOptionsGreeting.txt"
    
                    $ttsGreetings += ("click afterHoursCallFlowMenuOptionsGreeting$($aaafterHoursCallFlowAaObjectId) " + '"' + "$FilePath\$($aaafterHoursCallFlowAaObjectId)_afterHoursCallFlowMenuOptionsGreeting.txt" + '"')
    
                }    

    
                if ($afterHoursCallFlowMenuOptionsTTSGreetingValue.Length -gt $truncateGreetings) {
    
                    $afterHoursCallFlowMenuOptionsTTSGreetingValue = $afterHoursCallFlowMenuOptionsTTSGreetingValue.Remove($afterHoursCallFlowMenuOptionsTTSGreetingValue.Length - ($afterHoursCallFlowMenuOptionsTTSGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                
                }
    
                $afterHoursCallFlowMenuOptionsGreeting += " <br> ''$afterHoursCallFlowMenuOptionsTTSGreetingValue''"
            
            }
    
            elseif ($($afterHoursCallFlow.Menu.Prompts.ActiveType) -eq "AudioFile" -and $ShowAudioFileName) {
    
                $afterHoursCallFlowMenuOptionsTTSGreetingValue = $null
    
                # Audio File Greeting Name
                $audioFileName = Optimize-DisplayName -String ($afterHoursCallFlow.Menu.Prompts.AudioFilePrompt.FileName)

                if ($ExportAudioFiles) {

                    $content = Export-CsOnlineAudioFile -Identity $afterHoursCallFlow.Menu.Prompts.AudioFilePrompt.Id -ApplicationId OrgAutoAttendant
                    [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
    
                    $audioFileNames += ("click afterHoursCallFlowMenuOptionsGreeting$($aaAfterHoursCallFlowAaObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
    
                }
    
                if ($audioFileName.Length -gt $truncateGreetings) {
    
                    $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                    $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
    
                }
    
                $afterHoursCallFlowMenuOptionsGreeting += " <br> $audioFileName"
    
    
            }

            if ($afterHoursCallFlow.ForceListenMenuEnabled -eq $true) {

                $afterHoursCallFlowForceListen = "<br> Force Listen: True"

            }

            else {

                $afterHoursCallFlowForceListen = "<br> Force Listen: False"

            }

            if ($aaIsVoiceResponseEnabled) {

                $afterHoursCallFlowVoiceResponse = " or <br> Voice Response <br> Language: $($aa.LanguageId)<br>Voice Style: $($aa.VoiceId)$afterHoursCallFlowForceListen"

            }

            else {

                $afterHoursCallFlowVoiceResponse = $afterHoursCallFlowForceListen

            }

    
            $afterHoursCallFlowMenuOptionsKeyPress = @"

afterHoursCallFlowMenuOptions$($aaAfterHoursCallFlowAaObjectId){Key Press$afterHoursCallFlowVoiceResponse}
"@

            $mdafterHoursCallFlowGreeting =@"
afterHoursCallFlowGreeting$($aaafterHoursCallFlowAaObjectId)>$afterHoursCallFlowGreeting] --> afterHoursCallFlowMenuOptionsGreeting$($aaAfterHoursCallFlowAaObjectId)>$afterHoursCallFlowMenuOptionsGreeting] --> $afterHoursCallFlowMenuOptionsKeyPress

"@

            $mdAutoAttendantafterHoursCallFlowMenuOptions =@"

"@

            $allMermaidNodes += @("afterHoursCallFlowMenuOptions$($aaafterHoursCallFlowAaObjectId)","afterHoursCallFlowGreeting$($aaafterHoursCallFlowAaObjectId)","afterHoursCallFlowMenuOptionsGreeting$($aaafterHoursCallFlowAaObjectId)")

        }

        $afterHoursCallFlowSharedVoicemailCounter = 1
        $afterHoursCallFlowUserOptionCounter = 1
        $afterHoursCallFlowPSTNOptionCounter = 1

        foreach ($MenuOption in $afterHoursCallFlowMenuOptions) {

            # Check if auto attendant has no IVR / menu options configured
            if ($afterHoursCallFlowMenuOptions.Count -lt 2 -and !$afterHoursCallFlow.Menu.Prompts.ActiveType) {

                $mdDtmfLink = $null
                $DtmfKey = $null
                $voiceResponse = $null

            }

            # Auto attendant has an IVR menu options
            else {

                if ($aaIsVoiceResponseEnabled) {

                    if ($MenuOption.VoiceResponses) {

                        $voiceResponse = "/ Voice Response: ''$($MenuOption.VoiceResponses)''"

                    }

                    else {

                        $voiceResponse = "/ No Voice Response Configured"

                    }

                }

                else {

                    $voiceResponse = $null

                }

                [String]$DtmfKey = ($MenuOption.DtmfResponse)

                $DtmfKey = $DtmfKey.Replace("Tone","")

                $mdDtmfLink = "$afterHoursCallFlowMenuOptionsKeyPress --> |$DtmfKey $voiceResponse|"

            }

            # Get transfer target type
            $afterHoursCallFlowTargetType = $MenuOption.CallTarget.Type
            $afterHoursCallFlowAction = $MenuOption.Action

            if ($afterHoursCallFlowAction -eq "TransferCallToOperator") {

                if ($aaIsVoiceResponseEnabled) {

                    $afterHoursCallFlowOperatorVoiceResponse = "/ Voice Response: ''Operator''"

                    $mdDtmfLink = $mdDtmfLink.Replace($voiceResponse,$afterHoursCallFlowOperatorVoiceResponse)

                }

                else {

                    $afterHoursCallFlowOperatorVoiceResponse = $null

                }

                if ($ShowTTSGreetingText) {

                    $afterHoursCallFlowTransferGreetingOperatorValue = (. Get-IvrTransferMessage)[2]
                    $afterHoursCallFlowTransferGreetingOperatorValueExport = (. Get-IvrTransferMessage)[3]

                    if ($ExportTTSGreetings) {

                        $afterHoursCallFlowTransferGreetingOperatorValueExport | Out-File "$FilePath\$($aaAfterHoursCallFlowAaObjectId)_autoAttendantafterHoursCallFlowTransferGreetingOperator.txt"
                        $ttsGreetings += ("click afterHoursCallFlowTransferGreetingOperator$($aaAfterHoursCallFlowAaObjectId) " + '"' + "$FilePath\$($aaAfterHoursCallFlowAaObjectId)_autoAttendantafterHoursCallFlowTransferGreetingOperator.txt" + '"')

                    }

                    if ($afterHoursCallFlowTransferGreetingOperatorValue.Length -gt $truncateGreetings) {

                        $afterHoursCallFlowTransferGreetingOperatorValue = $afterHoursCallFlowTransferGreetingOperatorValue.Remove($afterHoursCallFlowTransferGreetingOperatorValue.Length - ($afterHoursCallFlowTransferGreetingOperatorValue.Length -$truncateGreetings)).TrimEnd() + "..."
            
                    }

                    $afterHoursCallFlowTransferGreetingOperatorValue = "afterHoursCallFlowTransferGreetingOperator$($aaAfterHoursCallFlowAaObjectId)>Greeting<br>Transfer Message<br>''$afterHoursCallFlowTransferGreetingOperatorValue''] -->"
    
                }

                else {

                    $afterHoursCallFlowTransferGreetingOperatorValueExport = $null
                    $afterHoursCallFlowTransferGreetingOperatorValue = "afterHoursCallFlowTransferGreetingOperator$($aaAfterHoursCallFlowAaObjectId)>Greeting<br>Transfer Message] -->"

                }

                $allMermaidNodes += "afterHoursCallFlowTransferGreetingOperator$($aaAfterHoursCallFlowAaObjectId)"

                $mdAutoAttendantafterHoursCallFlow = "$mdDtmfLink $afterHoursCallFlowTransferGreetingOperatorValue afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey($afterHoursCallFlowAction) --> $($OperatorIdentity)($OperatorTypeFriendly <br> $OperatorName)`n"

                $allMermaidNodes += @("afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey","$($OperatorIdentity)")

                $afterHoursCallFlowVoicemailSystemGreeting = $null

                if ($nestedVoiceApps -notcontains $OperatorIdentity -and $AddOperatorToNestedVoiceApps -eq $true) {

                    $nestedVoiceApps += $OperatorIdentity

                }

                if ($FindUserLinks -eq $true) {
         
                    . New-VoiceAppUserLinkProperties -userLinkUserId $OperatorIdentity -userLinkUserName $OperatorName -userLinkVoiceAppType "Auto Attendant" -userLinkVoiceAppActionType "TransferCallToOperator (DTMF Option: $DtmfKey)" -userLinkVoiceAppName $aa.Name -userLinkVoiceAppId $aa.Identity
                
                }

            }

            elseif ($afterHoursCallFlowAction -eq "Announcement") {
                
                $voiceMenuOptionAnnouncementType = $MenuOption.Prompt.ActiveType

                $afterHoursCallFlowMenuOptionsAnnouncement = "$voiceMenuOptionAnnouncementType"

                if ($voiceMenuOptionAnnouncementType -eq "TextToSpeech" -and $ShowTTSGreetingText) {
    
                    $audioFileName = $null
        
                    $afterHoursCallFlowMenuOptionsTTSAnnouncementValueExport = $MenuOption.Prompt.TextToSpeechPrompt
                    $afterHoursCallFlowMenuOptionsTTSAnnouncementValue = Optimize-DisplayName -String $MenuOption.Prompt.TextToSpeechPrompt

                    if ($ExportTTSGreetings) {

                        $afterHoursCallFlowMenuOptionsTTSAnnouncementValueExport | Out-File "$FilePath\$($aaafterHoursCallFlowAaObjectId)_afterHoursCallFlowMenuOptionsAnnouncement$DtmfKey.txt"
        
                        $ttsGreetings += ("click afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey " + '"' + "$FilePath\$($aaafterHoursCallFlowAaObjectId)_afterHoursCallFlowMenuOptionsAnnouncement$DtmfKey.txt" + '"')
        
                    }    

        
                    if ($afterHoursCallFlowMenuOptionsTTSAnnouncementValue.Length -gt $truncateGreetings) {
        
                        $afterHoursCallFlowMenuOptionsTTSAnnouncementValue = $afterHoursCallFlowMenuOptionsTTSAnnouncementValue.Remove($afterHoursCallFlowMenuOptionsTTSAnnouncementValue.Length - ($afterHoursCallFlowMenuOptionsTTSAnnouncementValue.Length -$truncateGreetings)).TrimEnd() + "..."
                    
                    }
        
                    $afterHoursCallFlowMenuOptionsAnnouncement += " <br> ''$afterHoursCallFlowMenuOptionsTTSAnnouncementValue''"
                
                }
        
                elseif ($voiceMenuOptionAnnouncementType -eq "AudioFile" -and $ShowAudioFileName) {
        
                    $afterHoursCallFlowMenuOptionsTTSAnnouncementValue = $null
        
                    # Audio File Announcement Name
                    $audioFileName = Optimize-DisplayName -String ($MenuOption.Prompt.AudioFilePrompt.FileName)

                    if ($ExportAudioFiles) {

                        $content = Export-CsOnlineAudioFile -Identity $MenuOption.Prompt.AudioFilePrompt.Id -ApplicationId OrgAutoAttendant
                        [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                        $audioFileNames += ("click afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey " + '"' + "$FilePath\$audioFileName" + '"')
        
                    }

                    
                    if ($audioFileName.Length -gt $truncateGreetings) {
        
                        $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                        $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                    }
        
                    $afterHoursCallFlowMenuOptionsAnnouncement += " <br> $audioFileName"
        
        
                }

                $mdAutoAttendantafterHoursCallFlow = "$mdDtmfLink afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey>$afterHoursCallFlowAction <br> $afterHoursCallFlowMenuOptionsAnnouncement] ---> afterHoursCallFlowMenuOptionsGreeting$($aaafterHoursCallFlowAaObjectId)`n"

                $allMermaidNodes += @("afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey","afterHoursCallFlowMenuOptionsGreeting$($aaafterHoursCallFlowAaObjectId)")

                $afterHoursCallFlowVoicemailSystemGreeting = $null

            }

            else {

                # Switch through transfer target type and set variables accordingly
                switch ($afterHoursCallFlowTargetType) {
                    User { 
                        $afterHoursCallFlowTargetTypeFriendly = "User"
                        $afterHoursCallFlowTargetUser = (Get-MgUser -UserId $($MenuOption.CallTarget.Id))
                        $afterHoursCallFlowTargetName = Optimize-DisplayName -String $afterHoursCallFlowTargetUser.DisplayName
                        $afterHoursCallFlowTargetIdentity = $afterHoursCallFlowTargetUser.Id

                        if ($FindUserLinks -eq $true) {

                            if ($DtmfKey) {

                                $DtmfOption = " (DTMF Option: $DtmfKey)"

                            }

                            else {

                                $DtmfOption =$null

                            }
         
                            . New-VoiceAppUserLinkProperties -userLinkUserId $($MenuOption.CallTarget.Id) -userLinkUserName $afterHoursCallFlowTargetUser.DisplayName -userLinkVoiceAppType "Auto Attendant" -userLinkVoiceAppActionType "AfterHoursCallFlowTransferCallToTarget$DtmfOption" -userLinkVoiceAppName $aa.Name -userLinkVoiceAppId $aa.Identity
                        
                        }                

                        if ($nestedVoiceApps -notcontains $afterHoursCallFlowTargetUser.Id) {

                            $nestedVoiceApps += $afterHoursCallFlowTargetUser.Id

                        }

                        $afterHoursCallFlowVoicemailSystemGreeting = $null

                        if ($MenuOption.DtmfResponse -ne "Automatic") {

                            if ($ShowTTSGreetingText) {

                                $afterHoursCallFlowTransferGreetingValue = (. Get-IvrTransferMessage)[0]
                                $afterHoursCallFlowTransferGreetingValueExport = (. Get-IvrTransferMessage)[1]
            
                                if ($ExportTTSGreetings) {
            
                                    $afterHoursCallFlowTransferGreetingValueExport | Out-File "$FilePath\$($aaafterHoursCallFlowAaObjectId)_autoAttendantafterHoursCallFlowTransferUserGreeting.txt"
                                    $ttsGreetings += ("click afterHoursCallFlowTransferUserGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowUserOptionCounter " + '"' + "$FilePath\$($aaafterHoursCallFlowAaObjectId)_autoAttendantafterHoursCallFlowTransferUserGreeting.txt" + '"')
            
                                }
            
                                if ($afterHoursCallFlowTransferGreetingValue.Length -gt $truncateGreetings) {
            
                                    $afterHoursCallFlowTransferGreetingValue = $afterHoursCallFlowTransferGreetingValue.Remove($afterHoursCallFlowTransferGreetingValue.Length - ($afterHoursCallFlowTransferGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                        
                                }
            
                                $afterHoursCallFlowTransferGreetingValue = "afterHoursCallFlowTransferUserGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowUserOptionCounter>Greeting<br>Transfer Message<br>''$afterHoursCallFlowTransferGreetingValue''] -->"
                
                            }
            
                            else {
            
                                $afterHoursCallFlowTransferGreetingValueExport = $null
                                $afterHoursCallFlowTransferGreetingValue = "afterHoursCallFlowTransferUserGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowUserOptionCounter>Greeting<br>Transfer Message] -->"
            
                            }

                        }

                        else {

                            $afterHoursCallFlowTransferGreetingValueExport = $null
                            $afterHoursCallFlowTransferGreetingValue = $null

                        }
        
                        $allMermaidNodes += "afterHoursCallFlowTransferUserGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowUserOptionCounter"

                        $afterHoursCallFlowUserOptionCounter ++

                    }
                    ExternalPstn { 
                        $afterHoursCallFlowTargetTypeFriendly = "External Number"
                        $afterHoursCallFlowTargetName = ($MenuOption.CallTarget.Id).Replace("tel:","")
                        $afterHoursCallFlowTargetIdentity = $afterHoursCallFlowTargetName

                        if ($ObfuscatePhoneNumbers -eq $true) {

                            $afterHoursCallFlowTargetName = $afterHoursCallFlowTargetName.Remove(($afterHoursCallFlowTargetName.Length -4)) + "****"
        
                        }

                        $afterHoursCallFlowVoicemailSystemGreeting = $null

                        if ($MenuOption.DtmfResponse -ne "Automatic") {

                            if ($ShowTTSGreetingText) {

                                $afterHoursCallFlowTransferGreetingValue = (. Get-IvrTransferMessage)[0]
                                $afterHoursCallFlowTransferGreetingValueExport = (. Get-IvrTransferMessage)[1]

                                if ($ExportTTSGreetings) {

                                    $afterHoursCallFlowTransferGreetingValueExport | Out-File "$FilePath\$($aaafterHoursCallFlowAaObjectId)_autoAttendantafterHoursCallFlowTransferPSTNGreeting.txt"
                                    $ttsGreetings += ("click afterHoursCallFlowTransferPSTNGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowPSTNOptionCounter " + '"' + "$FilePath\$($aaafterHoursCallFlowAaObjectId)_autoAttendantafterHoursCallFlowTransferPSTNGreeting.txt" + '"')

                                }

                                if ($afterHoursCallFlowTransferGreetingValue.Length -gt $truncateGreetings) {

                                    $afterHoursCallFlowTransferGreetingValue = $afterHoursCallFlowTransferGreetingValue.Remove($afterHoursCallFlowTransferGreetingValue.Length - ($afterHoursCallFlowTransferGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."

                                }

                                $afterHoursCallFlowTransferGreetingValue = "afterHoursCallFlowTransferPSTNGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowPSTNOptionCounter>Greeting<br>Transfer Message<br>''$afterHoursCallFlowTransferGreetingValue''] -->"

                            }

                            else {

                                $afterHoursCallFlowTransferGreetingValueExport = $null
                                $afterHoursCallFlowTransferGreetingValue = "afterHoursCallFlowTransferPSTNGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowPSTNOptionCounter>Greeting<br>Transfer Message] -->"

                            }

                        }

                        else {

                            $afterHoursCallFlowTransferGreetingValueExport = $null
                            $afterHoursCallFlowTransferGreetingValue = $null

                        }

                        $allMermaidNodes += "afterHoursCallFlowTransferPSTNGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowPSTNOptionCounter"

                        $afterHoursCallFlowPSTNOptionCounter ++

                    }
                    ApplicationEndpoint {

                        # Check if application endpoint is auto attendant or call queue

                        if ($ShowTransferCallToTargetType -eq $true) {

                            $afterHoursCallFlowAction = "$afterHoursCallFlowAction <br> Resource Account"       
                          
                        }

                        $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $MenuOption.CallTarget.Id -and $_.ApplicationId -eq $applicationIdAa}

                        if ($matchingApplicationInstanceCheckAa) {

                            $MatchingAaafterHoursCallFlowAa = $allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id -or $_.Identity -eq $MenuOption.CallTarget.Id}

                            $afterHoursCallFlowTargetTypeFriendly = "[Auto Attendant"
                            $afterHoursCallFlowTargetName = (Optimize-DisplayName -String $($MatchingAaafterHoursCallFlowAa.Name)) + "]"

                        }

                        else {

                            $MatchingCqAaafterHoursCallFlow = $allCallQueues | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id -or $_.Identity -eq $MenuOption.CallTarget.Id}

                            $afterHoursCallFlowTargetTypeFriendly = "[Call Queue"
                            $afterHoursCallFlowTargetName = (Optimize-DisplayName -String $($MatchingCqAaafterHoursCallFlow.Name)) + "]"

                        }

                        $afterHoursCallFlowVoicemailSystemGreeting = $null

                    }
                    ConfigurationEndpoint {

                        # Check if application endpoint is auto attendant or call queue

                        if ($ShowTransferCallToTargetType -eq $true) {

                            $afterHoursCallFlowAction = "$afterHoursCallFlowAction <br> Voice App"

                        }

                        if ($allAutoAttendantIds -contains $MenuOption.CallTarget.Id) {

                            $MatchingAaafterHoursCallFlowAa = $allAutoAttendants | Where-Object { $_.Identity -eq $MenuOption.CallTarget.Id }

                            $afterHoursCallFlowTargetTypeFriendly = "[Auto Attendant"
                            $afterHoursCallFlowTargetName = (Optimize-DisplayName -String $($MatchingAaafterHoursCallFlowAa.Name)) + "]"

                        }

                        else {

                            $MatchingCqAaafterHoursCallFlow = $allCallQueues | Where-Object { $_.Identity -eq $MenuOption.CallTarget.Id }

                            $afterHoursCallFlowTargetTypeFriendly = "[Call Queue"
                            $afterHoursCallFlowTargetName = (Optimize-DisplayName -String $($MatchingCqAaafterHoursCallFlow.Name)) + "]"

                        }

                        $afterHoursCallFlowVoicemailSystemGreeting = $null

                    }
                    SharedVoicemail {

                        $afterHoursCallFlowTargetTypeFriendly = "Shared Voicemail"
                        $afterHoursCallFlowTargetGroup = (Get-MgGroup -GroupId $MenuOption.CallTarget.Id)
                        $afterHoursCallFlowTargetName = Optimize-DisplayName -String $afterHoursCallFlowTargetGroup.DisplayName
                        $afterHoursCallFlowTargetIdentity = $afterHoursCallFlowTargetGroup.Id

                        if ($MenuOption.CallTarget.EnableSharedVoicemailSystemPromptSuppression -eq $false) {
                            
                            $afterHoursCallFlowVoicemailSystemGreeting = "afterHoursCallFlowSystemGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowSharedVoicemailCounter>Greeting <br> MS System Message] -->"

                            $afterHoursCallFlowVoicemailSystemGreetingValue = (. Get-MsSystemMessage)[-1]
                            $afterHoursCallFlowVoicemailSystemGreetingValueExport = (. Get-MsSystemMessage)[0]

                            if ($ShowTTSGreetingText) {
            
                                if ($ExportTTSGreetings) {
            
                                    $afterHoursCallFlowVoicemailSystemGreetingValueExport | Out-File "$FilePath\$($aaafterHoursCallFlowAaObjectId)_autoAttendantAfterHoursCallFlowMsSystemMessage.txt"
                    
                                    $ttsGreetings += ("click afterHoursCallFlowSystemGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowSharedVoicemailCounter " + '"' + "$FilePath\$($aaafterHoursCallFlowAaObjectId)_autoAttendantAfterHoursCallFlowMsSystemMessage.txt" + '"')
                    
                                }    
                
            
                                if ($afterHoursCallFlowVoicemailSystemGreetingValue.Length -gt $truncateGreetings) {
            
                                    $afterHoursCallFlowVoicemailSystemGreetingValue = $afterHoursCallFlowVoicemailSystemGreetingValue.Remove($afterHoursCallFlowVoicemailSystemGreetingValue.Length - ($afterHoursCallFlowVoicemailSystemGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
            
                                }
            
                                $afterHoursCallFlowVoicemailSystemGreeting = $afterHoursCallFlowVoicemailSystemGreeting.Replace("]"," <br> ''$afterHoursCallFlowVoicemailSystemGreetingValue'']")
            
                            }

                            $afterHoursCallFlowTransferGreetingValue = $null            

                        }

                        else {
                            
                            $afterHoursCallFlowVoicemailSystemGreeting = $null

                        }

                        if ($ShowSharedVoicemailGroupMembers -eq $true) {

                            . Get-SharedVoicemailGroupMembers -SharedVoicemailGroupId $MenuOption.CallTarget.Id

                            $afterHoursCallFlowTargetName = "$afterHoursCallFlowTargetName$mdSharedVoicemailGroupMembers"

                        }

                        if ($MenuOption.DtmfResponse -ne "Automatic") {

                            if ($ShowTTSGreetingText) {

                                $afterHoursCallFlowTransferGreetingValue = (. Get-IvrTransferMessage)[0]
                                $afterHoursCallFlowTransferGreetingValueExport = (. Get-IvrTransferMessage)[1]
            
                                if ($ExportTTSGreetings) {
            
                                    $afterHoursCallFlowTransferGreetingValueExport | Out-File "$FilePath\$($aaafterHoursCallFlowAaObjectId)_autoAttendantafterHoursCallFlowTransferSharedVoicemailGreeting.txt"
                                    $ttsGreetings += ("click afterHoursCallFlowTransferSharedVoicemailGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowSharedVoicemailCounter " + '"' + "$FilePath\$($aaafterHoursCallFlowAaObjectId)_autoAttendantafterHoursCallFlowTransferSharedVoicemailGreeting.txt" + '"')
            
                                }
            
                                if ($afterHoursCallFlowTransferGreetingValue.Length -gt $truncateGreetings) {
            
                                    $afterHoursCallFlowTransferGreetingValue = $afterHoursCallFlowTransferGreetingValue.Remove($afterHoursCallFlowTransferGreetingValue.Length - ($afterHoursCallFlowTransferGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                        
                                }
            
                                $afterHoursCallFlowTransferGreetingValue = "afterHoursCallFlowTransferSharedVoicemailGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowSharedVoicemailCounter>Greeting<br>Transfer Message<br>''$afterHoursCallFlowTransferGreetingValue''] -->"
                
                            }
            
                            else {
            
                                $afterHoursCallFlowTransferGreetingValueExport = $null
                                $afterHoursCallFlowTransferGreetingValue = "afterHoursCallFlowTransferSharedVoicemailGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowSharedVoicemailCounter>Greeting<br>Transfer Message] -->"
            
                            }

                        }

                        else {

                            $afterHoursCallFlowTransferGreetingValueExport = $null
                            $afterHoursCallFlowTransferGreetingValue = $null

                        }
        
                        $allMermaidNodes += "afterHoursCallFlowTransferSharedVoicemailGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowSharedVoicemailCounter"

                        $allMermaidNodes += "afterHoursCallFlowSystemGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowSharedVoicemailCounter"

                        $afterHoursCallFlowSharedVoicemailCounter ++

                    }
                }

                # Check if transfer target type is call queue
                if ($afterHoursCallFlowTargetTypeFriendly -eq "[Call Queue") {

                    $MatchingCQIdentity = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id -or $_.Identity -eq $MenuOption.CallTarget.Id}).Identity

                    $mdAutoAttendantafterHoursCallFlow = "$mdDtmfLink afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey($afterHoursCallFlowAction) --> $($MatchingCQIdentity)($afterHoursCallFlowTargetTypeFriendly <br> $afterHoursCallFlowTargetName)`n"
                    
                    if ($nestedVoiceApps -notcontains $MatchingCQIdentity) {

                        $nestedVoiceApps += $MatchingCQIdentity

                    }

                    $allMermaidNodes += @("afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey","$($MatchingCQIdentity)")

                
                } # End if transfer target type is call queue

                elseif ($afterHoursCallFlowTargetTypeFriendly -eq "[Auto Attendant") {

                    $mdAutoAttendantafterHoursCallFlow = "$mdDtmfLink afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey($afterHoursCallFlowAction) --> $($MatchingAaafterHoursCallFlowAa.Identity)($afterHoursCallFlowTargetTypeFriendly <br> $afterHoursCallFlowTargetName)`n"

                    if ($nestedVoiceApps -notcontains $MatchingAaafterHoursCallFlowAa.Identity) {

                        $nestedVoiceApps += $MatchingAaafterHoursCallFlowAa.Identity

                    }

                    $allMermaidNodes += @("afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey","$($MatchingAaafterHoursCallFlowAa.Identity)")

                }

                # Check if afterHours callflow action target is trasnfer call to target but something other than call queue
                else {

                    $mdAutoAttendantafterHoursCallFlow = $mdDtmfLink + "$afterHoursCallFlowTransferGreetingValue afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey($afterHoursCallFlowAction) --> $afterHoursCallFlowVoicemailSystemGreeting $($afterHoursCallFlowTargetIdentity)($afterHoursCallFlowTargetTypeFriendly <br> $afterHoursCallFlowTargetName)`n"
                    
                    $allMermaidNodes += @("afterHoursCallFlow$($aaafterHoursCallFlowAaObjectId)$DtmfKey","$($afterHoursCallFlowTargetIdentity)")

                }

            }

            $mdAutoAttendantafterHoursCallFlowMenuOptions += $mdAutoAttendantafterHoursCallFlow

        }

        # Greeting can only exist once. Add greeting before call flow and set mdAutoAttendantafterHoursCallFlow to the new variable.

        $mdafterHoursCallFlowGreeting += $mdAutoAttendantafterHoursCallFlowMenuOptions
        $mdAutoAttendantafterHoursCallFlow = $mdafterHoursCallFlowGreeting        
    
    }

    # Remove Greeting node, if none is configured
    if ($afterHoursCallFlowGreeting -eq "Greeting <br> None") {

        $mdAutoAttendantafterHoursCallFlow = ($mdAutoAttendantafterHoursCallFlow.Replace("afterHoursCallFlowGreeting$($aaafterHoursCallFlowAaObjectId)>$afterHoursCallFlowGreeting] --> ","")).TrimStart()

    }

}


function Get-CallQueueCallFlow {
    param (
        [Parameter(Mandatory = $true)][String]$MatchingCQIdentity
    )

    $MatchingCQ = $allCallQueues | Where-Object {$_.Identity -eq $MatchingCQIdentity}

    $cqCallFlowObjectId = $MatchingCQ.Identity

    Write-Host "Reading call flow for: $($MatchingCQ.Name)" -ForegroundColor Magenta
    Write-Host "##################################################" -ForegroundColor Magenta
    Write-Host "Voice App Id: $cqCallFlowObjectId" -ForegroundColor Magenta

    if ($($MatchingCQ.Name -ne (Optimize-DisplayName -String $MatchingCQ.Name))) {

        Write-Warning -Message "The Call Queue '$($MatchingCQ.Name)' contains special characters which will be removed in the diagram. New Name: '$(Optimize-DisplayName -String $MatchingCQ.Name)'"

    }

    # Store all neccessary call queue properties in variables
    $CqName = Optimize-DisplayName -String $MatchingCQ.Name
    $CqOverFlowThreshold = $MatchingCQ.OverflowThreshold

    $CqOverFlowAction = $MatchingCQ.OverflowAction

    $CqTimeoutAction = $MatchingCQ.TimeoutAction

    $CqRoutingMethod = $MatchingCQ.RoutingMethod


    $CqTimeOut = $MatchingCQ.TimeoutThreshold
    $CqAgents = $MatchingCQ.Agents
    $CqAgentOptOut = $MatchingCQ.AllowOptOut
    $CqConferenceMode = $MatchingCQ.ConferenceMode
    $CqAgentAlertTime = $MatchingCQ.AgentAlertTime
    $CqDistributionLists = $MatchingCQ.DistributionLists
    $CqDefaultMusicOnHold = $MatchingCQ.UseDefaultMusicOnHold
    $CqWelcomeMusicFileName = $MatchingCQ.WelcomeMusicFileName
    $CqWelcomeTTSGreeting = $MatchingCQ.WelcomeTextToSpeechPrompt
    $CqLanguageId = $MatchingCQ.LanguageId
    $CqOboResourceAccountIds = $MatchingCQ.OboResourceAccountIds.Guid

    if ($cqroutingMethod -eq "LongestIdle") {

        $CqPresenceBasedRouting = $true

    }

    else {

        $CqPresenceBasedRouting = $MatchingCQ.PresenceBasedRouting

    }
    
    $languageId = $CqLanguageId

    # Check if call queue uses default music on hold
    if ($CqDefaultMusicOnHold -eq $true) {
        $CqMusicOnHold = "Default"
    }

    else {
        $CqMusicOnHold = "Custom"

        if ($ShowAudioFileName) {

            $audioFileName = Optimize-DisplayName -String ($MatchingCQ.MusicOnHoldFileName)

            if ($ExportAudioFiles) {

                if ($MatchingCQ.MusicOnHoldFileDownloadUri) {

                    Invoke-WebRequest -Uri $MatchingCQ.MusicOnHoldFileDownloadUri -OutFile "$FilePath\$audioFileName"

                }


                else {

                    $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.MusicOnHoldResourceId -ApplicationId HuntGroup
                    [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)

                }

                $audioFileNames += ("click cqSettingsContainer$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')

            }
            
            if ($audioFileName.Length -gt $truncateGreetings) {
        
                $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"

            }

            $CqMusicOnHold += " <br> MoH File: $audioFileName"    

        }

    }

    # Check if call queue uses a greeting
    if (!$CqWelcomeMusicFileName -and !$CqWelcomeTTSGreeting) {
        $CqGreeting = "None"

        $cqGreetingNode = $null
    }

    else {

        if ($CqWelcomeMusicFileName) {

            $CqGreeting = "AudioFile"

            if ($ShowAudioFileName) {

                $audioFileName = Optimize-DisplayName -String ($CqWelcomeMusicFileName)

                if ($ExportAudioFiles) {

                    if ($MatchingCQ.WelcomeMusicFileDownloadUri) {

                        Invoke-WebRequest -Uri $MatchingCQ.WelcomeMusicFileDownloadUri -OutFile "$FilePath\$audioFileName"

                    }

                    else {

                        $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.WelcomeMusicResourceId -ApplicationId HuntGroup
                        [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)

                    }

                    $audioFileNames += ("click cqGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')

                }
                
                if ($audioFileName.Length -gt $truncateGreetings) {
            
                    $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                    $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"

                }

                $CqGreeting += " <br> $audioFileName"

            }

        }

        else {

            $CqGreeting = "TextToSpeech"

            if ($ShowTTSGreetingText) {

                $welcomeTTSGreetingValueExport = $MatchingCQ.WelcomeTextToSpeechPrompt
                $welcomeTTSGreetingValue = Optimize-DisplayName -String $MatchingCQ.WelcomeTextToSpeechPrompt

                if ($ExportTTSGreetings) {

                    $welcomeTTSGreetingValueExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqWelcomeGreeting.txt"
    
                    $ttsGreetings += ("click cqGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqWelcomeGreeting.txt" + '"')
    
                }    


                if ($welcomeTTSGreetingValue.Length -gt $truncateGreetings) {

                    $welcomeTTSGreetingValue = $welcomeTTSGreetingValue.Remove($welcomeTTSGreetingValue.Length - ($welcomeTTSGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."

                }

                $cqGreeting += " <br> ''$welcomeTTSGreetingValue''"

            }

        }

        if ($MatchingCQ.OverflowThreshold -ge 1) {

            $cqGreetingNode = " cqGreeting$($cqCallFlowObjectId)>Greeting <br> $CqGreeting] -->"

        }

        else {

            $cqGreetingNode = $null

        }

    }

    # Check if call queue useses users, group or teams channel as distribution list
    if (!$CqDistributionLists) {

        $CqAgentListType = "Users"

    }

    else {

        if (!$MatchingCQ.ChannelId) {

            $CqAgentListType = "Groups"

            foreach ($DistributionList in $MatchingCQ.DistributionLists.Guid) {

                $DistributionListName = Optimize-DisplayName -String (Get-MgGroup -GroupId $DistributionList).DisplayName

                $CqAgentListType += " <br> Group Name: $DistributionListName"

            }

            if ($MatchingCQ.DistributionLists.Count -lt 2) {

                $CqAgentListType = $CqAgentListType.Replace("Groups","Group")

            }

        }

        else {

            $TeamName = Optimize-DisplayName -String (Get-Team -GroupId $MatchingCQ.DistributionLists.Guid).DisplayName
            $ChannelName = Optimize-DisplayName -String (Get-TeamChannel -GroupId $MatchingCQ.DistributionLists.Guid | Where-Object {$_.Id -eq $MatchingCQ.ChannelId}).DisplayName

            $CqAgentListType = "Teams Channel <br> Team Name: $TeamName <br> Channel Name: $ChannelName"

        }

    }

    # Switch through call queue overflow action target
    switch ($CqOverFlowAction) {
        DisconnectWithBusy {

            if ($CombineDisconnectCallNodes -eq $true) {

                $CqOverFlowActionFriendly = "disconnectCall((DisconnectCall))"

                $allMermaidNodes += "disconnectCall"
    
            }
    
            else {
    
                $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)((DisconnectCall))"

                $allMermaidNodes += "cqOverFlowAction$($cqCallFlowObjectId)"
            
            }

            if ($MatchingCQ.OverflowDisconnectAudioFilePrompt -or $MatchingCQ.OverflowDisconnectTextToSpeechPrompt) {

                if ($MatchingCQ.OverflowDisconnectAudioFilePrompt) {

                    if ($ShowAudioFileName) {

                        $audioFileName = Optimize-DisplayName -String ($MatchingCQ.OverflowDisconnectAudioFilePromptFileName)

                        # If audio file name is not present on call queue properties
                        if (!$audioFileName) {

                            $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.OverflowDisconnectAudioFilePrompt -ApplicationId HuntGroup).FileName

                        }

                        if ($ExportAudioFiles) {

                            $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.OverflowDisconnectAudioFilePrompt -ApplicationId HuntGroup
                            [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                            $audioFileNames += ("click cqOverflowDisconnectAudioFilePrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                        }
        
                    
                        if ($audioFileName.Length -gt $truncateGreetings) {
                
                            $audioFileNameExtension = ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[3]
                            $audioFileName = $audioFileName.Remove($audioFileName.Length - ($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                        }
        
                        $CqOverFlowActionFriendly = "cqOverflowDisconnectAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile <br> $audioFileName] --> $cqOverFlowActionFriendly"
        
                    }

                    else {

                        $CqOverFlowActionFriendly = "cqOverflowDisconnectAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile] --> $cqOverFlowActionFriendly"

                    }

                    $allMermaidNodes += "cqOverflowDisconnectAudioFilePrompt$($cqCallFlowObjectId)"

                }

                else {

                    if ($ShowTTSGreetingText) {

                        $overflowDisconnectTextToSpeechPromptExport = $MatchingCQ.OverflowDisconnectTextToSpeechPrompt
                        $overflowDisconnectTextToSpeechPromptValue = Optimize-DisplayName -String $MatchingCQ.OverflowDisconnectTextToSpeechPrompt

                        if ($ExportTTSGreetings) {

                            $overflowDisconnectTextToSpeechPromptExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqOverFlowDisconnectGreeting.txt"
        
                            $ttsGreetings += ("click cqOverflowDisconnectTextToSpeechPrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqOverFlowDisconnectGreeting.txt" + '"')
        
                        }    

                        if ($overflowDisconnectTextToSpeechPromptValue.Length -gt $truncateGreetings) {

                            $overflowDisconnectTextToSpeechPromptValue = $overflowDisconnectTextToSpeechPromptValue.Remove($overflowDisconnectTextToSpeechPromptValue.Length - ($overflowDisconnectTextToSpeechPromptValue.Length - $truncateGreetings)).TrimEnd() + "..."

                        }

                        $CqOverFlowActionFriendly = "cqOverflowDisconnectTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech <br> $overflowDisconnectTextToSpeechPromptValue] --> $cqOverFlowActionFriendly"

                    }

                    else {

                        $CqOverFlowActionFriendly = "cqOverflowDisconnectTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech] --> $cqOverFlowActionFriendly"

                    }

                    $allMermaidNodes += "cqOverflowDisconnectTextToSpeechPrompt$($cqCallFlowObjectId)"

                }

            }

        }
        Forward {

            if ($MatchingCQ.OverflowActionTarget.Type -eq "User") {

                $MatchingOverFlowUserProperties = (Get-MgUser -UserId $MatchingCQ.OverflowActionTarget.Id)
                $MatchingOverFlowUser = Optimize-DisplayName -String $MatchingOverFlowUserProperties.DisplayName
                $MatchingOverFlowIdentity = $MatchingOverFlowUserProperties.Id

                if ($FindUserLinks -eq $true) {
         
                    . New-VoiceAppUserLinkProperties -userLinkUserId $MatchingCQ.OverflowActionTarget.Id -userLinkUserName $MatchingOverFlowUserProperties.DisplayName -userLinkVoiceAppType "Call Queue" -userLinkVoiceAppActionType "OverflowActionTarget" -userLinkVoiceAppName $MatchingCQ.Name -userLinkVoiceAppId $MatchingCQIdentity
                
                }        

                if ($nestedVoiceApps -notcontains $MatchingOverFlowUserProperties.Id -and $MatchingCQ.TimeoutThreshold -ge 1) {

                    $nestedVoiceApps += $MatchingOverFlowUserProperties.Id

                }

                $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingOverFlowIdentity)(User <br> $MatchingOverFlowUser)"

                $allMermaidNodes += @("cqOverFlowAction$($cqCallFlowObjectId)","$($MatchingOverFlowIdentity)")

            }

            elseif ($MatchingCQ.OverflowActionTarget.Type -eq "Phone") {

                $cqOverFlowPhoneNumber = ($MatchingCQ.OverflowActionTarget.Id).Replace("tel:","")

                if ($ObfuscatePhoneNumbers -eq $true) {

                    $cqOverFlowPhoneNumber = $cqOverFlowPhoneNumber.Remove(($cqOverFlowPhoneNumber.Length -4)) + "****"

                }        

                $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($cqOverFlowPhoneNumber)(External Number <br> $cqOverFlowPhoneNumber)"

                $allMermaidNodes += @("cqOverFlowAction$($cqCallFlowObjectId)","$($cqOverFlowPhoneNumber)")

                if ($MatchingCQ.OverflowRedirectPhoneNumberAudioFilePrompt -or $MatchingCQ.OverflowRedirectPhoneNumberTextToSpeechPrompt) {

                    if ($MatchingCQ.OverflowRedirectPhoneNumberAudioFilePrompt) {

                        if ($ShowAudioFileName) {

                            $audioFileName = Optimize-DisplayName -String ($MatchingCQ.OverflowRedirectPhoneNumberAudioFilePromptFileName)

                            # If audio file name is not present on call queue properties
                            if (!$audioFileName) {

                                $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.OverflowRedirectPhoneNumberAudioFilePrompt -ApplicationId HuntGroup).FileName

                            }

                            if ($ExportAudioFiles) {

                                $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.OverflowRedirectPhoneNumberAudioFilePrompt -ApplicationId HuntGroup
                                [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                                $audioFileNames += ("click cqOverflowRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                            }
        
                    
                            if ($audioFileName.Length -gt $truncateGreetings) {
                
                                $audioFileNameExtension = ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[3]
                                $audioFileName = $audioFileName.Remove($audioFileName.Length - ($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                            }
        
                            $CqOverFlowActionFriendly = "cqOverflowRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile <br> $audioFileName] --> $CqOverFlowActionFriendly"
        
                        }

                        else {

                            $CqOverFlowActionFriendly = "cqOverflowRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile] --> $CqOverFlowActionFriendly"

                        }

                        $allMermaidNodes += "cqOverflowRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId)"

                    }

                    else {

                        if ($ShowTTSGreetingText) {

                            $OverflowRedirectPhoneNumberTextToSpeechPromptExport = $MatchingCQ.OverflowRedirectPhoneNumberTextToSpeechPrompt
                            $OverflowRedirectPhoneNumberTextToSpeechPromptValue = Optimize-DisplayName -String $MatchingCQ.OverflowRedirectPhoneNumberTextToSpeechPrompt

                            if ($ExportTTSGreetings) {

                                $OverflowRedirectPhoneNumberTextToSpeechPromptExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqOverflowRedirectPhoneNumberGreeting.txt"
        
                                $ttsGreetings += ("click cqOverflowRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqOverflowRedirectPhoneNumberGreeting.txt" + '"')
        
                            }    

                            if ($OverflowRedirectPhoneNumberTextToSpeechPromptValue.Length -gt $truncateGreetings) {

                                $OverflowRedirectPhoneNumberTextToSpeechPromptValue = $OverflowRedirectPhoneNumberTextToSpeechPromptValue.Remove($OverflowRedirectPhoneNumberTextToSpeechPromptValue.Length - ($OverflowRedirectPhoneNumberTextToSpeechPromptValue.Length - $truncateGreetings)).TrimEnd() + "..."

                            }

                            $CqOverFlowActionFriendly = "cqOverflowRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech <br> $OverflowRedirectPhoneNumberTextToSpeechPromptValue] --> $CqOverFlowActionFriendly"

                        }

                        else {

                            $CqOverFlowActionFriendly = "cqOverflowRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech] --> $CqOverFlowActionFriendly"

                        }

                        $allMermaidNodes += "cqOverflowRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId)"

                    }

                }
                
            }

            else {

                if ($ShowTransferCallToTargetType -eq $true) {

                    switch ($MatchingCQ.OverflowActionTarget.Type) {
                        ApplicationEndpoint {

                            $cqTransferCallToTargetTypeAdditionalInfo = " <br> Resource Account"

                        }
                        ConfigurationEndpoint {

                            $cqTransferCallToTargetTypeAdditionalInfo = " <br> Voice App"

                        }
                        Default {

                            $cqTransferCallToTargetTypeAdditionalInfo = $null

                        }
                    }

                }

                else {

                    $cqTransferCallToTargetTypeAdditionalInfo = $null

                }

                $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $MatchingCQ.OverflowActionTarget.Id -and $_.ApplicationId -eq $applicationIdAa}

                if ($matchingApplicationInstanceCheckAa -or $allAutoAttendantIds -contains $MatchingCQ.OverflowActionTarget.Id) {

                    $MatchingOverFlowAA = ($allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.OverflowActionTarget.Id -or $_.Identity -eq $MatchingCQ.OverflowActionTarget.Id})

                    $MatchingOverFlowAA.Name = Optimize-DisplayName -String $MatchingOverFlowAA.Name
                    $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget$cqTransferCallToTargetTypeAdditionalInfo) --> $($MatchingOverFlowAA.Identity)([Auto Attendant <br> $($MatchingOverFlowAA.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingOverFlowAA.Identity -and $MatchingCQ.TimeoutThreshold -ge 1) {

                        $nestedVoiceApps += $MatchingOverFlowAA.Identity
        
                    }

                    $allMermaidNodes += @("cqOverFlowAction$($cqCallFlowObjectId)","$($MatchingOverFlowAA.Identity)")
        

                }

                else {

                    $MatchingOverFlowCQ = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.OverflowActionTarget.Id -or $_.Identity -eq $MatchingCQ.OverflowActionTarget.Id})

                    $MatchingOverFlowCQ.Name = Optimize-DisplayName -String $MatchingOverFlowCQ.Name

                    $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget$cqTransferCallToTargetTypeAdditionalInfo) --> $($MatchingOverFlowCQ.Identity)([Call Queue <br> $($MatchingOverFlowCQ.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingOverFlowCQ.Identity -and $MatchingCQ.TimeoutThreshold -ge 1) {

                        $nestedVoiceApps += $MatchingOverFlowCQ.Identity
        
                    }

                    $allMermaidNodes += @("cqOverFlowAction$($cqCallFlowObjectId)","$($MatchingOverFlowCQ.Identity)")

                }

                if ($MatchingCQ.OverflowRedirectVoiceAppAudioFilePrompt -or $MatchingCQ.OverflowRedirectVoiceAppTextToSpeechPrompt) {

                    if ($MatchingCQ.OverflowRedirectVoiceAppAudioFilePrompt) {

                        if ($ShowAudioFileName) {

                            $audioFileName = Optimize-DisplayName -String ($MatchingCQ.OverflowRedirectVoiceAppAudioFilePromptFileName)

                            # If audio file name is not present on call queue properties
                            if (!$audioFileName) {

                                $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.OverflowRedirectVoiceAppAudioFilePrompt -ApplicationId HuntGroup).FileName

                            }

                            if ($ExportAudioFiles) {

                                $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.OverflowRedirectVoiceAppAudioFilePrompt -ApplicationId HuntGroup
                                [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                                $audioFileNames += ("click cqOverflowRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                            }
        
                    
                            if ($audioFileName.Length -gt $truncateGreetings) {
                
                                $audioFileNameExtension = ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[3]
                                $audioFileName = $audioFileName.Remove($audioFileName.Length - ($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                            }
        
                            $CqOverFlowActionFriendly = "cqOverflowRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile <br> $audioFileName] --> $CqOverFlowActionFriendly"
        
                        }

                        else {

                            $CqOverFlowActionFriendly = "cqOverflowRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile] --> $CqOverFlowActionFriendly"

                        }

                        $allMermaidNodes += "cqOverflowRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId)"

                    }

                    else {

                        if ($ShowTTSGreetingText) {

                            $OverflowRedirectVoiceAppTextToSpeechPromptExport = $MatchingCQ.OverflowRedirectVoiceAppTextToSpeechPrompt
                            $OverflowRedirectVoiceAppTextToSpeechPromptValue = Optimize-DisplayName -String $MatchingCQ.OverflowRedirectVoiceAppTextToSpeechPrompt

                            if ($ExportTTSGreetings) {

                                $OverflowRedirectVoiceAppTextToSpeechPromptExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqOverflowRedirectVoiceAppGreeting.txt"
        
                                $ttsGreetings += ("click cqOverflowRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqOverflowRedirectVoiceAppGreeting.txt" + '"')
        
                            }    

                            if ($OverflowRedirectVoiceAppTextToSpeechPromptValue.Length -gt $truncateGreetings) {

                                $OverflowRedirectVoiceAppTextToSpeechPromptValue = $OverflowRedirectVoiceAppTextToSpeechPromptValue.Remove($OverflowRedirectVoiceAppTextToSpeechPromptValue.Length - ($OverflowRedirectVoiceAppTextToSpeechPromptValue.Length - $truncateGreetings)).TrimEnd() + "..."

                            }

                            $CqOverFlowActionFriendly = "cqOverflowRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech <br> $OverflowRedirectVoiceAppTextToSpeechPromptValue] --> $CqOverFlowActionFriendly"

                        }

                        else {

                            $CqOverFlowActionFriendly = "cqOverflowRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech] --> $CqOverFlowActionFriendly"

                        }

                        $allMermaidNodes += "cqOverflowRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId)"

                    }

                }

            }

        }
        Voicemail {
            $MatchingOverFlowPersonalVoicemailUserProperties = (Get-MgUser -UserId $MatchingCQ.OverflowActionTarget.Id)
            $MatchingOverFlowPersonalVoicemailUser = Optimize-DisplayName -String $MatchingOverFlowPersonalVoicemailUserProperties.DisplayName
            $MatchingOverFlowPersonalVoicemailIdentity = $MatchingOverFlowPersonalVoicemailUserProperties.Id

            if ($FindUserLinks -eq $true) {
         
                . New-VoiceAppUserLinkProperties -userLinkUserId $MatchingCQ.OverflowActionTarget.Id -userLinkUserName $MatchingOverFlowPersonalVoicemailUserProperties.DisplayName -userLinkVoiceAppType "Call Queue" -userLinkVoiceAppActionType "OverflowActionTargetPersonalVoicemail" -userLinkVoiceAppName $MatchingCQ.Name -userLinkVoiceAppId $MatchingCQIdentity
            
            }        

            $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget) --> cqPersonalVoicemail$($MatchingOverFlowPersonalVoicemailIdentity)(Personal Voicemail <br> $MatchingOverFlowPersonalVoicemailUser)"

            $allMermaidNodes += @("cqOverFlowAction$($cqCallFlowObjectId)","$($MatchingOverFlowPersonalVoicemailIdentity)","cqPersonalVoicemail$($MatchingOverFlowPersonalVoicemailIdentity)")
        }
        SharedVoicemail {
            $MatchingOverFlowVoicemailProperties = (Get-MgGroup -GroupId $MatchingCQ.OverflowActionTarget.Id)
            $MatchingOverFlowVoicemail = Optimize-DisplayName -String $MatchingOverFlowVoicemailProperties.DisplayName
            $MatchingOverFlowIdentity = $MatchingOverFlowVoicemailProperties.Id

            if ($ShowSharedVoicemailGroupMembers -eq $true) {

                . Get-SharedVoicemailGroupMembers -SharedVoicemailGroupId $MatchingCQ.OverflowActionTarget.Id

                $MatchingOverFlowVoicemail = "$MatchingOverFlowVoicemail$mdSharedVoicemailGroupMembers"

            }

            if ($MatchingCQ.OverflowSharedVoicemailTextToSpeechPrompt) {

                $CqOverFlowVoicemailGreeting = "TextToSpeech"

                if ($ShowTTSGreetingText) {

                    $overFlowVoicemailTTSGreetingValueExport = $MatchingCQ.OverflowSharedVoicemailTextToSpeechPrompt
                    $overFlowVoicemailTTSGreetingValue = Optimize-DisplayName -String $MatchingCQ.OverflowSharedVoicemailTextToSpeechPrompt

                    if ($ExportTTSGreetings) {

                        $overFlowVoicemailTTSGreetingValueExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqOverFlowVoicemailGreeting.txt"
        
                        $ttsGreetings += ("click cqOverFlowVoicemailGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqOverFlowVoicemailGreeting.txt" + '"')
        
                    }

                    if ($overFlowVoicemailTTSGreetingValue.Length -gt $truncateGreetings) {

                        $overFlowVoicemailTTSGreetingValue = $overFlowVoicemailTTSGreetingValue.Remove($overFlowVoicemailTTSGreetingValue.Length - ($overFlowVoicemailTTSGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."

                    }

                    $CqOverFlowVoicemailGreeting += " <br> ''$overFlowVoicemailTTSGreetingValue''"

                }

                if ($MatchingCQ.EnableOverflowSharedVoicemailSystemPromptSuppression -eq $false) {

                    $CQOverFlowVoicemailSystemGreeting = "--> cqOverFlowVoicemailSystemGreeting$($cqCallFlowObjectId)>Greeting <br> MS System Message] "

                    $CQOverFlowVoicemailSystemGreetingValue = (. Get-MsSystemMessage)[-1]
                    $CQOverFlowVoicemailSystemGreetingValueExport = (. Get-MsSystemMessage)[0]


                    if ($ShowTTSGreetingText) {

                        if ($ExportTTSGreetings) {

                            $CQOverFlowVoicemailSystemGreetingValueExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqOverFlowMsSystemMessage.txt"
            
                            $ttsGreetings += ("click cqOverFlowVoicemailSystemGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqOverFlowMsSystemMessage.txt" + '"')
            
                        }    
        

                        if ($CQOverFlowVoicemailSystemGreetingValue.Length -gt $truncateGreetings) {

                            $CQOverFlowVoicemailSystemGreetingValue = $CQOverFlowVoicemailSystemGreetingValue.Remove($CQOverFlowVoicemailSystemGreetingValue.Length - ($CQOverFlowVoicemailSystemGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."

                        }

                        $CQOverFlowVoicemailSystemGreeting = $CQOverFlowVoicemailSystemGreeting.Replace("] "," <br> ''$CQOverFlowVoicemailSystemGreetingValue''] ")

                    }

                    $allMermaidNodes += "cqOverFlowVoicemailSystemGreeting$($cqCallFlowObjectId)"

                }

                else {

                    $CQOverFlowVoicemailSystemGreeting = $null

                }

            }

            else {

                $CqOverFlowVoicemailGreeting = "AudioFile"

                if ($ShowAudioFileName) {

                    $audioFileName = Optimize-DisplayName -String ($MatchingCQ.OverflowSharedVoicemailAudioFilePromptFileName)

                    # If audio file name is not present on call queue properties
                    if (!$audioFileName) {

                        $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.OverflowSharedVoicemailAudioFilePrompt -ApplicationId HuntGroup).FileName

                    }

                    if ($ExportAudioFiles) {

                        $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.OverflowSharedVoicemailAudioFilePrompt -ApplicationId HuntGroup
                        [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                        $audioFileNames += ("click cqOverFlowVoicemailGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                    }
        
                    
                    if ($audioFileName.Length -gt $truncateGreetings) {
                
                        $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                        $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                    }
        
                    $CqOverFlowVoicemailGreeting += " <br> $audioFileName"
        
                }        

                if ($MatchingCQ.EnableOverflowSharedVoicemailSystemPromptSuppression -eq $false) {

                    $CQOverFlowVoicemailSystemGreeting = "--> cqOverFlowVoicemailSystemGreeting$($cqCallFlowObjectId)>Greeting <br> MS System Message] "

                    $CQOverFlowVoicemailSystemGreetingValue = (. Get-MsSystemMessage)[-1]
                    $CQOverFlowVoicemailSystemGreetingValueExport = (. Get-MsSystemMessage)[0]

                    if ($ShowTTSGreetingText) {

                        if ($ExportTTSGreetings) {

                            $CQOverFlowVoicemailSystemGreetingValueExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqOverFlowMsSystemMessage.txt"
            
                            $ttsGreetings += ("click cqOverFlowVoicemailSystemGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqOverFlowMsSystemMessage.txt" + '"')
            
                        }    
        

                        if ($CQOverFlowVoicemailSystemGreetingValue.Length -gt $truncateGreetings) {

                            $CQOverFlowVoicemailSystemGreetingValue = $CQOverFlowVoicemailSystemGreetingValue.Remove($CQOverFlowVoicemailSystemGreetingValue.Length - ($CQOverFlowVoicemailSystemGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."

                        }

                        $CQOverFlowVoicemailSystemGreeting = $CQOverFlowVoicemailSystemGreeting.Replace("] "," <br> ''$CQOverFlowVoicemailSystemGreetingValue''] ")

                    }

                    $allMermaidNodes += "cqOverFlowVoicemailSystemGreeting$($cqCallFlowObjectId)"

                }

                else {

                    $CQOverFlowVoicemailSystemGreeting = $null

                }

            }

            $CqOverFlowActionFriendly = "cqOverFlowVoicemailGreeting$($cqCallFlowObjectId)>Greeting <br> $CqOverFlowVoicemailGreeting] $CQOverFlowVoicemailSystemGreeting--> cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingOverFlowIdentity)(Shared Voicemail <br> $MatchingOverFlowVoicemail)"

            $allMermaidNodes += @("cqOverFlowAction$($cqCallFlowObjectId)","cqOverFlowVoicemailGreeting$($cqCallFlowObjectId)","$($MatchingOverFlowIdentity)")

        }

    }

    # Switch through call queue timeout overflow action
    switch ($CqTimeoutAction) {
        Disconnect {

            if ($CombineDisconnectCallNodes -eq $true) {

                $CqTimeoutActionFriendly = "disconnectCall((DisconnectCall))"

                $allMermaidNodes += "disconnectCall"
    
            }
    
            else {
    
                $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)((DisconnectCall))"

                $allMermaidNodes += "cqTimeoutAction$($cqCallFlowObjectId)"
                
            }

            if ($MatchingCQ.TimeoutDisconnectAudioFilePrompt -or $MatchingCQ.TimeoutDisconnectTextToSpeechPrompt) {

                if ($MatchingCQ.TimeoutDisconnectAudioFilePrompt) {

                    if ($ShowAudioFileName) {

                        $audioFileName = Optimize-DisplayName -String ($MatchingCQ.TimeoutDisconnectAudioFilePromptFileName)

                        # If audio file name is not present on call queue properties
                        if (!$audioFileName) {

                            $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.TimeoutDisconnectAudioFilePrompt -ApplicationId HuntGroup).FileName

                        }

                        if ($ExportAudioFiles) {

                            $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.TimeoutDisconnectAudioFilePrompt -ApplicationId HuntGroup
                            [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                            $audioFileNames += ("click cqTimeoutDisconnectAudioFilePrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                        }
        
                    
                        if ($audioFileName.Length -gt $truncateGreetings) {
                
                            $audioFileNameExtension = ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[3]
                            $audioFileName = $audioFileName.Remove($audioFileName.Length - ($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                        }
        
                        $CqTimeoutActionFriendly = "cqTimeoutDisconnectAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile <br> $audioFileName] --> $cqTimeoutActionFriendly"
        
                    }

                    else {

                        $CqTimeoutActionFriendly = "cqTimeoutDisconnectAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile] --> $cqTimeoutActionFriendly"

                    }

                    $allMermaidNodes += "cqTimeoutDisconnectAudioFilePrompt$($cqCallFlowObjectId)"

                }

                else {

                    if ($ShowTTSGreetingText) {

                        $TimeoutDisconnectTextToSpeechPromptExport = $MatchingCQ.TimeoutDisconnectTextToSpeechPrompt
                        $TimeoutDisconnectTextToSpeechPromptValue = Optimize-DisplayName -String $MatchingCQ.TimeoutDisconnectTextToSpeechPrompt

                        if ($ExportTTSGreetings) {

                            $TimeoutDisconnectTextToSpeechPromptExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqTimeoutDisconnectGreeting.txt"
        
                            $ttsGreetings += ("click cqTimeoutDisconnectTextToSpeechPrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqTimeoutDisconnectGreeting.txt" + '"')
        
                        }    

                        if ($TimeoutDisconnectTextToSpeechPromptValue.Length -gt $truncateGreetings) {

                            $TimeoutDisconnectTextToSpeechPromptValue = $TimeoutDisconnectTextToSpeechPromptValue.Remove($TimeoutDisconnectTextToSpeechPromptValue.Length - ($TimeoutDisconnectTextToSpeechPromptValue.Length - $truncateGreetings)).TrimEnd() + "..."

                        }

                        $CqTimeoutActionFriendly = "cqTimeoutDisconnectTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech <br> $TimeoutDisconnectTextToSpeechPromptValue] --> $cqTimeoutActionFriendly"

                    }

                    else {

                        $CqTimeoutActionFriendly = "cqTimeoutDisconnectTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech] --> $cqTimeoutActionFriendly"

                    }

                    $allMermaidNodes += "cqTimeoutDisconnectTextToSpeechPrompt$($cqCallFlowObjectId)"

                }

            }

        }
        Forward {
    
            if ($MatchingCQ.TimeoutActionTarget.Type -eq "User") {

                $MatchingTimeoutUserProperties = (Get-MgUser -UserId $MatchingCQ.TimeoutActionTarget.Id)
                $MatchingTimeoutUser = Optimize-DisplayName -String $MatchingTimeoutUserProperties.DisplayName
                $MatchingTimeoutIdentity = $MatchingTimeoutUserProperties.Id

                if ($FindUserLinks -eq $true) {
         
                    . New-VoiceAppUserLinkProperties -userLinkUserId $MatchingCQ.TimeoutActionTarget.Id -userLinkUserName $MatchingTimeoutUserProperties.DisplayName -userLinkVoiceAppType "Call Queue" -userLinkVoiceAppActionType "TimoutActionTarget" -userLinkVoiceAppName $MatchingCQ.Name -userLinkVoiceAppId $MatchingCQIdentity
                
                }        

                if ($nestedVoiceApps -notcontains $MatchingTimeoutUserProperties.Id -and $MatchingCQ.OverflowThreshold -ge 1) {

                    $nestedVoiceApps += $MatchingTimeoutUserProperties.Id

                }
    
                $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingTimeoutIdentity)(User <br> $MatchingTimeoutUser)"

                $allMermaidNodes += @("cqTimeoutAction$($cqCallFlowObjectId)","$($MatchingTimeoutIdentity)")
    
            }
    
            elseif ($MatchingCQ.TimeoutActionTarget.Type -eq "Phone") {
    
                $cqTimeoutPhoneNumber = ($MatchingCQ.TimeoutActionTarget.Id).Replace("tel:","")

                if ($ObfuscatePhoneNumbers -eq $true) {

                    $cqTimeoutPhoneNumber = $cqTimeoutPhoneNumber.Remove(($cqTimeoutPhoneNumber.Length -4)) + "****"

                }
    
                $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($cqTimeoutPhoneNumber)(External Number <br> $cqTimeoutPhoneNumber)"

                $allMermaidNodes += @("cqTimeoutAction$($cqCallFlowObjectId)","$($cqTimeoutPhoneNumber)")

                if ($MatchingCQ.TimeoutRedirectPhoneNumberAudioFilePrompt -or $MatchingCQ.TimeoutRedirectPhoneNumberTextToSpeechPrompt) {

                    if ($MatchingCQ.TimeoutRedirectPhoneNumberAudioFilePrompt) {

                        if ($ShowAudioFileName) {

                            $audioFileName = Optimize-DisplayName -String ($MatchingCQ.TimeoutRedirectPhoneNumberAudioFilePromptFileName)

                            # If audio file name is not present on call queue properties
                            if (!$audioFileName) {

                                $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.TimeoutRedirectPhoneNumberAudioFilePrompt -ApplicationId HuntGroup).FileName

                            }

                            if ($ExportAudioFiles) {

                                $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.TimeoutRedirectPhoneNumberAudioFilePrompt -ApplicationId HuntGroup
                                [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                                $audioFileNames += ("click cqTimeoutRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                            }
        
                    
                            if ($audioFileName.Length -gt $truncateGreetings) {
                
                                $audioFileNameExtension = ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[3]
                                $audioFileName = $audioFileName.Remove($audioFileName.Length - ($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                            }
        
                            $CqTimeoutActionFriendly = "cqTimeoutRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile <br> $audioFileName] --> $CqTimeoutActionFriendly"
        
                        }

                        else {

                            $CqTimeoutActionFriendly = "cqTimeoutRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile] --> $CqTimeoutActionFriendly"

                        }

                        $allMermaidNodes += "cqTimeoutRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId)"

                    }

                    else {

                        if ($ShowTTSGreetingText) {

                            $TimeoutRedirectPhoneNumberTextToSpeechPromptExport = $MatchingCQ.TimeoutRedirectPhoneNumberTextToSpeechPrompt
                            $TimeoutRedirectPhoneNumberTextToSpeechPromptValue = Optimize-DisplayName -String $MatchingCQ.TimeoutRedirectPhoneNumberTextToSpeechPrompt

                            if ($ExportTTSGreetings) {

                                $TimeoutRedirectPhoneNumberTextToSpeechPromptExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqTimeoutRedirectPhoneNumberGreeting.txt"
        
                                $ttsGreetings += ("click cqTimeoutRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqTimeoutRedirectPhoneNumberGreeting.txt" + '"')
        
                            }    

                            if ($TimeoutRedirectPhoneNumberTextToSpeechPromptValue.Length -gt $truncateGreetings) {

                                $TimeoutRedirectPhoneNumberTextToSpeechPromptValue = $TimeoutRedirectPhoneNumberTextToSpeechPromptValue.Remove($TimeoutRedirectPhoneNumberTextToSpeechPromptValue.Length - ($TimeoutRedirectPhoneNumberTextToSpeechPromptValue.Length - $truncateGreetings)).TrimEnd() + "..."

                            }

                            $CqTimeoutActionFriendly = "cqTimeoutRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech <br> $TimeoutRedirectPhoneNumberTextToSpeechPromptValue] --> $CqTimeoutActionFriendly"

                        }

                        else {

                            $CqTimeoutActionFriendly = "cqTimeoutRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech] --> $CqTimeoutActionFriendly"

                        }

                        $allMermaidNodes += "cqTimeoutRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId)"

                    }

                }
                
            }
    
            else {

                if ($ShowTransferCallToTargetType -eq $true) {

                    switch ($MatchingCQ.TimeoutActionTarget.Type) {
                        ApplicationEndpoint {

                            $cqTransferCallToTargetTypeAdditionalInfo = " <br> Resource Account"

                        }
                        ConfigurationEndpoint {

                            $cqTransferCallToTargetTypeAdditionalInfo = " <br> Voice App"

                        }
                        Default {

                            $cqTransferCallToTargetTypeAdditionalInfo = $null

                        }
                    }

                }

                else {

                    $cqTransferCallToTargetTypeAdditionalInfo = $null

                }

                $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $MatchingCQ.TimeoutActionTarget.Id -and $_.ApplicationId -eq $applicationIdAa}
        
                if ($matchingApplicationInstanceCheckAa -or $allAutoAttendantIds -contains $MatchingCQ.TimeoutActionTarget.Id) {

                    $MatchingTimeoutAA = ($allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.TimeoutActionTarget.Id -or $_.Identity -eq $MatchingCQ.TimeoutActionTarget.Id})

                    $MatchingTimeoutAA.Name = Optimize-DisplayName -String $MatchingTimeoutAA.Name
    
                    $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget$cqTransferCallToTargetTypeAdditionalInfo) --> $($MatchingTimeoutAA.Identity)([Auto Attendant <br> $($MatchingTimeoutAA.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingTimeoutAA.Identity -and $MatchingCQ.OverflowThreshold -ge 1) {

                        $nestedVoiceApps += $MatchingTimeoutAA.Identity
        
                    }

                    $allMermaidNodes += @("cqTimeoutAction$($cqCallFlowObjectId)","$($MatchingTimeoutAA.Identity)")
    
                }
    
                else {
    
                    $MatchingTimeoutCQ = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.TimeoutActionTarget.Id -or $_.Identity -eq $MatchingCQ.TimeoutActionTarget.Id})

                    $MatchingTimeoutCQ.Name = Optimize-DisplayName -String $MatchingTimeoutCQ.Name

                    $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget$cqTransferCallToTargetTypeAdditionalInfo) --> $($MatchingTimeoutCQ.Identity)([Call Queue <br> $($MatchingTimeoutCQ.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingTimeoutCQ.Identity -and $MatchingCQ.OverflowThreshold -ge 1) {

                        $nestedVoiceApps += $MatchingTimeoutCQ.Identity
        
                    }

                    $allMermaidNodes += @("cqTimeoutAction$($cqCallFlowObjectId)","$($MatchingTimeoutCQ.Identity)")
    
                }

                if ($MatchingCQ.TimeoutRedirectVoiceAppAudioFilePrompt -or $MatchingCQ.TimeoutRedirectVoiceAppTextToSpeechPrompt) {

                    if ($MatchingCQ.TimeoutRedirectVoiceAppAudioFilePrompt) {

                        if ($ShowAudioFileName) {

                            $audioFileName = Optimize-DisplayName -String ($MatchingCQ.TimeoutRedirectVoiceAppAudioFilePromptFileName)

                            # If audio file name is not present on call queue properties
                            if (!$audioFileName) {

                                $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.TimeoutRedirectVoiceAppAudioFilePrompt -ApplicationId HuntGroup).FileName

                            }

                            if ($ExportAudioFiles) {

                                $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.TimeoutRedirectVoiceAppAudioFilePrompt -ApplicationId HuntGroup
                                [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                                $audioFileNames += ("click cqTimeoutRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                            }
        
                    
                            if ($audioFileName.Length -gt $truncateGreetings) {
                
                                $audioFileNameExtension = ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[3]
                                $audioFileName = $audioFileName.Remove($audioFileName.Length - ($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                            }
        
                            $CqTimeoutActionFriendly = "cqTimeoutRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile <br> $audioFileName] --> $CqTimeoutActionFriendly"
        
                        }

                        else {

                            $CqTimeoutActionFriendly = "cqTimeoutRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile] --> $CqTimeoutActionFriendly"

                        }

                        $allMermaidNodes += "cqTimeoutRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId)"

                    }

                    else {

                        if ($ShowTTSGreetingText) {

                            $TimeoutRedirectVoiceAppTextToSpeechPromptExport = $MatchingCQ.TimeoutRedirectVoiceAppTextToSpeechPrompt
                            $TimeoutRedirectVoiceAppTextToSpeechPromptValue = Optimize-DisplayName -String $MatchingCQ.TimeoutRedirectVoiceAppTextToSpeechPrompt

                            if ($ExportTTSGreetings) {

                                $TimeoutRedirectVoiceAppTextToSpeechPromptExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqTimeoutRedirectVoiceAppGreeting.txt"
        
                                $ttsGreetings += ("click cqTimeoutRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqTimeoutRedirectVoiceAppGreeting.txt" + '"')
        
                            }    

                            if ($TimeoutRedirectVoiceAppTextToSpeechPromptValue.Length -gt $truncateGreetings) {

                                $TimeoutRedirectVoiceAppTextToSpeechPromptValue = $TimeoutRedirectVoiceAppTextToSpeechPromptValue.Remove($TimeoutRedirectVoiceAppTextToSpeechPromptValue.Length - ($TimeoutRedirectVoiceAppTextToSpeechPromptValue.Length - $truncateGreetings)).TrimEnd() + "..."

                            }

                            $CqTimeoutActionFriendly = "cqTimeoutRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech <br> $TimeoutRedirectVoiceAppTextToSpeechPromptValue] --> $CqTimeoutActionFriendly"

                        }

                        else {

                            $CqTimeoutActionFriendly = "cqTimeoutRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech] --> $CqTimeoutActionFriendly"

                        }

                        $allMermaidNodes += "cqTimeoutRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId)"

                    }

                }

            }
    
        }
        Voicemail {
            $MatchingTimeoutPersonalVoicemailUserProperties = (Get-MgUser -UserId $MatchingCQ.TimeoutActionTarget.Id)
            $MatchingTimeoutPersonalVoicemailUser = Optimize-DisplayName -String $MatchingTimeoutPersonalVoicemailUserProperties.DisplayName
            $MatchingTimeoutPersonalVoicemailIdentity = $MatchingTimeoutPersonalVoicemailUserProperties.Id

            if ($FindUserLinks -eq $true) {
         
                . New-VoiceAppUserLinkProperties -userLinkUserId $MatchingCQ.TimeoutActionTarget.Id -userLinkUserName $MatchingTimeoutPersonalVoicemailUserProperties.DisplayName -userLinkVoiceAppType "Call Queue" -userLinkVoiceAppActionType "TimoutActionTargetPersonalVoicemail" -userLinkVoiceAppName $MatchingCQ.Name -userLinkVoiceAppId $MatchingCQIdentity
            
            }        

            $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget) --> cqPersonalVoicemail$($MatchingTimeoutPersonalVoicemailIdentity)(Personal Voicemail <br> $MatchingTimeoutPersonalVoicemailUser)"

            $allMermaidNodes += @("cqTimeoutAction$($cqCallFlowObjectId)","$($MatchingTimeoutPersonalVoicemailIdentity)","cqPersonalVoicemail$($MatchingTimeoutPersonalVoicemailIdentity)")
        }
        SharedVoicemail {
            $MatchingTimeoutVoicemailProperties = (Get-MgGroup -GroupId $MatchingCQ.TimeoutActionTarget.Id)
            $MatchingTimeoutVoicemail = Optimize-DisplayName -String $MatchingTimeoutVoicemailProperties.DisplayName
            $MatchingTimeoutIdentity = $MatchingTimeoutVoicemailProperties.Id

            if ($ShowSharedVoicemailGroupMembers -eq $true) {

                . Get-SharedVoicemailGroupMembers -SharedVoicemailGroupId $MatchingCQ.TimeoutActionTarget.Id

                $MatchingTimeoutVoicemail = "$MatchingTimeoutVoicemail$mdSharedVoicemailGroupMembers"

            }
    
            if ($MatchingCQ.TimeoutSharedVoicemailTextToSpeechPrompt) {
    
                $CqTimeoutVoicemailGreeting = "TextToSpeech"
                
                if ($ShowTTSGreetingText) {

                    $TimeOutVoicemailTTSGreetingValueExport = $MatchingCQ.TimeOutSharedVoicemailTextToSpeechPrompt
                    $TimeOutVoicemailTTSGreetingValue = Optimize-DisplayName -String $MatchingCQ.TimeOutSharedVoicemailTextToSpeechPrompt

                    if ($ExportTTSGreetings) {

                        $TimeOutVoicemailTTSGreetingValueExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqTimeoutVoicemailGreeting.txt"
        
                        $ttsGreetings += ("click cqTimeoutVoicemailGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqTimeoutVoicemailGreeting.txt" + '"')
        
                    }    


                    if ($TimeOutVoicemailTTSGreetingValue.Length -gt $truncateGreetings) {

                        $TimeOutVoicemailTTSGreetingValue = $TimeOutVoicemailTTSGreetingValue.Remove($TimeOutVoicemailTTSGreetingValue.Length - ($TimeOutVoicemailTTSGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."

                    }

                    $CqTimeOutVoicemailGreeting += " <br> ''$TimeOutVoicemailTTSGreetingValue''"

                }

                if ($MatchingCQ.EnableTimeoutSharedVoicemailSystemPromptSuppression -eq $false) {

                    $CQTimeoutVoicemailSystemGreeting = "--> cqTimeoutVoicemailSystemGreeting$($cqCallFlowObjectId)>Greeting <br> MS System Message] "

                    $CQTimeOutVoicemailSystemGreetingValue = (. Get-MsSystemMessage)[-1]
                    $CQTimeOutVoicemailSystemGreetingValueExport = (. Get-MsSystemMessage)[0]

                    if ($ShowTTSGreetingText) {

                        if ($ExportTTSGreetings) {

                            $CQTimeOutVoicemailSystemGreetingValueExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqTimeoutMsSystemMessage.txt"
            
                            $ttsGreetings += ("click cqTimeoutVoicemailSystemGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqTimeoutMsSystemMessage.txt" + '"')
            
                        }

                        if ($CQTimeOutVoicemailSystemGreetingValue.Length -gt $truncateGreetings) {

                            $CQTimeOutVoicemailSystemGreetingValue = $CQTimeOutVoicemailSystemGreetingValue.Remove($CQTimeOutVoicemailSystemGreetingValue.Length - ($CQTimeOutVoicemailSystemGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."

                        }

                        $CQTimeOutVoicemailSystemGreeting = $CQTimeOutVoicemailSystemGreeting.Replace("] "," <br> ''$CQTimeOutVoicemailSystemGreetingValue''] ")

                    }

                    $allMermaidNodes += "cqTimeoutVoicemailSystemGreeting$($cqCallFlowObjectId)"

                }

                else {

                    $CQTimeoutVoicemailSystemGreeting = $null

                }

            }
    
            else {
    
                $CqTimeoutVoicemailGreeting = "AudioFile"

                if ($ShowAudioFileName) {

                    $audioFileName = Optimize-DisplayName -String ($MatchingCQ.TimeoutSharedVoicemailAudioFilePromptFileName)

                    # If audio file name is not present on call queue properties
                    if (!$audioFileName) {

                        $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.TimeoutSharedVoicemailAudioFilePrompt -ApplicationId HuntGroup).FileName

                    }

                    if ($ExportAudioFiles) {

                        $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.TimeoutSharedVoicemailAudioFilePrompt -ApplicationId HuntGroup
                        [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                        $audioFileNames += ("click cqTimeoutVoicemailGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                    }
                    
                    if ($audioFileName.Length -gt $truncateGreetings) {
                
                        $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                        $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                    }
        
                    $CqTimeOutVoicemailGreeting += " <br> $audioFileName"
        
                }

                if ($MatchingCQ.EnableTimeoutSharedVoicemailSystemPromptSuppression -eq $false) {

                    $CQTimeoutVoicemailSystemGreeting = "--> cqTimeoutVoicemailSystemGreeting$($cqCallFlowObjectId)>Greeting <br> MS System Message] "

                    $CQTimeOutVoicemailSystemGreetingValue = (. Get-MsSystemMessage)[-1]
                    $CQTimeOutVoicemailSystemGreetingValueExport = (. Get-MsSystemMessage)[0]

                    if ($ShowTTSGreetingText) {

                        if ($ExportTTSGreetings) {

                            $CQTimeOutVoicemailSystemGreetingValueExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqTimeoutMsSystemMessage.txt"
            
                            $ttsGreetings += ("click cqTimeoutVoicemailSystemGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqTimeoutMsSystemMessage.txt" + '"')
            
                        }

                        if ($CQTimeOutVoicemailSystemGreetingValue.Length -gt $truncateGreetings) {

                            $CQTimeOutVoicemailSystemGreetingValue = $CQTimeOutVoicemailSystemGreetingValue.Remove($CQTimeOutVoicemailSystemGreetingValue.Length - ($CQTimeOutVoicemailSystemGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."

                        }

                        $CQTimeOutVoicemailSystemGreeting = $CQTimeOutVoicemailSystemGreeting.Replace("] "," <br> ''$CQTimeOutVoicemailSystemGreetingValue''] ")

                    }

                    $allMermaidNodes += "cqTimeoutVoicemailSystemGreeting$($cqCallFlowObjectId)"

                }

                else {

                    $CQTimeoutVoicemailSystemGreeting = $null

                }
    
            }
    
            $CqTimeoutActionFriendly = "cqTimeoutVoicemailGreeting$($cqCallFlowObjectId)>Greeting <br> $CqTimeoutVoicemailGreeting] $CQTimeoutVoicemailSystemGreeting--> cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingTimeoutIdentity)(Shared Voicemail <br> $MatchingTimeoutVoicemail)"
    
            $allMermaidNodes += @("cqTimeoutAction$($cqCallFlowObjectId)","cqTimeoutVoicemailGreeting$($cqCallFlowObjectId)","$($MatchingTimeoutIdentity)")

        }
    
    }

    switch ($MatchingCQ.NoAgentApplyTo) {
        AllCalls {
            $mdNoAgentApplyTo = "|New and Queued Calls|"
        }
        NewCalls {
            $mdNoAgentApplyTo = "|New Calls Only|"
        }
        Default {}
    }

    switch ($MatchingCQ.NoAgentAction) {
        Queue {

            $mdCqNoAgentAction = @"
cqNoAgent$($cqCallFlowObjectId){Agent Available?} --> |No| cqNoAgentAction$($cqCallFlowObjectId)(Queue Call) --> $mdNoAgentApplyTo cqResult$($cqCallFlowObjectId)
"@

            $mdCqNoAgentActionDisconnect = $null
            $mdCqNoAgentActionForward = $null

        }
        Disconnect {

            $mdCqNoAgentActionForward = $null

            if ($MatchingCQ.NoAgentDisconnectAudioFilePrompt -or $MatchingCQ.NoAgentDisconnectTextToSpeechPrompt) {

                if ($MatchingCQ.NoAgentDisconnectAudioFilePrompt) {

                    if ($ShowAudioFileName) {

                        $audioFileName = Optimize-DisplayName -String ($MatchingCQ.NoAgentDisconnectAudioFilePromptFileName)

                        # If audio file name is not present on call queue properties
                        if (!$audioFileName) {

                            $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.NoAgentDisconnectAudioFilePrompt -ApplicationId HuntGroup).FileName

                        }

                        if ($ExportAudioFiles) {

                            $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.NoAgentDisconnectAudioFilePrompt -ApplicationId HuntGroup
                            [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                            $audioFileNames += ("click cqNoAgentDisconnectAudioFilePrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                        }
        
                    
                        if ($audioFileName.Length -gt $truncateGreetings) {
                
                            $audioFileNameExtension = ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[3]
                            $audioFileName = $audioFileName.Remove($audioFileName.Length - ($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                        }
        
                        $mdCqNoAgentDisconnectGreeting = " cqNoAgentDisconnectAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile <br> $audioFileName] -->"
        
                    }

                    else {

                        $mdCqNoAgentDisconnectGreeting = " cqNoAgentDisconnectAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile] -->"

                    }

                    $allMermaidNodes += "cqNoAgentDisconnectAudioFilePrompt$($cqCallFlowObjectId)"

                }

                else {

                    if ($ShowTTSGreetingText) {

                        $NoAgentDisconnectTextToSpeechPromptExport = $MatchingCQ.NoAgentDisconnectTextToSpeechPrompt
                        $NoAgentDisconnectTextToSpeechPromptValue = Optimize-DisplayName -String $MatchingCQ.NoAgentDisconnectTextToSpeechPrompt

                        if ($ExportTTSGreetings) {

                            $NoAgentDisconnectTextToSpeechPromptExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqNoAgentDisconnectGreeting.txt"
        
                            $ttsGreetings += ("click cqNoAgentDisconnectTextToSpeechPrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqNoAgentDisconnectGreeting.txt" + '"')
        
                        }    

                        if ($NoAgentDisconnectTextToSpeechPromptValue.Length -gt $truncateGreetings) {

                            $NoAgentDisconnectTextToSpeechPromptValue = $NoAgentDisconnectTextToSpeechPromptValue.Remove($NoAgentDisconnectTextToSpeechPromptValue.Length - ($NoAgentDisconnectTextToSpeechPromptValue.Length - $truncateGreetings)).TrimEnd() + "..."

                        }

                        $mdCqNoAgentDisconnectGreeting = " cqNoAgentDisconnectTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech <br> $NoAgentDisconnectTextToSpeechPromptValue] -->"

                    }

                    else {

                        $mdCqNoAgentDisconnectGreeting = " cqNoAgentDisconnectTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech] -->"

                    }

                    $allMermaidNodes += "cqNoAgentDisconnectTextToSpeechPrompt$($cqCallFlowObjectId)"

                }

            }

            else {

                $mdCqNoAgentDisconnectGreeting = $null

            }

            if ($CombineDisconnectCallNodes -eq $true) {

                $mdCqNoAgentAction = "cqNoAgent$($cqCallFlowObjectId){Agent Available?} --> |No| cqNoAgentAction$($cqCallFlowObjectId)(ApplyTo)"

                $mdCqNoAgentActionDisconnect = "cqNoAgentAction$($cqCallFlowObjectId) ---> $mdNoAgentApplyTo$mdCqNoAgentDisconnectGreeting disconnectCall((DisconnectCall))"

                $allMermaidNodes += "disconnectCall"
    
            }
    
            else {
    
                $mdCqNoAgentAction = "cqNoAgent$($cqCallFlowObjectId){Agent Available?} --> |No| cqNoAgentAction$($cqCallFlowObjectId)(ApplyTo)"

                $mdCqNoAgentActionDisconnect = "cqNoAgentAction$($cqCallFlowObjectId) ---> $mdNoAgentApplyTo$mdCqNoAgentDisconnectGreeting cqNoAgentDisconnect$($cqCallFlowObjectId)((DisconnectCall))"

                $allMermaidNodes += "cqNoAgentDisconnect$($cqCallFlowObjectId)"
                
            }

        }
        Forward {

            $mdCqNoAgentActionDisconnect = $null
    
            if ($MatchingCQ.NoAgentActionTarget.Type -eq "User") {
        
                $MatchingNoAgentUserProperties = (Get-MgUser -UserId $MatchingCQ.NoAgentActionTarget.Id)
                $MatchingNoAgentUser = Optimize-DisplayName -String $MatchingNoAgentUserProperties.DisplayName
                $MatchingNoAgentIdentity = $MatchingNoAgentUserProperties.Id
        
                if ($FindUserLinks -eq $true) {
         
                    . New-VoiceAppUserLinkProperties -userLinkUserId $MatchingCQ.NoAgentActionTarget.Id -userLinkUserName $MatchingNoAgentUserProperties.DisplayName -userLinkVoiceAppType "Call Queue" -userLinkVoiceAppActionType "TimoutActionTarget" -userLinkVoiceAppName $MatchingCQ.Name -userLinkVoiceAppId $MatchingCQIdentity
                
                }        
        
                if ($nestedVoiceApps -notcontains $MatchingNoAgentUserProperties.Id -and $MatchingCQ.OverflowThreshold -ge 1) {
        
                    $nestedVoiceApps += $MatchingNoAgentUserProperties.Id
        
                }
        
                $CqNoAgentActionFriendly = "cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingNoAgentIdentity)(User <br> $MatchingNoAgentUser)"
        
                $allMermaidNodes += @("cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)", "$($MatchingNoAgentIdentity)")
        
            }
        
            elseif ($MatchingCQ.NoAgentActionTarget.Type -eq "Phone") {
        
                $cqNoAgentPhoneNumber = ($MatchingCQ.NoAgentActionTarget.Id).Replace("tel:","")
        
                if ($ObfuscatePhoneNumbers -eq $true) {
        
                    $cqNoAgentPhoneNumber = $cqNoAgentPhoneNumber.Remove(($cqNoAgentPhoneNumber.Length -4)) + "****"
        
                }
        
                $CqNoAgentActionFriendly = "cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)(TransferCallToTarget) --> $($cqNoAgentPhoneNumber)(External Number <br> $cqNoAgentPhoneNumber)"
        
                $allMermaidNodes += @("cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)","$($cqNoAgentPhoneNumber)")

                if ($MatchingCQ.NoAgentRedirectPhoneNumberAudioFilePrompt -or $MatchingCQ.NoAgentRedirectPhoneNumberTextToSpeechPrompt) {

                    if ($MatchingCQ.NoAgentRedirectPhoneNumberAudioFilePrompt) {

                        if ($ShowAudioFileName) {

                            $audioFileName = Optimize-DisplayName -String ($MatchingCQ.NoAgentRedirectPhoneNumberAudioFilePromptFileName)

                            # If audio file name is not present on call queue properties
                            if (!$audioFileName) {

                                $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.NoAgentRedirectPhoneNumberAudioFilePrompt -ApplicationId HuntGroup).FileName

                            }

                            if ($ExportAudioFiles) {

                                $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.NoAgentRedirectPhoneNumberAudioFilePrompt -ApplicationId HuntGroup
                                [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                                $audioFileNames += ("click cqNoAgentRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                            }
        
                    
                            if ($audioFileName.Length -gt $truncateGreetings) {
                
                                $audioFileNameExtension = ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[3]
                                $audioFileName = $audioFileName.Remove($audioFileName.Length - ($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                            }
        
                            $CqNoAgentActionFriendly = "cqNoAgentRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile <br> $audioFileName] --> $CqNoAgentActionFriendly"
        
                        }

                        else {

                            $CqNoAgentActionFriendly = "cqNoAgentRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile] --> $CqNoAgentActionFriendly"

                        }

                        $allMermaidNodes += "cqNoAgentRedirectPhoneNumberAudioFilePrompt$($cqCallFlowObjectId)"

                    }

                    else {

                        if ($ShowTTSGreetingText) {

                            $NoAgentRedirectPhoneNumberTextToSpeechPromptExport = $MatchingCQ.NoAgentRedirectPhoneNumberTextToSpeechPrompt
                            $NoAgentRedirectPhoneNumberTextToSpeechPromptValue = Optimize-DisplayName -String $MatchingCQ.NoAgentRedirectPhoneNumberTextToSpeechPrompt

                            if ($ExportTTSGreetings) {

                                $NoAgentRedirectPhoneNumberTextToSpeechPromptExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqNoAgentRedirectPhoneNumberGreeting.txt"
        
                                $ttsGreetings += ("click cqNoAgentRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqNoAgentRedirectPhoneNumberGreeting.txt" + '"')
        
                            }    

                            if ($NoAgentRedirectPhoneNumberTextToSpeechPromptValue.Length -gt $truncateGreetings) {

                                $NoAgentRedirectPhoneNumberTextToSpeechPromptValue = $NoAgentRedirectPhoneNumberTextToSpeechPromptValue.Remove($NoAgentRedirectPhoneNumberTextToSpeechPromptValue.Length - ($NoAgentRedirectPhoneNumberTextToSpeechPromptValue.Length - $truncateGreetings)).TrimEnd() + "..."

                            }

                            $CqNoAgentActionFriendly = "cqNoAgentRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech <br> $NoAgentRedirectPhoneNumberTextToSpeechPromptValue] --> $CqNoAgentActionFriendly"

                        }

                        else {

                            $CqNoAgentActionFriendly = "cqNoAgentRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech] --> $CqNoAgentActionFriendly"

                        }

                        $allMermaidNodes += "cqNoAgentRedirectPhoneNumberTextToSpeechPrompt$($cqCallFlowObjectId)"

                    }

                }
                
            }
        
            else {

                if ($ShowTransferCallToTargetType -eq $true) {

                    switch ($MatchingCQ.NoAgentActionTarget.Type) {
                        ApplicationEndpoint {

                            $cqTransferCallToTargetTypeAdditionalInfo = " <br> Resource Account"

                        }
                        ConfigurationEndpoint {

                            $cqTransferCallToTargetTypeAdditionalInfo = " <br> Voice App"

                        }
                        Default {

                            $cqTransferCallToTargetTypeAdditionalInfo = $null

                        }
                    }

                }

                else {

                    $cqTransferCallToTargetTypeAdditionalInfo = $null

                }
        
                $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $MatchingCQ.NoAgentActionTarget.Id -and $_.ApplicationId -eq $applicationIdAa}
        
                if ($matchingApplicationInstanceCheckAa -or $allAutoAttendantIds -contains $MatchingCQ.NoAgentActionTarget.Id) {
        
                    $MatchingNoAgentAA = ($allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.NoAgentActionTarget.Id -or $_.Identity -eq $MatchingCQ.NoAgentActionTarget.Id})
        
                    $MatchingNoAgentAA.Name = Optimize-DisplayName -String $MatchingNoAgentAA.Name
        
                    $CqNoAgentActionFriendly = "cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)(TransferCallToTarget$cqTransferCallToTargetTypeAdditionalInfo) --> $($MatchingNoAgentAA.Identity)([Auto Attendant <br> $($MatchingNoAgentAA.Name)])"
        
                    if ($nestedVoiceApps -notcontains $MatchingNoAgentAA.Identity -and $MatchingCQ.OverflowThreshold -ge 1 -and $MatchingCQ.TimeoutThreshold -ge 1) {
        
                        $nestedVoiceApps += $MatchingNoAgentAA.Identity
        
                    }
        
                    $allMermaidNodes += @("cqNoAgentAction$($cqCallFlowObjectId)", "$($MatchingNoAgentAA.Identity)", "cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)")
        
                }
        
                else {
        
                    $MatchingNoAgentCQ = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.NoAgentActionTarget.Id -or $_.Identity -eq $MatchingCQ.NoAgentActionTarget.Id})
        
                    $MatchingNoAgentCQ.Name = Optimize-DisplayName -String $MatchingNoAgentCQ.Name
        
                    $CqNoAgentActionFriendly = "cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)(TransferCallToTarget$cqTransferCallToTargetTypeAdditionalInfo) --> $($MatchingNoAgentCQ.Identity)([Call Queue <br> $($MatchingNoAgentCQ.Name)])"
        
                    if ($nestedVoiceApps -notcontains $MatchingNoAgentCQ.Identity -and $MatchingCQ.OverflowThreshold -ge 1 -and $MatchingCQ.TimeoutThreshold -ge 1) {
        
                        $nestedVoiceApps += $MatchingNoAgentCQ.Identity
        
                    }
        
                    $allMermaidNodes += @("cqNoAgentAction$($cqCallFlowObjectId)", "$($MatchingNoAgentCQ.Identity)", "cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)")
        
                }

                if ($MatchingCQ.NoAgentRedirectVoiceAppAudioFilePrompt -or $MatchingCQ.NoAgentRedirectVoiceAppTextToSpeechPrompt) {

                    if ($MatchingCQ.NoAgentRedirectVoiceAppAudioFilePrompt) {

                        if ($ShowAudioFileName) {

                            $audioFileName = Optimize-DisplayName -String ($MatchingCQ.NoAgentRedirectVoiceAppAudioFilePromptFileName)

                            # If audio file name is not present on call queue properties
                            if (!$audioFileName) {

                                $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.NoAgentRedirectVoiceAppAudioFilePrompt -ApplicationId HuntGroup).FileName

                            }

                            if ($ExportAudioFiles) {

                                $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.NoAgentRedirectVoiceAppAudioFilePrompt -ApplicationId HuntGroup
                                [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                                $audioFileNames += ("click cqNoAgentRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                            }
        
                    
                            if ($audioFileName.Length -gt $truncateGreetings) {
                
                                $audioFileNameExtension = ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length - 4)..$audioFileName.Length])[3]
                                $audioFileName = $audioFileName.Remove($audioFileName.Length - ($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                            }
        
                            $CqNoAgentActionFriendly = "cqNoAgentRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile <br> $audioFileName] --> $CqNoAgentActionFriendly"
        
                        }

                        else {

                            $CqNoAgentActionFriendly = "cqNoAgentRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId)>Greeting <br> AudioFile] --> $CqNoAgentActionFriendly"

                        }

                        $allMermaidNodes += "cqNoAgentRedirectVoiceAppAudioFilePrompt$($cqCallFlowObjectId)"

                    }

                    else {

                        if ($ShowTTSGreetingText) {

                            $NoAgentRedirectVoiceAppTextToSpeechPromptExport = $MatchingCQ.NoAgentRedirectVoiceAppTextToSpeechPrompt
                            $NoAgentRedirectVoiceAppTextToSpeechPromptValue = Optimize-DisplayName -String $MatchingCQ.NoAgentRedirectVoiceAppTextToSpeechPrompt

                            if ($ExportTTSGreetings) {

                                $NoAgentRedirectVoiceAppTextToSpeechPromptExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqNoAgentRedirectVoiceAppGreeting.txt"
        
                                $ttsGreetings += ("click cqNoAgentRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqNoAgentRedirectVoiceAppGreeting.txt" + '"')
        
                            }    

                            if ($NoAgentRedirectVoiceAppTextToSpeechPromptValue.Length -gt $truncateGreetings) {

                                $NoAgentRedirectVoiceAppTextToSpeechPromptValue = $NoAgentRedirectVoiceAppTextToSpeechPromptValue.Remove($NoAgentRedirectVoiceAppTextToSpeechPromptValue.Length - ($NoAgentRedirectVoiceAppTextToSpeechPromptValue.Length - $truncateGreetings)).TrimEnd() + "..."

                            }

                            $CqNoAgentActionFriendly = "cqNoAgentRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech <br> $NoAgentRedirectVoiceAppTextToSpeechPromptValue] --> $CqNoAgentActionFriendly"

                        }

                        else {

                            $CqNoAgentActionFriendly = "cqNoAgentRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId)>Greeting <br> TextToSpeech] --> $CqNoAgentActionFriendly"

                        }

                        $allMermaidNodes += "cqNoAgentRedirectVoiceAppTextToSpeechPrompt$($cqCallFlowObjectId)"

                    }

                }
        
            }

            $mdCqNoAgentAction = "cqNoAgent$($cqCallFlowObjectId){Agent Available?} --> |No| cqNoAgentAction$($cqCallFlowObjectId)(ApplyTo)"

            $mdCqNoAgentActionForward = "cqNoAgentAction$($cqCallFlowObjectId) ---> $mdNoAgentApplyTo $CqNoAgentActionFriendly"
        
        }
        Voicemail {

            $mdCqNoAgentActionDisconnect = $null

            $MatchingNoAgentPersonalVoicemailUserProperties = (Get-MgUser -UserId $MatchingCQ.NoAgentActionTarget.Id)
            $MatchingNoAgentPersonalVoicemailUser = Optimize-DisplayName -String $MatchingNoAgentPersonalVoicemailUserProperties.DisplayName
            $MatchingNoAgentPersonalVoicemailIdentity = $MatchingNoAgentPersonalVoicemailUserProperties.Id
        
            if ($FindUserLinks -eq $true) {
         
                . New-VoiceAppUserLinkProperties -userLinkUserId $MatchingCQ.NoAgentActionTarget.Id -userLinkUserName $MatchingNoAgentPersonalVoicemailUserProperties.DisplayName -userLinkVoiceAppType "Call Queue" -userLinkVoiceAppActionType "TimoutActionTargetPersonalVoicemail" -userLinkVoiceAppName $MatchingCQ.Name -userLinkVoiceAppId $MatchingCQIdentity
            
            }
        
            $CqNoAgentActionFriendly = "cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)(TransferCallToTarget) --> cqPersonalVoicemail$($MatchingNoAgentPersonalVoicemailIdentity)(Personal Voicemail <br> $MatchingNoAgentPersonalVoicemailUser)"
        
            $mdCqNoAgentAction = "cqNoAgent$($cqCallFlowObjectId){Agent Available?} --> |No| cqNoAgentAction$($cqCallFlowObjectId)(ApplyTo)"
        
            $mdCqNoAgentActionForward = "cqNoAgentAction$($cqCallFlowObjectId) ---> $mdNoAgentApplyTo $CqNoAgentActionFriendly"
        
            $allMermaidNodes += @("cqNoAgentAction$($cqCallFlowObjectId)", "$($MatchingNoAgentPersonalVoicemailIdentity)", "cqPersonalVoicemail$($MatchingNoAgentPersonalVoicemailIdentity)", "cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)")
        
        }
        SharedVoicemail {

            $mdCqNoAgentActionDisconnect = $null

            $MatchingNoAgentVoicemailProperties = (Get-MgGroup -GroupId $MatchingCQ.NoAgentActionTarget.Id)
            $MatchingNoAgentVoicemail = Optimize-DisplayName -String $MatchingNoAgentVoicemailProperties.DisplayName
            $MatchingNoAgentIdentity = $MatchingNoAgentVoicemailProperties.Id
        
            if ($ShowSharedVoicemailGroupMembers -eq $true) {
        
                . Get-SharedVoicemailGroupMembers -SharedVoicemailGroupId $MatchingCQ.NoAgentActionTarget.Id
        
                $MatchingNoAgentVoicemail = "$MatchingNoAgentVoicemail$mdSharedVoicemailGroupMembers"
        
            }
        
            if ($MatchingCQ.NoAgentSharedVoicemailTextToSpeechPrompt) {
        
                $CqNoAgentVoicemailGreeting = "TextToSpeech"
                
                if ($ShowTTSGreetingText) {
        
                    $NoAgentVoicemailTTSGreetingValueExport = $MatchingCQ.NoAgentSharedVoicemailTextToSpeechPrompt
                    $NoAgentVoicemailTTSGreetingValue = Optimize-DisplayName -String $MatchingCQ.NoAgentSharedVoicemailTextToSpeechPrompt
        
                    if ($ExportTTSGreetings) {
        
                        $NoAgentVoicemailTTSGreetingValueExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqNoAgentVoicemailGreeting.txt"
        
                        $ttsGreetings += ("click cqNoAgentVoicemailGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqNoAgentVoicemailGreeting.txt" + '"')
        
                    }    
        
        
                    if ($NoAgentVoicemailTTSGreetingValue.Length -gt $truncateGreetings) {
        
                        $NoAgentVoicemailTTSGreetingValue = $NoAgentVoicemailTTSGreetingValue.Remove($NoAgentVoicemailTTSGreetingValue.Length - ($NoAgentVoicemailTTSGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
        
                    }
        
                    $CqNoAgentVoicemailGreeting += " <br> ''$NoAgentVoicemailTTSGreetingValue''"
        
                }
        
                if ($MatchingCQ.EnableNoAgentSharedVoicemailSystemPromptSuppression -eq $false) {
        
                    $CQNoAgentVoicemailSystemGreeting = "--> cqNoAgentVoicemailSystemGreeting$($cqCallFlowObjectId)>Greeting <br> MS System Message] "
        
                    $CQNoAgentVoicemailSystemGreetingValue = (. Get-MsSystemMessage)[-1]
                    $CQNoAgentVoicemailSystemGreetingValueExport = (. Get-MsSystemMessage)[0]
        
                    if ($ShowTTSGreetingText) {
        
                        if ($ExportTTSGreetings) {
        
                            $CQNoAgentVoicemailSystemGreetingValueExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqNoAgentMsSystemMessage.txt"
            
                            $ttsGreetings += ("click cqNoAgentVoicemailSystemGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqNoAgentMsSystemMessage.txt" + '"')
            
                        }
        
                        if ($CQNoAgentVoicemailSystemGreetingValue.Length -gt $truncateGreetings) {
        
                            $CQNoAgentVoicemailSystemGreetingValue = $CQNoAgentVoicemailSystemGreetingValue.Remove($CQNoAgentVoicemailSystemGreetingValue.Length - ($CQNoAgentVoicemailSystemGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
        
                        }
        
                        $CQNoAgentVoicemailSystemGreeting = $CQNoAgentVoicemailSystemGreeting.Replace("] "," <br> ''$CQNoAgentVoicemailSystemGreetingValue''] ")
        
                    }
        
                    $allMermaidNodes += "cqNoAgentVoicemailSystemGreeting$($cqCallFlowObjectId)"
        
                }
        
                else {
        
                    $CQNoAgentVoicemailSystemGreeting = $null
        
                }
        
            }
        
            else {
        
                $CqNoAgentVoicemailGreeting = "AudioFile"
        
                if ($ShowAudioFileName) {
        
                    $audioFileName = Optimize-DisplayName -String ($MatchingCQ.NoAgentSharedVoicemailAudioFilePromptFileName)
        
                    # If audio file name is not present on call queue properties
                    if (!$audioFileName) {
        
                        $audioFileName = Optimize-DisplayName -String (Get-CsOnlineAudioFile -Identity $MatchingCQ.NoAgentSharedVoicemailAudioFilePrompt -ApplicationId HuntGroup).FileName
        
                    }
        
                    if ($ExportAudioFiles) {
        
                        $content = Export-CsOnlineAudioFile -Identity $MatchingCQ.NoAgentSharedVoicemailAudioFilePrompt -ApplicationId HuntGroup
                        [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                        $audioFileNames += ("click cqNoAgentVoicemailGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$audioFileName" + '"')
        
                    }
                    
                    if ($audioFileName.Length -gt $truncateGreetings) {
                
                        $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                        $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                    }
        
                    $CqNoAgentVoicemailGreeting += " <br> $audioFileName"
        
                }
        
                if ($MatchingCQ.EnableNoAgentSharedVoicemailSystemPromptSuppression -eq $false) {
        
                    $CQNoAgentVoicemailSystemGreeting = "--> cqNoAgentVoicemailSystemGreeting$($cqCallFlowObjectId)>Greeting <br> MS System Message] "
        
                    $CQNoAgentVoicemailSystemGreetingValue = (. Get-MsSystemMessage)[-1]
                    $CQNoAgentVoicemailSystemGreetingValueExport = (. Get-MsSystemMessage)[0]
        
                    if ($ShowTTSGreetingText) {
        
                        if ($ExportTTSGreetings) {
        
                            $CQNoAgentVoicemailSystemGreetingValueExport | Out-File "$FilePath\$($cqCallFlowObjectId)_cqNoAgentMsSystemMessage.txt"
            
                            $ttsGreetings += ("click cqNoAgentVoicemailSystemGreeting$($cqCallFlowObjectId) " + '"' + "$FilePath\$($cqCallFlowObjectId)_cqNoAgentMsSystemMessage.txt" + '"')
            
                        }
        
                        if ($CQNoAgentVoicemailSystemGreetingValue.Length -gt $truncateGreetings) {
        
                            $CQNoAgentVoicemailSystemGreetingValue = $CQNoAgentVoicemailSystemGreetingValue.Remove($CQNoAgentVoicemailSystemGreetingValue.Length - ($CQNoAgentVoicemailSystemGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
        
                        }
        
                        $CQNoAgentVoicemailSystemGreeting = $CQNoAgentVoicemailSystemGreeting.Replace("] "," <br> ''$CQNoAgentVoicemailSystemGreetingValue''] ")
        
                    }
        
                    $allMermaidNodes += "cqNoAgentVoicemailSystemGreeting$($cqCallFlowObjectId)"
        
                }
        
                else {
        
                    $CQNoAgentVoicemailSystemGreeting = $null
        
                }
        
            }
        
            $CqNoAgentActionFriendly = "cqNoAgentVoicemailGreeting$($cqCallFlowObjectId)>Greeting <br> $CqNoAgentVoicemailGreeting] $CQNoAgentVoicemailSystemGreeting--> cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingNoAgentIdentity)(Shared Voicemail <br> $MatchingNoAgentVoicemail)"
        
            $mdCqNoAgentAction = "cqNoAgent$($cqCallFlowObjectId){Agent Available?} --> |No| cqNoAgentAction$($cqCallFlowObjectId)(ApplyTo)"
        
            $mdCqNoAgentActionForward = "cqNoAgentAction$($cqCallFlowObjectId) ---> $mdNoAgentApplyTo $CqNoAgentActionFriendly"

            $allMermaidNodes += @("cqNoAgentAction$($cqCallFlowObjectId)", "cqNoAgentVoicemailGreeting$($cqCallFlowObjectId)", "$($MatchingNoAgentIdentity)", "cqNoAgentActionTransferCallToTarget$($cqCallFlowObjectId)")
        
        }
        Default {}
    }

    $allMermaidNodes += @("cqNoAgent$($cqCallFlowObjectId)","cqNoAgentApplyTo$($cqCallFlowObjectId)","cqNoAgentAction$($cqCallFlowObjectId)")

    # Create empty mermaid element for agent list
    $mdCqAgentsDisplayNames = @"
"@

    # Define agent counter for unique mermaid element names
    $AgentCounter = 1

    # add each agent to the empty agents mermaid element
    foreach ($CqAgent in $CqAgents) {
        $AgentDisplayName = Optimize-DisplayName -String (Get-MgUser -UserId $CqAgent.ObjectId).DisplayName

        if ($FindUserLinks -eq $true) {
         
            . New-VoiceAppUserLinkProperties -userLinkUserId $CqAgent.ObjectId -userLinkUserName $AgentDisplayName -userLinkVoiceAppType "Call Queue" -userLinkVoiceAppActionType "Agent" -userLinkVoiceAppName $MatchingCQ.Name -userLinkVoiceAppId $MatchingCQIdentity
        
        }

        if ($ShowCqAgentPhoneNumbers -eq $true) {

            $CqAgentCsOnlineUser = (Get-CsOnlineUser -Identity $($CqAgent.ObjectId))
            $CqAgentPhoneNumber = $CqAgentCsOnlineUser.LineUri

            if (!$CqAgentPhoneNumber) {

                $CqAgentPhoneNumber = "No Number Assigned"

            }

            else {

                if ($ShowPhoneNumberType -eq $true) {

                    $CqAgentPhoneNumberType = (Get-CsPhoneNumberAssignment -TelephoneNumber $CqAgentPhoneNumber.Replace("tel:","")).NumberType
                
                }

                if ($CqAgentPhoneNumber -match "tel:") {

                    $CqAgentPhoneNumber = $CqAgentPhoneNumber.Replace("tel:","")
    
                }
    
                if ($CqAgentPhoneNumber -notmatch "\+") {
    
                    $CqAgentPhoneNumber = "+" + $CqAgentPhoneNumber
    
                }
    
            }

            if ($ObfuscatePhoneNumbers -eq $true) {

                $CqAgentPhoneNumber = $CqAgentPhoneNumber.Remove(($CqAgentPhoneNumber.Length -4)) + "****"

            }

            if ($ShowPhoneNumberType -eq $true) {

                $CqAgentPhoneNumber = $CqAgentPhoneNumber + "<br>$CqAgentPhoneNumberType"

            }

            $AgentDisplayName = "$AgentDisplayName <br> $CqAgentPhoneNumber"

        }

        if ($ShowUserOutboundCallingIds -eq $true){

            if ($CqAgentCsOnlineUser.CallingLineIdentity.Name) {

                $checkCallingLineId = Get-CsCallingLineIdentity -Identity $CqAgentCsOnlineUser.CallingLineIdentity.Name

                if ($checkCallingLineId.CallingIDSubstitute -ne "LineUri") {

                    switch ($checkCallingLineId.CallingIDSubstitute) {
                        Service {

                            Write-Warning -Message "Service Number Calling Line Identities are going to be deprecated soon. Migrate to Resource Account substitute soon. (MC505122)"

                            $serviceNumber = "+$($checkCallingLineId.ServiceNumber)"

                            if ($ObfuscatePhoneNumbers -eq $true) {

                                $serviceNumber = $serviceNumber.Remove(($serviceNumber.Length -4)) + "****"
                
                            }

                            $cqAgentCallingLineId = "Outbound Calling Id:<br>" + $($checkCallingLineId.Identity).Replace("Tag:","") + " $serviceNumber<br>Type: $($checkCallingLineId.CallingIDSubstitute)"

                        }
                        Resource {

                            $resourceAccountNumber = ($allResourceAccounts | Where-Object {$_.ObjectId -eq $($checkCallingLineId.ResourceAccount)}).PhoneNumber.Replace("tel:","")

                            if ($ObfuscatePhoneNumbers -eq $true) {

                                $resourceAccountNumber = $resourceAccountNumber.Remove(($resourceAccountNumber.Length -4)) + "****"
                
                            }

                            $cqAgentCallingLineId = "Outbound Calling Id:<br>" + $($checkCallingLineId.Identity).Replace("Tag:","") + " $resourceAccountNumber<br>Type: $($checkCallingLineId.CallingIDSubstitute)"

                        }

                        Anonymous {

                            $cqAgentCallingLineId = "Outbound Calling Id:<br>Anonymous"

                        }
                        Default {}
                    } 

                }

                $AgentDisplayName = "$AgentDisplayName<br>$cqAgentCallingLineId"

            }

        }


        if ($ShowCqAgentOptInStatus -eq $true) {

            $AgentDisplayName = "$AgentDisplayName <br> OptIn: $($CqAgent.OptIn)"

        }

        if ($CqRoutingMethod -eq "Serial") {

            $serialAgentNumber = "|$AgentCounter|"

        }

        else {

            $serialAgentNumber = $null

        }

        $AgentDisplayNames = "agentListType$($cqCallFlowObjectId) -.-> $serialAgentNumber agent$($cqCallFlowObjectId)$($AgentCounter)($AgentDisplayName)`n"

        $allMermaidNodes += @("agentListType$($cqCallFlowObjectId)","agent$($cqCallFlowObjectId)$($AgentCounter)")

        $mdCqAgentsDisplayNames += $AgentDisplayNames

        $AgentCounter ++

    }

    $allMermaidNodes += "$($MatchingCQIdentity)"

    # Add outbound calling IDs if available and if selected
    if ($ShowCqOutboundCallingIds -eq $true) {

        $oboResourceAccounts = "Outbound Calling Ids:"

        foreach ($CqOboResourceAccountId in $CqOboResourceAccountIds) {

            $oboResourceAccount = $allResourceAccounts | Where-Object {$_.ObjectId -eq $CqOboResourceAccountId}
            $oboResourceAccountDisplayName = Optimize-DisplayName -String $oboResourceAccount.DisplayName
            $oboResourceAccountPhoneNumber = $oboResourceAccount.PhoneNumber.Replace("tel:","")

            if ($ShowPhoneNumberType -eq $true) {

                $oboResourceAccountPhoneNumberPhoneNumberType = (Get-CsPhoneNumberAssignment -TelephoneNumber $oboResourceAccount.PhoneNumber.Replace("tel:","")).NumberType
            
            }

            if ($ObfuscatePhoneNumbers -eq $true) {

                $oboResourceAccountPhoneNumber = $oboResourceAccountPhoneNumber.Remove(($oboResourceAccountPhoneNumber.Length -4)) + "****"

            }

            if ($ShowPhoneNumberType -eq $true) {

                $oboResourceAccountPhoneNumber = $oboResourceAccountPhoneNumber + "<br>$oboResourceAccountPhoneNumberPhoneNumberType"

            }

            $oboResourceAccounts += "<br>$oboResourceAccountDisplayName $oboResourceAccountPhoneNumber"

        }

        $mdOutboundCallingIds = @"

cqSettingsContainer$($cqCallFlowObjectId) -.- cqOutboundCallingIds$($cqCallFlowObjectId)[($($oboResourceAccounts))] -.- 
    
"@
    
        $allMermaidNodes += "cqOutboundCallingIds$($cqCallFlowObjectId)"

    }

    else {

        if ($ShowCqAuthorizedUsers -eq $true -and $MatchingCQ.AuthorizedUsers) {

            $mdOutboundCallingIds = $null

        }

        else {

            $mdOutboundCallingIds = "cqSettingsContainer$($cqCallFlowObjectId) -.- timeOut$($cqCallFlowObjectId)"

        }

    }

    if ($ShowCqAuthorizedUsers -eq $true -and $MatchingCQ.AuthorizedUsers) {

        if ($ShowCqOutboundCallingIds -eq $true -and $CqOboResourceAccountIds) {

            $mdOutboundCallingIds = $mdOutboundCallingIds.Replace(")] -.- ",")]")

            $mdCqAuthorizedUsers = "cqOutboundCallingIds$($cqCallFlowObjectId) -.- cqAuthorizedUsers$($cqCallFlowObjectId)[(Authorized Users<br>"

        }

        else {

            $mdCqAuthorizedUsers = "cqSettingsContainer$($cqCallFlowObjectId) -.- cqAuthorizedUsers$($cqCallFlowObjectId)[(Authorized Users<br>"

        }

        foreach ($cqAuthorizedUser in $MatchingCQ.AuthorizedUsers.Guid) {

            $cqAuthorizedCsOnlineUser = Get-CsOnlineUser -Identity $cqAuthorizedUser

            if (!$cqAuthorizedCsOnlineUser.TeamsVoiceApplicationsPolicy.Name) {

                Write-Warning -Message "User $($cqAuthorizedCsOnlineUser.DisplayName) is an authorized user of CQ $($MatchingCQ.Name) but doesn't have a Voice Application Policy assigned."

                $mdCqAuthorizedUserVoiceApplicationPolicy = ", Assigned Policy: None"

            }

            else {

                # $cqAuthorizedUserVoiceApplicationPolicy = Get-CsTeamsVoiceApplicationsPolicy -Identity $cqAuthorizedCsOnlineUser.TeamsVoiceApplicationsPolicy.Name

                $mdCqAuthorizedUserVoiceApplicationPolicy = ", Assigned Policy: $($cqAuthorizedCsOnlineUser.TeamsVoiceApplicationsPolicy.Name)"

            }

            $mdCqAuthorizedUsers += ($cqAuthorizedCsOnlineUser.DisplayName) + $mdCqAuthorizedUserVoiceApplicationPolicy + "<br>"

        }

        $mdCqAuthorizedUsers = $mdCqAuthorizedUsers.Remove(($mdCqAuthorizedUsers.Length -4),4)
        $mdCqAuthorizedUsers += ")] -.-"

        $allMermaidNodes += "cqAuthorizedUsers$($cqCallFlowObjectId)"

    }

    else {
        
        $mdCqAuthorizedUsers = $null

    }
    
    # Create default callflow mermaid code

    if ($CombineCallConnectedNodes -eq $true) {

        $mdCallSuccess = "callSuccess((Call Connected))"

        $allMermaidNodes += "cqEnd"

    }

    else {

        $mdCallSuccess = "cqEnd$($cqCallFlowObjectId)((Call Connected))"

        $allMermaidNodes += "cqEnd$($cqCallFlowObjectId)"

    }

$mdCallQueueCallFlow =@"
$($MatchingCQIdentity)([Call Queue <br> $($CqName)]) -->$cqGreetingNode overFlow$($cqCallFlowObjectId){More than $CqOverFlowThreshold <br> Active Calls?}
overFlow$($cqCallFlowObjectId) --> |Yes| $CqOverFlowActionFriendly
overFlow$($cqCallFlowObjectId) ---> |No| routingMethod$($cqCallFlowObjectId)

subgraph subgraphCallDistribution$($cqCallFlowObjectId)[Call Distribution: $CqName]
subgraph subgraphCqSettings$($cqCallFlowObjectId)[CQ Settings]
routingMethod$($cqCallFlowObjectId)[(Routing Method: $CqRoutingMethod)] --> agentAlertTime$($cqCallFlowObjectId)
agentAlertTime$($cqCallFlowObjectId)[(Agent Alert Time: $CqAgentAlertTime)] -.- cqSettingsContainer$($cqCallFlowObjectId)
cqSettingsContainer$($cqCallFlowObjectId)[(Music On Hold: $CqMusicOnHold <br> Conference Mode Enabled: $CqConferenceMode <br> Agent Opt Out Allowed: $CqAgentOptOut <br> Presence Based Routing: $CqPresenceBasedRouting <br> TTS Greeting Language: $CqLanguageId)]
$mdOutboundCallingIds
$mdCqAuthorizedUsers
timeOut$($cqCallFlowObjectId)[(Timeout: $CqTimeOut Seconds)]
end
agentAlertTime$($cqCallFlowObjectId) --> subgraphAgents$($cqCallFlowObjectId)
subgraph subgraphAgents$($cqCallFlowObjectId)[Agents List]
agentListType$($cqCallFlowObjectId)[(Agent List Type: $CqAgentListType)]
$mdCqAgentsDisplayNames
end
subgraphAgents$($cqCallFlowObjectId) --> cqNoAgent$($cqCallFlowObjectId){Agent Available?} --> |Yes| cqResult$($cqCallFlowObjectId){Agent Answered?}
$mdCqNoAgentAction
end

cqResult$($cqCallFlowObjectId) --> |Yes| $mdCallSuccess
cqResult$($cqCallFlowObjectId) --> |No| timeOut$($cqCallFlowObjectId) --> $CqTimeoutActionFriendly

$mdCqNoAgentActionDisconnect
$mdCqNoAgentActionForward

"@

    if ($mermaidCode -notcontains $mdCallQueueCallFlow) {

        # Add complete call queue call flow if overflow and timeout thresholds are set
        if ($MatchingCQ.OverflowThreshold -ge 1 -and $MatchingCQ.TimeoutThreshold -ge 1) {

            $mermaidCode += $mdCallQueueCallFlow

        }

        else {

            # Add only overflow call flow if overflow threshold is 0
            if ($MatchingCQ.OverflowThreshold -eq 0) {

                if ($ShowCqOutboundCallingIds -eq $true) {

                    $mdOutboundCallingIds = " -.- " + $mdOutboundCallingIds.Split(" -.- ")[1]

                }

                else {

                    $mdOutboundCallingIds = $null

                }

                $mdCallQueueCallFlow =@"
$($MatchingCQIdentity)([Call Queue <br> $($CqName)]) -->$cqGreetingNode overFlow$($cqCallFlowObjectId)[(Overflow Threshold: $CqOverFlowThreshold <br> Immediate Overflow Action <br> TTS Greeting Language: $CqLanguageId)]$mdOutboundCallingIds
overFlow$($cqCallFlowObjectId) --> $CqOverFlowActionFriendly

"@

            $mermaidCode += $mdCallQueueCallFlow
            }

            # Add only timeout call flow if timeout threshold is 0 and overflow threshold is not 0
            else {
                $mdCallQueueCallFlow =@"
$($MatchingCQIdentity)([Call Queue <br> $($CqName)]) -->$cqGreetingNode timeOut$($cqCallFlowObjectId)[(Timeout Threshold: $CqTimeOut <br> Immediate Timeout Action <br> TTS Greeting Language: $CqLanguageId)]
timeOut$($cqCallFlowObjectId) --> $CqTimeoutActionFriendly

"@
            $mermaidCode += $mdCallQueueCallFlow
            }
        }

        $allMermaidNodes += @("cqGreeting$($cqCallFlowObjectId)","overFlow$($cqCallFlowObjectId)","routingMethod$($cqCallFlowObjectId)","agentAlertTime$($cqCallFlowObjectId)","cqSettingsContainer$($cqCallFlowObjectId)","timeOut$($cqCallFlowObjectId)","agentListType$($cqCallFlowObjectId)","cqResult$($cqCallFlowObjectId)")
        $allSubgraphs += @("subgraphCallDistribution$($cqCallFlowObjectId)","subgraphCqSettings$($cqCallFlowObjectId)","subgraphAgents$($cqCallFlowObjectId)")

    }

}

. Set-Mermaid -DocType $DocType

#This is needed to determine if the Get-CallFlow function is running for the first time or not.
$mdNodePhoneNumbersCounter = 0

function Get-CallFlow {
    param (
        [Parameter(Mandatory = $false)][String]$VoiceAppId,
        [Parameter(Mandatory = $false)][String]$VoiceAppName,
        [Parameter(Mandatory = $false)][String]$voiceAppType
    )
    
    if (!$VoiceAppName -and !$voiceAppType -and !$VoiceAppId) {
        
        $VoiceApps = @()

        $VoiceAppAas = $allAutoAttendants
        $VoiceAppCqs = $allCallQueues

        foreach ($VoiceApp in $VoiceAppAas) {

            $VoiceAppProperties = New-Object -TypeName psobject
            $ResourceAccountPhoneNumbers = ""

            $VoiceAppProperties | Add-Member -MemberType NoteProperty -Name "Name" -Value $VoiceApp.Name
            $VoiceAppProperties | Add-Member -MemberType NoteProperty -Name "Type" -Value "Auto Attendant"

            $ApplicationInstanceAssociationCounter = 0

            foreach ($ResourceAccount in $VoiceApp.ApplicationInstances) {

                $ResourceAccountPhoneNumber = ($allResourceAccounts | Where-Object {$_.ObjectId -eq $ResourceAccount}).PhoneNumber

                if ($ResourceAccountPhoneNumber) {

                    $ResourceAccountPhoneNumber = $ResourceAccountPhoneNumber.Replace("tel:","")

                    # Add leading + if PS fails to read it from online application
                    if ($ResourceAccountPhoneNumber -notmatch "\+") {

                        $ResourceAccountPhoneNumber = "+$ResourceAccountPhoneNumber"

                    }

                    $ResourceAccountPhoneNumbers += "$ResourceAccountPhoneNumber, "

                    $ApplicationInstanceAssociationCounter ++
    
                }

            }

            if ($ApplicationInstanceAssociationCounter -lt 2) {

                $ResourceAccountPhoneNumbers = $ResourceAccountPhoneNumbers.Replace(",","")

            }

            else {

                $ResourceAccountPhoneNumbers = $ResourceAccountPhoneNumbers.TrimEnd(", ")

            }

            $VoiceAppProperties | Add-Member -MemberType NoteProperty -Name "PhoneNumbers" -Value $ResourceAccountPhoneNumbers

            $VoiceApps += $VoiceAppProperties

        }

        foreach ($VoiceApp in $VoiceAppCqs) {

            $VoiceAppProperties = New-Object -TypeName psobject
            $ResourceAccountPhoneNumbers = ""

            $VoiceAppProperties | Add-Member -MemberType NoteProperty -Name "Name" -Value $VoiceApp.Name
            $VoiceAppProperties | Add-Member -MemberType NoteProperty -Name "Type" -Value "Call Queue"

            $ApplicationInstanceAssociationCounter = 0

            foreach ($ResourceAccount in $VoiceApp.ApplicationInstances) {

                $ResourceAccountPhoneNumber = ($allResourceAccounts | Where-Object {$_.ObjectId -eq $ResourceAccount}).PhoneNumber

                if ($ResourceAccountPhoneNumber) {

                    $ResourceAccountPhoneNumber = $ResourceAccountPhoneNumber.Replace("tel:","")

                    # Add leading + if PS fails to read it from online application
                    if ($ResourceAccountPhoneNumber -notmatch "\+") {

                        $ResourceAccountPhoneNumber = "+$ResourceAccountPhoneNumber"

                    }

                    $ResourceAccountPhoneNumbers += "$ResourceAccountPhoneNumber, "

                    $ApplicationInstanceAssociationCounter ++
    
                }

            }

            if ($ApplicationInstanceAssociationCounter -lt 2) {

                $ResourceAccountPhoneNumbers = $ResourceAccountPhoneNumbers.Replace(",","")

            }

            else {

                $ResourceAccountPhoneNumbers = $ResourceAccountPhoneNumbers.TrimEnd(", ")
                
            }

            $VoiceAppProperties | Add-Member -MemberType NoteProperty -Name "PhoneNumbers" -Value $ResourceAccountPhoneNumbers

            $VoiceApps += $VoiceAppProperties

        }

        do {

            $VoiceAppSelection = $VoiceApps | Out-GridView -Title "Choose an Auto Attendant or Call Queue from the list." -PassThru

            if (!$VoiceAppSelection) {

                Write-Warning -Message "No voice app was selected. The script cannot work it's magic until you select a top-level voice app."

                do {

                    $abortPrompt = Read-Host -Prompt "Do you want to abort the script? [Y] = Yes or [N] = No"

                    if ($abortPrompt -ne "y" -and $abortPrompt -ne "n") {

                        Write-Warning "Invalid input. Please enter either [Y] = Yes or [N] = No"

                        $continueScript = $false

                    }

                    else {
                    
                        $continueScript = $true
                    
                    }

                } until (
                    $continueScript -eq $true
                )

                if ($abortPrompt -eq "y") {

                    Write-Host "Script aborted!" -ForegroundColor Red
                    exit

                }
            
            }

        } until (
            $VoiceAppSelection
        )

        if ($VoiceAppSelection.Type -eq "Auto Attendant") {

            $VoiceApp = $allAutoAttendants | Where-Object {$_.Name -eq $VoiceAppSelection.Name}
            $voiceAppType = "Auto Attendant"

        }

        else {

            $VoiceApp = $allCallQueues | Where-Object {$_.Name -eq $VoiceAppSelection.Name}
            $voiceAppType = "Call Queue"

        }


    }

    elseif ($VoiceAppId) {

        if ($allAutoAttendantIds -contains $VoiceAppId) {

            $VoiceApp = $allAutoAttendants | Where-Object {$_.Identity -eq $VoiceAppId}
            $voiceAppType = "Auto Attendant"

        }

        else {

            $VoiceApp = $allCallQueues | Where-Object {$_.Identity -eq $VoiceAppId}
            $voiceAppType = "Call Queue"

        }

    }

    else {

        if ($voiceAppType -eq "Auto Attendant") {

            $VoiceApp = $allAutoAttendants | Where-Object {$_.Name -eq $VoiceAppName}

        }

        else {

            $VoiceApp = $allCallQueues | Where-Object {$_.Name -eq $VoiceAppName}

        }

    }

    $mdNodePhoneNumbers = @()

    foreach ($ApplicationInstance in ($VoiceApp.ApplicationInstances)) {

        if ($mdNodePhoneNumbersCounter -eq 0) {

            $mdPhoneNumberLinkType = "-->"
            $VoiceAppFileName = Optimize-DisplayName -String $VoiceApp.Name

        }

        else {

            $mdPhoneNumberLinkType = "-.->"

        }

        $ApplicationInstancePhoneNumber = ($allResourceAccounts | Where-Object {$_.ObjectId -eq $ApplicationInstance}).PhoneNumber -replace ("tel:","")

        if ($ShowPhoneNumberType -eq $true) {

            $ApplicationInstancePhoneNumberType = (Get-CsPhoneNumberAssignment -TelephoneNumber $ApplicationInstancePhoneNumber).NumberType

            $ApplicationInstancePhoneNumberName = $ApplicationInstancePhoneNumberName + "<br>"

        }

        else {

            $ApplicationInstancePhoneNumberType = $null

        }

        if ($ApplicationInstancePhoneNumber) {

            if ($ApplicationInstancePhoneNumber -notmatch "\+") {
            
                $ApplicationInstancePhoneNumber = "+$ApplicationInstancePhoneNumber"
    
            }

            if ($ObfuscatePhoneNumbers -eq $true) {

                $ApplicationInstancePhoneNumberName = $ApplicationInstancePhoneNumber.Remove(($ApplicationInstancePhoneNumber.Length -4)) + "****"

            }

            else {

                $ApplicationInstancePhoneNumberName = $ApplicationInstancePhoneNumber
            
            }

            if ($ShowPhoneNumberType -eq $true) {
    
                $ApplicationInstancePhoneNumberName = $ApplicationInstancePhoneNumberName + "<br>$ApplicationInstancePhoneNumberType"
    
            }    

            $mdNodeNumber = "start$($ApplicationInstancePhoneNumber)((Incoming Call at <br> $ApplicationInstancePhoneNumberName)) $mdPhoneNumberLinkType $($VoiceApp.Identity)([$($voiceAppType) <br> $VoiceAppFileName])"

            $mdNodePhoneNumbers += $mdNodeNumber
    
            $mdNodePhoneNumbersCounter ++

            $allMermaidNodes += "start$($ApplicationInstancePhoneNumber)"

        }

        $mdNodePhoneNumbersCounter ++

    }

    # Normalize mermaid code for comparison
    $normalizedMermaidCode = $mermaidCode -replace "-\.->", "-->"

    # Loop through original lines
    foreach ($startPhoneNumber in @($mdNodePhoneNumbers)) {

        $normalizedLine = $startPhoneNumber.Replace("-.->", "-->")

        if ($normalizedMermaidCode -notcontains $normalizedLine) {

            $mermaidCode += $startPhoneNumber  # Append the original (with correct arrow type)

        }

    }

    if ($voiceAppType -eq "Auto Attendant") {
        . Find-Holidays -VoiceAppId $VoiceApp.Identity
        . Find-AfterHours -VoiceAppId $VoiceApp.Identity

        if ($aaHasHolidays -eq $true -and $aaHasAfterHours -eq $false) {
    
            . Get-AutoAttendantDefaultCallFlow -VoiceAppId $VoiceApp.Identity
        
            . Get-AutoAttendantHolidaysAndAfterHours -VoiceAppId $VoiceApp.Identity
    
        }
    
        elseif ($aaHasHolidays -eq $true -or $aaHasAfterHours -eq $true) {
    
            . Get-AutoAttendantDefaultCallFlow -VoiceAppId $VoiceApp.Identity
    
            . Get-AutoAttendantAfterHoursCallFlow -VoiceAppId $VoiceApp.Identity
    
            . Get-AutoAttendantHolidaysAndAfterHours -VoiceAppId $VoiceApp.Identity
    
        }

        else {
    
            . Get-AutoAttendantDefaultCallFlow -VoiceAppId $VoiceApp.Identity

            $aa.Name = Optimize-DisplayName -String $aa.Name

            $nodeElementHolidayLink = "$($aa.Identity)([Auto Attendant <br> $($aa.Name)])"
    
            $mdHolidayAndAfterHoursCheck =@"
            $nodeElementHolidayLink --> $mdAutoAttendantDefaultCallFlow
            
"@

            $allMermaidNodes += "$($aa.Identity)"

            if ($mermaidCode -notcontains $mdHolidayAndAfterHoursCheck) {

                $mermaidCode += $mdHolidayAndAfterHoursCheck

            }
    
        }
        
    }
    
    elseif ($voiceAppType -eq "Call Queue") {
        . Get-CallQueueCallFlow -MatchingCQIdentity $VoiceApp.Identity
    }

}

# Get First Call Flow

if ($Identity) {

    . Get-CallFlow -VoiceAppId $Identity

}

else {

    . Get-CallFlow -VoiceAppName $VoiceAppName -voiceAppType $VoiceAppType

}

function Get-NestedCallFlow {
    param (
    )

    foreach ($nestedVoiceApp in $nestedVoiceApps) {

        if ($processedVoiceApps -notcontains $nestedVoiceApp) {

            $processedVoiceApps += $nestedVoiceApp

            . Get-AccountType -Id $nestedVoiceApp

            Write-Host "Account Type: $accountType" -ForegroundColor Magenta

            if ($accountType -eq "VoiceApp") {

                . Get-CallFlow -VoiceAppId $nestedVoiceApp

            }

            if ($accountType -eq "UserAccount" -and $ShowUserCallingSettings) {

                . Get-TeamsUserCallFlow -UserId $nestedVoiceApp -PreviewSvg $false -SetClipBoard $false -StandAlone $false -ExportSvg $false -CustomFilePath $CustomFilePath -ObfuscatePhoneNumbers $ObfuscatePhoneNumbers

                if ($mermaidCode -notcontains $mdUserCallingSettings) {
    
                    $mermaidCode += $mdUserCallingSettings
        
                }    

            }


        }

    }

    if (Compare-Object -ReferenceObject $nestedVoiceApps -DifferenceObject $processedVoiceApps) {

        . Get-NestedCallFlow

    }

}

if ($ShowNestedCallFlows -eq $true) {

    . Get-NestedCallFlow

}

else {
    
    if ($nestedVoiceApps) {

        Write-Warning -Message "Your call flow contains nested call queues or auto attendants. They won't be expanded because 'ShowNestedCallFlows' is set to false."
        Write-Host "Nested Voice App Ids:" -ForegroundColor Yellow
        $nestedVoiceApps

    }

}

#Remove invalid characters from mermaid syntax
$mermaidCode = $mermaidCode.Replace(";",",")

#Add H2 Title to Markdown code
$mermaidCode = $mermaidCode.Replace("## CallFlowNamePlaceHolder","## $VoiceAppFileName")

if ($OverrideVoiceIdToFemale) {

    $mermaidCode = $mermaidCode.Replace("Male<br>","Female<br>")

}

# Remove duplicate nodes from the mermaid code to avoid 
$mermaidCode = $mermaidCode | Select-Object -Unique

# Custom Mermaid Color Themes
function Set-CustomMermaidTheme {
    param (
        [Parameter(Mandatory = $false)][String]$NodeColor,
        [Parameter(Mandatory = $false)][String]$NodeBorderColor,
        [Parameter(Mandatory = $false)][String]$FontColor,
        [Parameter(Mandatory = $false)][String]$LinkColor,
        [Parameter(Mandatory = $false)][String]$LinkTextColor
    )


    $themedNodes = "classDef customTheme fill:$NodeColor,stroke:$NodeBorderColor,stroke-width:2px,color:$FontColor`n`nclass "

    $allMermaidNodes = $allMermaidNodes | Sort-Object -Unique

    foreach ($node in $allMermaidNodes) {

        $themedNodes += "$node,"

    }

    $mermaidString = ($mermaidCode | Out-String)
    $NumberOfMermaidLinks = (Select-String -InputObject $mermaidString -Pattern '(--)|(-.-)|( -. )' -AllMatches).Matches.Count

    $themedNodes = ($themedNodes += " customTheme").Replace(", customTheme", " customTheme")

    $themedLinks = "`nlinkStyle "

    $currentMermaidLink = 0

    do {
        $themedLinks += "$currentMermaidLink,"

        $currentMermaidLink ++

    } until ($currentMermaidLink -eq ($NumberOfMermaidLinks))

    $themedLinks = ($themedLinks += " stroke:$LinkColor,stroke-width:2px,color:$LinkTextColor").Replace(", stroke:"," stroke:")

    if ($allSubgraphs) {

        $themedSubgraphs = "`nclassDef customSubgraphTheme fill:$SubgraphColor,color:$FontColor,stroke:$NodeBorderColor`n`nclass "

        foreach ($subgraph in $allSubgraphs) {
            
            $themedSubgraphs += "$subgraph,"

        }

        $themedSubgraphs = ($themedSubgraphs += " customSubgraphTheme").Replace(", customSubgraphTheme", " customSubgraphTheme")
    
    }

    else {

        $themedSubgraphs = $null

    }

    $mermaidCode += @($themedNodes,$themedLinks,$themedSubgraphs)

}

if ($Theme -eq "custom") {

    . Set-CustomMermaidTheme -NodeColor $NodeColor -NodeBorderColor $NodeBorderColor -FontColor $FontColor -LinkColor $LinkColor -LinkTextColor $LinkTextColor

}



if ($SaveToFile -eq $true) {

    if ($ExportAudioFiles -and $audioFileNames) {

        if ($CustomFilePath) {

            $audioFileNames = $audioFileNames.Replace($CustomFilePath,".")

        }
        
        $mermaidCode += $audioFileNames

    }

    if ($ExportTTSGreetings -and $ttsGreetings) {

        if ($CustomFilePath) {

            $ttsGreetings = $ttsGreetings.Replace($CustomFilePath,".")

        }


        $mermaidCode += $ttsGreetings

    }

    $mermaidCode += $mdEnd

    if ($ExportAudioFiles -eq $true -or $ExportTTSGreetings -eq $true -and $DocFxMode -eq $true) {

        $docFxFriendlyMermaidCode = $mermaidCode.Replace('".\',"`"$voiceAppIdentity\")

        Set-Content -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow$fileExtension" -Value $docFxFriendlyMermaidCode -Encoding UTF8

    }

    else {

        Set-Content -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow$fileExtension" -Value $mermaidCode -Encoding UTF8

    }

}

if ($ExportPng -eq $true) {

    if ($Theme -eq "custom") {

        $pngTheme = "dark"

    }

    else {

        $pngTheme = $Theme

    }

    if (Test-Path -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.png") {

        Remove-Item -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.png" -Force

    }

    mmdc -i "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow$fileExtension" -o "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.png" -b transparent -t "$pngTheme" -s 5 --configFile=".\mermaidRenderConfig.json"

    if ($DocType -eq "Markdown") {

        $createdPng = Get-ChildItem -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow-1.png"
        Rename-Item -Path $createdPng.FullName -NewName "$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.png" -Force

    }

}

if ($ExportPDF -eq $true) {

    if ($Theme -eq "custom") {

        $pdfTheme = "dark"

    }

    else {

        $pdfTheme = $Theme

    }

    if (Test-Path -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.pdf") {

        Remove-Item -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.pdf" -Force

    }

    mmdc -i "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow$fileExtension" -o "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.pdf" -b transparent "$pdfTheme" -s 10 --pdfFit

    if ($DocType -eq "Markdown") {

        $createdPDF = Get-ChildItem -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow-1.pdf"
        Rename-Item -Path $createdPDF.FullName -NewName "$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.pdf" -Force

    }

}

if ($SetClipBoard -eq $true) {
    $mermaidCode -Replace('```mermaid','') `
    -Replace('```','') `
    -Replace("## $VoiceAppFileName","") `
    -Replace($MarkdownTheme,"") | Set-Clipboard

    Write-Host "Mermaid code copied to clipboard. Paste it on https://mermaid.live" -ForegroundColor Cyan
}

if ($ExportHtml -eq $true) {

    $HtmlOutput = Get-Content -Path .\HtmlTemplate.html | Out-String

    if ($Theme -eq "custom") {

        $MarkdownTheme = '<div class="mermaid">'
        $MarkdownThemeHtml = '<div class="mermaid">'
        
    }

    else {

        $MarkdownThemeHtml = '<div class="mermaid">' + $MarkdownTheme 

    }


    if ($DocType -eq "Markdown") {

        $HtmlOutput -Replace "VoiceAppNamePlaceHolder","Call Flow $VoiceAppFileName" `
        -Replace "VoiceAppNameHtmlIdPlaceHolder",($($VoiceAppFileName).Replace(" ","-")) `
        -Replace '<div class="mermaid">ThemePlaceHolder',$MarkdownThemeHtml `
        -Replace "MermaidPlaceHolder",($mermaidCode | Out-String).Replace($MarkdownTheme,"") `
        -Replace "## $VoiceAppFileName","" `
        -Replace('```mermaid','') `
        -Replace('```','') | Set-Content -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.htm" -Encoding UTF8

    }

    else {

        $HtmlOutput -Replace "VoiceAppNamePlaceHolder","Call Flow $VoiceAppFileName" `
        -Replace "VoiceAppNameHtmlIdPlaceHolder",($($VoiceAppFileName).Replace(" ","-")) `
        -Replace '<div class="mermaid">ThemePlaceHolder',$MarkdownThemeHtml `
        -Replace "MermaidPlaceHolder",($mermaidCode | Out-String).Replace($MarkdownTheme,"") | Set-Content -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.htm" -Encoding UTF8

    }

    if ($PreviewHtml) {

        Start-Process "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.htm"

    }

}

if ($CheckCallFlowRouting -eq $true) {

    Write-Host "Check Call Flow Routing Results for $($VoiceAppFileName)" -ForegroundColor Cyan

    $allCallFlowRoutingChecks | Format-Table -AutoSize

    $allCallFlowRoutingChecks | Out-GridView -Title "Check Call Flow Routing Results for $($VoiceAppFileName)"

}
