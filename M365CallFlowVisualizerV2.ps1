<#
    .SYNOPSIS
    Reads the configuration from a Microsoft 365 Phone System auto attendant or call queue and visualizes the call flow using mermaid-js.

    .DESCRIPTION
    Reads the configuration from a Microsoft 365 Phone System auto attendant or call queue by either specifying the voice app name and type, unique identity of the voice app or presents a selection of available auto attendants or call queues if none of the identifiers are supplied.
    The call flow is then written into either a mermaid (*.mmd) or a markdown (*.md) file containing the mermaid syntax.

    Author:             Martin Heusser
    Version:            2.8.7b
    Changelog:          Moved to repository at .\Changelog.md

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

    -PreviewHtml
        Specifies if the exported html file should be opened in default / last active browser (only works on Windows systems)
        Required:           false
        Type:               switch
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

    -ShowUserCallingSettings
        Specifies whether or not to also display the user calling settings of a Teams user. If set to false, only the name of a user will be rendered.
        Required:           false
        Type:               boolean
        Default value:      true

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
        Specifies if the audio files of greetings, announcements and music on hold should be exported to the specified directory. If this is enabled, Markdown and HTML output will have clickable links on the greeting nodes which open an audio file in the browser. This is an experimental feature.
        Required:           false
        Type:               switch
        Default value:      false

    -ExportTTSGreetings
        Specifies if the value of TTS greetings and announcements should be exported to the specified directory. If this is enabled, Markdown and HTML output will have clickable links on the greeting nodes with which open a text file in the browser. This is an experimental feature.
        Required:           false
        Type:               switch
        Default value:      false

    -FindUserLinks
        This paramter can only be used by the external function Find-CallQueueAndAutoAttendantUserLinks. When this parameter is true, the script will write every user which is part of a call flow into an external variable.
        Required:           false
        Type:               switch
        Default value:      false

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

#Requires -Modules @{ ModuleName = "MicrosoftTeams"; ModuleVersion = "4.6.0" }, "Microsoft.Graph.Users", "Microsoft.Graph.Groups"

[CmdletBinding(DefaultParametersetName="None")]
param(
    [Parameter(Mandatory=$false)][String]$Identity,
    [Parameter(Mandatory=$false)][Bool]$SetClipBoard = $false,
    [Parameter(Mandatory=$false)][Bool]$SaveToFile = $true,
    [Parameter(Mandatory=$false)][Bool]$ExportHtml = $true,
    [Parameter(Mandatory=$false)][Bool]$ExportPng = $false,
    [Parameter(Mandatory=$false)][Switch]$PreviewHtml,
    [Parameter(Mandatory=$false)][String]$CustomFilePath = ".\Output\$(Get-Date -Format "yyyy-MM-dd")",
    [Parameter(Mandatory=$false)][Bool]$ShowNestedCallFlows = $true,
    [Parameter(Mandatory=$false)][Bool]$ShowUserCallingSettings = $true,
    [Parameter(Mandatory=$false)][Switch]$ShowCqAgentPhoneNumbers,
    [Parameter(Mandatory=$false)][Switch]$ShowCqAgentOptInStatus,
    [Parameter(Mandatory=$false)][Switch]$ShowTTSGreetingText,
    [Parameter(Mandatory=$false)][Switch]$ShowAudioFileName,
    [Parameter(Mandatory=$false)][Single]$TruncateGreetings = 20,
    [Parameter(Mandatory=$false)][Switch]$ExportAudioFiles,
    [Parameter(Mandatory=$false)][Switch]$ExportTTSGreetings,
    [Parameter(Mandatory=$false)][Switch]$FindUserLinks,
    [Parameter(Mandatory=$false)][Bool]$ObfuscatePhoneNumbers = $false,
    [Parameter(Mandatory=$false)][Bool]$ShowSharedVoicemailGroupMembers = $false,
    [Parameter(Mandatory=$false)][ValidateSet("Markdown","Mermaid")][String]$DocType = "Markdown",
    [Parameter(Mandatory=$false)][ValidateSet("EU","US")][String]$DateFormat = "EU",
    [Parameter(Mandatory=$false)][ValidateSet("default","forest","dark","neutral","custom")][String]$Theme = "default",
    [Parameter(Mandatory=$false)][String]$NodeColor = "#505AC9",
    [Parameter(Mandatory=$false)][String]$NodeBorderColor = "#464EB8",
    [Parameter(Mandatory=$false)][String]$FontColor = "#FFFFFF",
    [Parameter(Mandatory=$false)][String]$LinkColor = "#505AC9",
    [Parameter(Mandatory=$false)][String]$LinkTextColor = "#000000",
    [Parameter(Mandatory=$false)][String]$SubgraphColor = "#7B83EB",
    [Parameter(ParameterSetName="VoiceAppProperties",Mandatory=$false)][String]$VoiceAppName,
    [Parameter(ParameterSetName="VoiceAppProperties",Mandatory=$true)][ValidateSet("Auto Attendant","Call Queue")][String]$VoiceAppType
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

# Connect to MicrosoftTeams and Microsoft.Graph
. Connect-M365CFV

if ($SaveToFile -eq $false -and $CustomFilePath -ne ".\Output") {

    Write-Warning -Message "Custom file path is specified but SaveToFile is set to false. The call flow won't be saved!"

}

if ($ObfuscatePhoneNumbers -eq $true) {

    Write-Warning -Message "Obfuscate phone numbers is True. This will only obfuscate phone numbers in node descriptions. The complete phone number will still be included in Markdown, Mermaid and HTML output!"

}

if ($ExportPng -eq $true) {

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
    }

}

# Common arrays and variables
$nestedVoiceApps = @()
$processedVoiceApps = @()
$allMermaidNodes = @()
$allSubgraphs = @()
$audioFileNames = @()
$ttsGreetings = @()

Write-Host "Retrieving all Auto Attendants (max. 1000)... this can take a while..." -ForegroundColor Magenta
if ([String]::IsNullOrEmpty($Global:allAutoAttendants)) {$Global:allAutoAttendants = Get-CsAutoAttendant -First 1000}
Write-Host "Retrieving all Call Queues (max. 1000)... this can take a while..." -ForegroundColor Magenta
if ([String]::IsNullOrEmpty($Global:allCallQueues)) {$Global:allCallQueues = Get-CsCallQueue -WarningAction SilentlyContinue -First 1000}

$allAutoAttendantIds = $allAutoAttendants.Identity
$allCallQueueIds = $allCallQueues.Identity
Write-Host "Retrieving all Resource Accounts (max. 1000)... this can take a while..." -ForegroundColor Magenta
if ([String]::IsNullOrEmpty($Global:allResourceAccounts)) {$Global:allResourceAccounts = Get-CsOnlineApplicationInstance -ResultSize 1000}

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
        [Parameter(Mandatory=$true)][String]$DocType
        )

    if ($Theme -eq "custom") {

        $MarkdownTheme = ""

    }

    else {

        $Theme = $Theme.ToLower()

        $MarkdownTheme =@"
%%{init: {'theme': '$($Theme)', 'flowchart' : { 'curve' : 'basis' } } }%%

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
        [Parameter(Mandatory=$true)][String]$VoiceAppId

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

                    $MatchingOperatorAa = $allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $Operator.Id}

                    $OperatorTypeFriendly = "[Auto Attendant"
                    $OperatorName = "$($MatchingOperatorAa.Name)]"
                    $OperatorIdentity = $MatchingOperatorAa.Identity

                }

                else {

                    $MatchingOperatorCq = $allCallQueues | Where-Object {$_.ApplicationInstances -contains $Operator.Id}

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
        [Parameter(Mandatory=$true)][String]$VoiceAppId

    )

    $aa = $allAutoAttendants | Where-Object {$_.Identity -eq $VoiceAppId}

    Write-Host "Reading call flow for: $($aa.Name)" -ForegroundColor Magenta
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
    
        # Check if the auto attendant has business hours by comparing the ps object to the actual config of the current auto attendant
        if ($aaDefaultScheduleProperties -eq ($aaEffectiveScheduleProperties | Out-String)) {
            $aaHasAfterHours = $false
        }

        else {
            $aaHasAfterHours = $true
        }

    }

    else {

        $aaHasAfterHours = $false

    }
    
}

function Get-AutoAttendantHolidaysAndAfterHours {
    param (
        [Parameter(Mandatory=$true)][String]$VoiceAppId
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

        foreach ($HolidayCallHandling in $aaHolidays) {

            $holidayCallFlow = $aa.CallFlows | Where-Object {$_.Id -eq $HolidayCallHandling.CallFlowId}
            $holidaySchedule = $aa.Schedules | Where-Object {$_.Id -eq $HolidayCallHandling.ScheduleId}

            if (!$holidayCallFlow.Greetings) {

                $holidayGreeting = "Greeting <br> None"

            }

            else {

                $holidayGreeting = "Greeting <br> $($holidayCallFlow.Greetings.ActiveType)"

                if ($($holidayCallFlow.Greetings.ActiveType) -eq "TextToSpeech" -and $ShowTTSGreetingText) {

                    $audioFileName = $null

                    # Optmimize-DisplayName for TTS Greetings (needs to be added for every greeting...)
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

            $holidayAction = $holidayCallFlow.Menu.MenuOptions.Action

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
                        
                        $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $holidayCallFlow.Menu.MenuOptions.CallTarget.Id -and $_.ApplicationId -eq $applicationIdAa}

                        if ($matchingApplicationInstanceCheckAa) {

                            $MatchingAA = $allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $holidayCallFlow.Menu.MenuOptions.CallTarget.Id}

                            $holidayActionTargetTypeFriendly = "[Auto Attendant"
                            $holidayActionTargetName = (Optimize-DisplayName -String $($MatchingAA.Name)) + "]"

                        }

                        else {

                            $MatchingCQ = $allCallQueues | Where-Object {$_.ApplicationInstances -contains $holidayCallFlow.Menu.MenuOptions.CallTarget.Id}

                            $holidayActionTargetTypeFriendly = "[Call Queue"
                            $holidayActionTargetName = (Optimize-DisplayName -String $($MatchingCQ.Name)) + "]"

                        }

                        $holidayVoicemailSystemGreeting = $null

                    }
                
                }

                # Create mermaid code for the holiday action node based on the variables created in the switch statemenet
                $nodeElementHolidayAction = "elementAAHolidayAction$($aaObjectId)-$($HolidayCounter)($holidayAction) $holidayVoicemailSystemGreeting--> elementAAHolidayActionTargetType$($aaObjectId)-$($HolidayCounter)($holidayActionTargetTypeFriendly <br> $holidayActionTargetName)"

                $allMermaidNodes += @("elementAAHolidayAction$($aaObjectId)-$($HolidayCounter)","elementAAHolidayActionTargetType$($aaObjectId)-$($HolidayCounter)")

            }

            # Create subgraph per holiday call handling inside the Holidays subgraph

            $mermaidFriendlyHolidayName = Optimize-DisplayName -String $($holidayCallFlow.Name)

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

            # Increase the counter by 1
            $HolidayCounter ++

            # Add holiday call handling subgraph to holiday subgraph
            $mdSubGraphHolidays += $nodeElementHolidayDetails

            $allSubgraphs += @("subgraphHolidayCallHandling$($HolidayCallHandling.CallFlowId)","subgraphHolidaySchedule$($HolidayCallHandling.CallFlowId)")

        } # End of for-each loop

        # Create end for the holiday subgraph
        $mdSubGraphHolidaysEnd =@"

    end
"@
            
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
$nodeElementAfterHoursCheck -->|No| $mdAutoAttendantAfterHoursCallFlow
$nodeElementAfterHoursCheck -->|Yes| $mdAutoAttendantDefaultCallFlow

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
$nodeElementAfterHoursCheck -->|No| $mdAutoAttendantAfterHoursCallFlow
$nodeElementAfterHoursCheck -->|Yes| $mdAutoAttendantDefaultCallFlow


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

    
}

function Get-AutoAttendantDefaultCallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceAppId
    )

    $aaDefaultCallFlowAaObjectId = $aa.Identity

    $languageId = $aa.LanguageId

    # Get the current auto attendants default call flow and default call flow action
    $defaultCallFlow = $aa.DefaultCallFlow
    $defaultCallFlowAction = $aa.DefaultCallFlow.Menu.MenuOptions.Action

    # Get the current auto attentans default call flow greeting
    if (!$defaultCallFlow.Greetings.ActiveType){
        $defaultCallFlowGreeting = "Greeting <br> None"
    }

    else {

        $defaultCallFlowGreeting = "Greeting <br> $($defaultCallFlow.Greetings.ActiveType)"

        if ($($defaultCallFlow.Greetings.ActiveType) -eq "TextToSpeech" -and $ShowTTSGreetingText) {

            $audioFileName = $null

            $defaultTTSGreetingValue = $defaultCallFlow.Greetings.TextToSpeechPrompt

            if ($ExportTTSGreetings) {

                $defaultTTSGreetingValue | Out-File "$FilePath\$($aaDefaultCallFlowAaObjectId)_defaultCallFlowGreeting.txt"

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

        $mdAutoAttendantDefaultCallFlow = "defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowGreeting] --> defaultCallFlow$($aaDefaultCallFlowAaObjectId)(($defaultCallFlowAction))`n"

        $allMermaidNodes += @("defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)","defaultCallFlow$($aaDefaultCallFlowAaObjectId)")

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
    
                $defaultCallFlowMenuOptionsTTSGreetingValue = $defaultCallFlow.Menu.Prompts.TextToSpeechPrompt

                if ($ExportTTSGreetings) {

                    $defaultCallFlowMenuOptionsTTSGreetingValue | Out-File "$FilePath\$($aaDefaultCallFlowAaObjectId)_defaultCallFlowMenuOptionsGreeting.txt"
    
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

                $defaultCallFlowVoiceResponse = " or <br> Voice Response <br> Language: $($aa.LanguageId)$defaultCallFlowForceListen"


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
        
                        $allMermaidNodes += "defaultCallFlowTransferPSTNGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowPSTNOptionCounter"

                        $defaultCallFlowPSTNOptionCounter ++

                    }
                    ApplicationEndpoint {

                        # Check if application endpoint is auto attendant or call queue

                        $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $MenuOption.CallTarget.Id -and $_.ApplicationId -eq $applicationIdAa}

                        if ($matchingApplicationInstanceCheckAa) {

                            $MatchingAaDefaultCallFlowAa = $allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id}

                            $defaultCallFlowTargetTypeFriendly = "[Auto Attendant"
                            $defaultCallFlowTargetName = (Optimize-DisplayName -String $($MatchingAaDefaultCallFlowAa.Name)) + "]"

                        }

                        else {

                            $MatchingCqAaDefaultCallFlow = $allCallQueues | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id}

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
        
                        $allMermaidNodes += "defaultCallFlowTransferSharedVoicemailGreeting$($aadefaultCallFlowAaObjectId)-$defaultCallFlowSharedVoicemailCounter"

                        $allMermaidNodes += "defaultCallFlowSystemGreeting$($aaDefaultCallFlowAaObjectId)-$defaultCallFlowSharedVoicemailCounter"

                        $defaultCallFlowSharedVoicemailCounter ++

                    }
                }

                # Check if transfer target type is call queue
                if ($defaultCallFlowTargetTypeFriendly -eq "[Call Queue") {

                    $MatchingCQIdentity = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id}).Identity

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

            # Remove Greeting node, if none is configured
            if ($defaultCallFlowGreeting -eq "Greeting <br> None") {

                $mdAutoAttendantDefaultCallFlow = ($mdAutoAttendantDefaultCallFlow.Replace("defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowGreeting] --> ","")).TrimStart()

            }
    
    }
    
    
}

function Get-AutoAttendantAfterHoursCallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceAppId
    )

    $aaAfterHoursCallFlowAaObjectId = $aa.Identity

    $languageId = $aa.LanguageId

    # Get after hours call flow
    $afterHoursAssociatedCallFlowId = ($aa.CallHandlingAssociations | Where-Object {$_.Type -eq "AfterHours"}).CallFlowId
    $afterHoursCallFlow = ($aa.CallFlows | Where-Object {$_.Id -eq $afterHoursAssociatedCallFlowId})
    $afterHoursCallFlowAction = ($aa.CallFlows | Where-Object {$_.Id -eq $afterHoursAssociatedCallFlowId}).Menu.MenuOptions.Action

    # Get after hours greeting
    if (!$afterHoursCallFlow.Greetings.ActiveType){
        $afterHoursCallFlowGreeting = "Greeting <br> None"
    }

    else {
        $afterHoursCallFlowGreeting = "Greeting <br> $($afterHoursCallFlow.Greetings.ActiveType)"

        if ($($afterHoursCallFlow.Greetings.ActiveType) -eq "TextToSpeech" -and $ShowTTSGreetingText) {

            $audioFileName = $null

            $afterHoursTTSGreetingValue = $afterHoursCallFlow.Greetings.TextToSpeechPrompt

            if ($ExportTTSGreetings) {

                $afterHoursTTSGreetingValue | Out-File "$FilePath\$($aaDefaultCallFlowAaObjectId)_afterHoursCallFlowGreeting.txt"

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

        $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId)>$AfterHoursCallFlowGreeting] --> afterHoursCallFlow$($aaAfterHoursCallFlowAaObjectId)(($afterHoursCallFlowAction))`n"

        $allMermaidNodes += @("afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId)","afterHoursCallFlow$($aaAfterHoursCallFlowAaObjectId)")

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
    
                $afterHoursCallFlowMenuOptionsTTSGreetingValue = $afterHoursCallFlow.Menu.Prompts.TextToSpeechPrompt

                if ($ExportTTSGreetings) {

                    $afterHoursCallFlowMenuOptionsTTSGreetingValue | Out-File "$FilePath\$($aaafterHoursCallFlowAaObjectId)_afterHoursCallFlowMenuOptionsGreeting.txt"
    
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

                $afterHoursCallFlowVoiceResponse = " or <br> Voice Response <br> Language: $($aa.LanguageId)$afterHoursCallFlowForceListen"


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

            if ($afterHoursCallFlowMenuOptions.Count -lt 2 -and !$afterHoursCallFlow.Menu.Prompts.ActiveType) {

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
        
                    $afterHoursCallFlowMenuOptionsTTSAnnouncementValue = $MenuOption.Prompt.TextToSpeechPrompt

                    if ($ExportTTSGreetings) {

                        $afterHoursCallFlowMenuOptionsTTSAnnouncementValue | Out-File "$FilePath\$($aaafterHoursCallFlowAaObjectId)_afterHoursCallFlowMenuOptionsAnnouncement$DtmfKey.txt"
        
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

                        $allMermaidNodes += "afterHoursCallFlowTransferPSTNGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowPSTNOptionCounter"

                        $afterHoursCallFlowPSTNOptionCounter ++

                    }
                    ApplicationEndpoint {

                        # Check if application endpoint is auto attendant or call queue

                        $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $MenuOption.CallTarget.Id -and $_.ApplicationId -eq $applicationIdAa}

                        if ($matchingApplicationInstanceCheckAa) {

                            $MatchingAaafterHoursCallFlowAa = $allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id}

                            $afterHoursCallFlowTargetTypeFriendly = "[Auto Attendant"
                            $afterHoursCallFlowTargetName = (Optimize-DisplayName -String $($MatchingAaafterHoursCallFlowAa.Name)) + "]"

                        }

                        else {

                            $MatchingCqAaafterHoursCallFlow = $allCallQueues | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id}

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
        
                        $allMermaidNodes += "afterHoursCallFlowTransferSharedVoicemailGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowSharedVoicemailCounter"

                        $allMermaidNodes += "afterHoursCallFlowSystemGreeting$($aaafterHoursCallFlowAaObjectId)-$afterHoursCallFlowSharedVoicemailCounter"

                        $afterHoursCallFlowSharedVoicemailCounter ++

                    }
                }

                # Check if transfer target type is call queue
                if ($afterHoursCallFlowTargetTypeFriendly -eq "[Call Queue") {

                    $MatchingCQIdentity = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id}).Identity

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

        # Remove Greeting node, if none is configured
        if ($afterHoursCallFlowGreeting -eq "Greeting <br> None") {

            $mdAutoAttendantafterHoursCallFlow = ($mdAutoAttendantafterHoursCallFlow.Replace("afterHoursCallFlowGreeting$($aaafterHoursCallFlowAaObjectId)>$afterHoursCallFlowGreeting] --> ","")).TrimStart()

        }
        
    
    }


}


function Get-CallQueueCallFlow {
    param (
        [Parameter(Mandatory=$true)][String]$MatchingCQIdentity
    )

    $MatchingCQ = $allCallQueues | Where-Object {$_.Identity -eq $MatchingCQIdentity}

    $cqCallFlowObjectId = $MatchingCQ.Identity

    Write-Host "Reading call flow for: $($MatchingCQ.Name)" -ForegroundColor Magenta
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
    $CqPresenceBasedRouting = $MatchingCQ.PresenceBasedRouting
    $CqDistributionLists = $MatchingCQ.DistributionLists
    $CqDefaultMusicOnHold = $MatchingCQ.UseDefaultMusicOnHold
    $CqWelcomeMusicFileName = $MatchingCQ.WelcomeMusicFileName
    $CqLanguageId = $MatchingCQ.LanguageId
    
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
    if (!$CqWelcomeMusicFileName) {
        $CqGreeting = "None"

        $cqGreetingNode = $null
    }

    else {
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

            $TeamName = (Get-Team -GroupId $MatchingCQ.DistributionLists.Guid).DisplayName
            $ChannelName = Optimize-DisplayName -String (Get-TeamChannel -GroupId $MatchingCQ.DistributionLists.Guid | Where-Object {$_.Id -eq $MatchingCQ.ChannelId}).DisplayName

            $CqAgentListType = "Teams Channel <br> Team Name: $TeamName <br> Channel Name: $ChannelName"

        }

    }

    # Switch through call queue overflow action target
    switch ($CqOverFlowAction) {
        DisconnectWithBusy {
            $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)((DisconnectCall))"

            $allMermaidNodes += "cqOverFlowAction$($cqCallFlowObjectId)"

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
                
            }

            else {

                $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $MatchingCQ.OverflowActionTarget.Id -and $_.ApplicationId -eq $applicationIdAa}

                if ($matchingApplicationInstanceCheckAa) {

                    $MatchingOverFlowAA = ($allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.OverflowActionTarget.Id})

                    $MatchingOverFlowAA.Name = Optimize-DisplayName -String $MatchingOverFlowAA.Name
                    $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingOverFlowAA.Identity)([Auto Attendant <br> $($MatchingOverFlowAA.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingOverFlowAA.Identity -and $MatchingCQ.TimeoutThreshold -ge 1) {

                        $nestedVoiceApps += $MatchingOverFlowAA.Identity
        
                    }

                    $allMermaidNodes += @("cqOverFlowAction$($cqCallFlowObjectId)","$($MatchingOverFlowAA.Identity)")
        

                }

                else {

                    $MatchingOverFlowCQ = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.OverflowActionTarget.Id})

                    $MatchingOverFlowCQ.Name = Optimize-DisplayName -String $MatchingOverFlowCQ.Name

                    $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingOverFlowCQ.Identity)([Call Queue <br> $($MatchingOverFlowCQ.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingOverFlowCQ.Identity -and $MatchingCQ.TimeoutThreshold -ge 1) {

                        $nestedVoiceApps += $MatchingOverFlowCQ.Identity
        
                    }

                    $allMermaidNodes += @("cqOverFlowAction$($cqCallFlowObjectId)","$($MatchingOverFlowCQ.Identity)")

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

                    $overFlowVoicemailTTSGreetingValue = $MatchingCQ.OverflowSharedVoicemailTextToSpeechPrompt

                    if ($ExportTTSGreetings) {

                        $overFlowVoicemailTTSGreetingValue | Out-File "$FilePath\$($cqCallFlowObjectId)_cqOverFlowVoicemailGreeting.txt"
        
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
            $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)((DisconnectCall))"

            $allMermaidNodes += "cqTimeoutAction$($cqCallFlowObjectId)"

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
                
            }
    
            else {

                $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $MatchingCQ.TimeoutActionTarget.Id -and $_.ApplicationId -eq $applicationIdAa}
        
                if ($matchingApplicationInstanceCheckAa) {

                    $MatchingTimeoutAA = ($allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.TimeoutActionTarget.Id})

                    $MatchingTimeoutAA.Name = Optimize-DisplayName -String $MatchingTimeoutAA.Name
    
                    $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingTimeoutAA.Identity)([Auto Attendant <br> $($MatchingTimeoutAA.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingTimeoutAA.Identity -and $MatchingCQ.OverflowThreshold -ge 1) {

                        $nestedVoiceApps += $MatchingTimeoutAA.Identity
        
                    }

                    $allMermaidNodes += @("cqTimeoutAction$($cqCallFlowObjectId)","$($MatchingTimeoutAA.Identity)")
    
                }
    
                else {
    
                    $MatchingTimeoutCQ = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.TimeoutActionTarget.Id})

                    $MatchingTimeoutCQ.Name = Optimize-DisplayName -String $MatchingTimeoutCQ.Name

                    $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingTimeoutCQ.Identity)([Call Queue <br> $($MatchingTimeoutCQ.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingTimeoutCQ.Identity -and $MatchingCQ.OverflowThreshold -ge 1) {

                        $nestedVoiceApps += $MatchingTimeoutCQ.Identity
        
                    }

                    $allMermaidNodes += @("cqTimeoutAction$($cqCallFlowObjectId)","$($MatchingTimeoutCQ.Identity)")
    
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

                    $TimeOutVoicemailTTSGreetingValue = $MatchingCQ.TimeOutSharedVoicemailTextToSpeechPrompt

                    if ($ExportTTSGreetings) {

                        $TimeOutVoicemailTTSGreetingValue | Out-File "$FilePath\$($cqCallFlowObjectId)_cqTimeoutVoicemailGreeting.txt"
        
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

            $CqAgentPhoneNumber = ((Get-CsOnlineUser -Identity $($CqAgent.ObjectId)).LineUri)

            if (!$CqAgentPhoneNumber) {

                $CqAgentPhoneNumber = "No Number Assigned"

            }

            else {

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

            $AgentDisplayName = "$AgentDisplayName <br> $CqAgentPhoneNumber"

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

    
    # Create default callflow mermaid code

$mdCallQueueCallFlow =@"
$($MatchingCQIdentity)([Call Queue <br> $($CqName)]) -->$cqGreetingNode overFlow$($cqCallFlowObjectId){More than $CqOverFlowThreshold <br> Active Calls?}
overFlow$($cqCallFlowObjectId) --> |Yes| $CqOverFlowActionFriendly
overFlow$($cqCallFlowObjectId) ---> |No| routingMethod$($cqCallFlowObjectId)

subgraph subgraphCallDistribution$($cqCallFlowObjectId)[Call Distribution: $($MatchingCQ.Name)]
subgraph subgraphCqSettings$($cqCallFlowObjectId)[CQ Settings]
routingMethod$($cqCallFlowObjectId)[(Routing Method: $CqRoutingMethod)] --> agentAlertTime$($cqCallFlowObjectId)
agentAlertTime$($cqCallFlowObjectId)[(Agent Alert Time: $CqAgentAlertTime)] -.- cqSettingsContainer$($cqCallFlowObjectId)
cqSettingsContainer$($cqCallFlowObjectId)[(Music On Hold: $CqMusicOnHold <br> Conference Mode Enabled: $CqConferenceMode <br> Agent Opt Out Allowed: $CqAgentOptOut <br> Presence Based Routing: $CqPresenceBasedRouting <br> TTS Greeting Language: $CqLanguageId)] -.- timeOut$($cqCallFlowObjectId)
timeOut$($cqCallFlowObjectId)[(Timeout: $CqTimeOut Seconds)]
end
agentAlertTime$($cqCallFlowObjectId) --> subgraphAgents$($cqCallFlowObjectId)
subgraph subgraphAgents$($cqCallFlowObjectId)[Agents List]
agentListType$($cqCallFlowObjectId)[(Agent List Type: $CqAgentListType)]
$mdCqAgentsDisplayNames
end
subgraphAgents$($cqCallFlowObjectId) --> cqResult$($cqCallFlowObjectId){Agent Answered?}
end

cqResult$($cqCallFlowObjectId) --> |Yes| cqEnd$($cqCallFlowObjectId)((Call Connected))
cqResult$($cqCallFlowObjectId) --> |No| timeOut$($cqCallFlowObjectId) --> $CqTimeoutActionFriendly

"@

    if ($mermaidCode -notcontains $mdCallQueueCallFlow) {

        if ($MatchingCQ.OverflowThreshold -ge 1 -and $MatchingCQ.TimeoutThreshold -ge 1) {

            $mermaidCode += $mdCallQueueCallFlow

        }

        else {

            if ($MatchingCQ.OverflowThreshold -eq 0) {

                $mdCallQueueCallFlow =@"
$($MatchingCQIdentity)([Call Queue <br> $($CqName)]) -->$cqGreetingNode overFlow$($cqCallFlowObjectId)[(Overflow Threshold: $CqOverFlowThreshold <br> Immediate Overflow Action <br> TTS Greeting Language: $CqLanguageId)]
overFlow$($cqCallFlowObjectId) --> $CqOverFlowActionFriendly

"@

            $mermaidCode += $mdCallQueueCallFlow
            }

            else {
                $mdCallQueueCallFlow =@"
$($MatchingCQIdentity)([Call Queue <br> $($CqName)]) -->$cqGreetingNode timeOut$($cqCallFlowObjectId)[(Timeout Threshold: $CqTimeOut <br> Immediate Timeout Action <br> TTS Greeting Language: $CqLanguageId)]
timeOut$($cqCallFlowObjectId) --> $CqTimeoutActionFriendly

"@
            $mermaidCode += $mdCallQueueCallFlow
            }
        }

        $allMermaidNodes += @("cqGreeting$($cqCallFlowObjectId)","overFlow$($cqCallFlowObjectId)","routingMethod$($cqCallFlowObjectId)","agentAlertTime$($cqCallFlowObjectId)","cqSettingsContainer$($cqCallFlowObjectId)","timeOut$($cqCallFlowObjectId)","agentListType$($cqCallFlowObjectId)","cqResult$($cqCallFlowObjectId)","cqEnd$($cqCallFlowObjectId)")
        $allSubgraphs += @("subgraphCallDistribution$($cqCallFlowObjectId)","subgraphCqSettings$($cqCallFlowObjectId)","subgraphAgents$($cqCallFlowObjectId)")

    }

}

. Set-Mermaid -DocType $DocType

#This is needed to determine if the Get-CallFlow function is running for the first time or not.
$mdNodePhoneNumbersCounter = 0

function Get-CallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceAppId,
        [Parameter(Mandatory=$false)][String]$VoiceAppName,
        [Parameter(Mandatory=$false)][String]$voiceAppType
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

                $ResourceAccountPhoneNumber = ((Get-CsOnlineApplicationInstance -Identity $ResourceAccount).PhoneNumber)

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

                $ResourceAccountPhoneNumber = ((Get-CsOnlineApplicationInstance -Identity $ResourceAccount).PhoneNumber)

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

        $VoiceAppSelection = $VoiceApps | Out-GridView -Title "Choose an Auto Attendant or Call Queue from the list." -PassThru

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

        $ApplicationInstancePhoneNumber = ((Get-CsOnlineApplicationInstance -Identity $ApplicationInstance).PhoneNumber) -replace ("tel:","")

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

            $mdNodeNumber = "start$($ApplicationInstancePhoneNumber)((Incoming Call at <br> $ApplicationInstancePhoneNumberName)) $mdPhoneNumberLinkType $($VoiceApp.Identity)([$($voiceAppType) <br> $VoiceAppFileName])"

            $mdNodePhoneNumbers += $mdNodeNumber
    
            $mdNodePhoneNumbersCounter ++

            $allMermaidNodes += "start$($ApplicationInstancePhoneNumber)"

        }

        $mdNodePhoneNumbersCounter ++

    }

    if ($mermaidCode -notcontains $mdNodePhoneNumbers) {

        $mermaidCode += $mdNodePhoneNumbers

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

#Add H1 Title to Markdown code
$mermaidCode = $mermaidCode.Replace("## CallFlowNamePlaceHolder","# Call Flow $VoiceAppFileName")

# Custom Mermaid Color Themes
function Set-CustomMermaidTheme {
    param (
        [Parameter(Mandatory=$false)][String]$NodeColor,
        [Parameter(Mandatory=$false)][String]$NodeBorderColor,
        [Parameter(Mandatory=$false)][String]$FontColor,
        [Parameter(Mandatory=$false)][String]$LinkColor,
        [Parameter(Mandatory=$false)][String]$LinkTextColor
    )


    $themedNodes = "classDef customTheme fill:$NodeColor,stroke:$NodeBorderColor,stroke-width:2px,color:$FontColor`n`nclass "

    $allMermaidNodes = $allMermaidNodes | Sort-Object -Unique

    foreach ($node in $allMermaidNodes) {

        $themedNodes += "$node,"

    }

    $mermaidString = ($mermaidCode | Out-String)
    $NumberOfMermaidLinks = (Select-String -InputObject $mermaidString -Pattern '(--)|(-.-)' -AllMatches).Matches.Count

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

    Set-Content -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow$fileExtension" -Value $mermaidCode -Encoding UTF8

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

    mmdc -i "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow$fileExtension" -o "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.png" -b transparent -w "3840" -H "2160" -t "$pngTheme"

    if ($DocType -eq "Markdown") {

        $createdPng = Get-ChildItem -Path "$FilePath\$(($VoiceAppFileName).Replace(" ","_"))_CallFlow-1.png"
        Rename-Item -Path $createdPng.FullName -NewName "$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.png" -Force

    }

}

if ($SetClipBoard -eq $true) {
    $mermaidCode -Replace('```mermaid','') `
    -Replace('```','') `
    -Replace("# Call Flow $VoiceAppFileName","") `
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
        -Replace "# Call Flow $VoiceAppFileName","" `
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

