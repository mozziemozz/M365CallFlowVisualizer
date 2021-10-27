#Requires -Modules MsOnline, MicrosoftTeams

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)][ValidateSet("Markdown","Mermaid")][String]$docType = "Markdown",
    [Parameter(Mandatory=$false)][Bool]$ShowNestedDepth = $false,
    [Parameter(Mandatory=$false)][Switch]$SubSequentRun,
    [Parameter(Mandatory=$false)][string]$PhoneNumber

)

Write-Host "Show nested depth $ShowNestedDepth" -ForegroundColor Cyan

# From: https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/clearing-all-user-variables
function Get-UserVariable ($Name = '*') {
# these variables may exist in certain environments (like ISE, or after use of foreach)
$special = 'ps','psise','psunsupportedconsoleapplications', 'foreach', 'profile'

$ps = [PowerShell]::Create()
$null = $ps.AddScript('$null=$host;Get-Variable') 
$reserved = $ps.Invoke() | Select-Object -ExpandProperty Name
$ps.Runspace.Close()
$ps.Dispose()
Get-Variable -Scope Global | 
    Where-Object Name -like $Name |
    Where-Object { $reserved -notcontains $_.Name } |
    Where-Object { $special -notcontains $_.Name } |
    Where-Object Name 
}

if ($SubSequentRun) {
    Get-UserVariable | Remove-Variable
}

function Set-Mermaid {
    param (
        [Parameter(Mandatory=$true)][String]$docType
        )

    if ($docType -eq "Markdown") {
        $mdStart =@"
``````mermaid
flowchart TB
"@

        $mdEnd =@"

``````
"@

        $fileExtension = ".md"
    }

    else {
        $mdStart =@"
flowchart TB
"@

        $mdEnd =@"

"@

        $fileExtension = ".mmd"
    }

    $mermaidCode = @()

    $mermaidCode += $mdStart
    $mermaidCode += $mdIncomingCall
    $mermaidCode += $mdVoiceApp
    $mermaidCode += $mdNodeAdditionalNumbers
    $mermaidCode += $mdHolidayAndAfterHoursCheck
    $mermaidCode += $mdCallQueueCallFlow
    $mermaidCode += $mdEnd
    
}

function Get-VoiceApp {
    param (
        [Parameter(Mandatory=$false)][String]$PhoneNumber
        )

        if ($PhoneNumber) {
            $resourceAccount = Get-CsOnlineApplicationInstance | Where-Object {$_.PhoneNumber -match $PhoneNumber}
        }

        else {
            # Get resource account (it was a design choice to select a resource account instead of a voice app, people tend to know the phone number and want to know what happens when a particular number is called.)
            $resourceAccount = Get-CsOnlineApplicationInstance | Where-Object {$_.PhoneNumber -notlike ""} | Select-Object DisplayName, PhoneNumber, ObjectId, ApplicationId | Out-GridView -PassThru -Title "Choose an auto attendant or a call queue from the list:"

        }

        switch ($resourceAccount.ApplicationId) {
            # Application Id for auto attendants
            "ce933385-9390-45d1-9512-c8d228074e07" {
                $voiceAppType = "Auto Attendant"
                $voiceApp = Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -contains $resourceAccount.ObjectId}
            }
            # Application Id for call queues
            "11cd3e2e-fccb-42ad-ad00-878b93575e07" {
                $voiceAppType = "Call Queue"
                $voiceApp = Get-CsCallQueue | Where-Object {$_.ApplicationInstances -contains $resourceAccount.ObjectId}
            }
        }

        # Create ps object to store properties from voice app and resource account
        $voiceAppProperties = New-Object -TypeName psobject
        $voiceAppProperties | Add-Member -MemberType NoteProperty -Name "PhoneNumber" -Value $($resourceAccount.PhoneNumber).Replace("tel:","")
        $voiceAppProperties | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $VoiceApp.Name

        if (!$voiceAppCounter) {
            $voiceAppCounter = 0
        }

        $voiceAppCounter ++
        

        if (!$resourceAccountCounter) {
            $resourceAccountCounter = 0
        }

        $resourceAccountCounter ++


        $mdIncomingCall = "start$($resourceAccountCounter)((Incoming Call at <br> $($voiceAppProperties.PhoneNumber))) --> "
        $mdVoiceApp = "voiceApp$($voiceAppCounter)([$($voiceAppType) <br> $($voiceAppProperties.DisplayName)])"

        $mdNodeAdditionalNumbers = @()

        foreach ($ApplicationInstance in ($VoiceApp.ApplicationInstances | Where-Object {$_ -notcontains $resourceAccount.ObjectId})) {

            $resourceAccountCounter ++

            $additionalResourceAccount = ((Get-CsOnlineApplicationInstance -Identity $ApplicationInstance).PhoneNumber) -replace ("tel:","")

            $mdNodeAdditionalNumber = "start$($resourceAccountCounter)((Incoming Call at <br> $additionalResourceAccount)) -.-> voiceApp$($voiceAppCounter)"

            $mdNodeAdditionalNumbers += $mdNodeAdditionalNumber

        }

        

}

function Find-Holidays {
    param (
        [Parameter(Mandatory=$true)][String]$VoiceAppId

    )

    $aa = Get-CsAutoAttendant -Identity $VoiceAppId

    if ($aa.CallHandlingAssociations.Type.Value -contains "Holiday") {
        $aaHasHolidays = $true    
    }

    else {
        $aaHasHolidays = $false
    }
    
}

function Find-AfterHours {
    param (
        [Parameter(Mandatory=$true)][String]$VoiceAppId

    )

    $aa = Get-CsAutoAttendant -Identity $VoiceAppId

    # Create ps object which has no business hours, needed to check if it matches an auto attendants after hours schedule
    $aaDefaultScheduleProperties = New-Object -TypeName psobject

    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "ComplementEnabled" -Value $true
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "MondayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "TuesdayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "WednesdayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "ThursdayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "FridayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "SaturdayHours" -Value "00:00:00-1.00:00:00"
    $aaDefaultScheduleProperties | Add-Member -MemberType NoteProperty -Name "SundayHours" -Value "00:00:00-1.00:00:00"

    # Convert to string for comparison
    $aaDefaultScheduleProperties = $aaDefaultScheduleProperties | Out-String
    
    # Get the current auto attendants after hours schedule and convert to string
    $aaAfterHoursScheduleProperties = ($aa.Schedules | Where-Object {$_.name -match "after"}).WeeklyRecurrentSchedule | Out-String

    # Check if the auto attendant has business hours by comparing the ps object to the actual config of the current auto attendant
    if ($aaDefaultScheduleProperties -eq $aaAfterHoursScheduleProperties) {
        $aaHasAfterHours = $false
    }

    else {
        $aaHasAfterHours = $true
    }
    
}

function Get-AutoAttendantHolidaysAndAfterHours {
    param (
    )

    if (!$aaCounter) {
        $aaCounter = 0
    }

    $aaCounter ++

    if ($aaHasHolidays -eq $true) {

        # The counter is here so that each element is unique in Mermaid
        $HolidayCounter = 1

        # Create empty mermaid subgraph for holidays
        $mdSubGraphHolidays =@"
subgraph Holidays
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

                $holidayGreeting = "Greeting <br> $($holidayCallFlow.Greetings.ActiveType.Value)"

            }

            $holidayAction = $holidayCallFlow.Menu.MenuOptions.Action.Value

            # Check if holiday call handling is disconnect call
            if ($holidayAction -eq "DisconnectCall") {

                $nodeElementHolidayAction = "elementAAHolidayAction$($HolidayCounter)(($holidayAction))"

            }

            else {

                $holidayActionTargetType = $holidayCallFlow.Menu.MenuOptions.CallTarget.Type.Value

                # Switch through different transfer call to target types
                switch ($holidayActionTargetType) {
                    User { $holidayActionTargetTypeFriendly = "User" 
                    $holidayActionTargetName = (Get-MsolUser -ObjectId $($holidayCallFlow.Menu.MenuOptions.CallTarget.Id)).DisplayName
                }
                    SharedVoicemail { $holidayActionTargetTypeFriendly = "Voicemail"
                    $holidayActionTargetName = (Get-MsolGroup -ObjectId $($holidayCallFlow.Menu.MenuOptions.CallTarget.Id)).DisplayName
                }
                    ExternalPstn { $holidayActionTargetTypeFriendly = "External Number" 
                    $holidayActionTargetName =  ($holidayCallFlow.Menu.MenuOptions.CallTarget.Id).Replace("tel:","")
                }
                    # Check if the application endpoint is an auto attendant or a call queue
                    ApplicationEndpoint {                    
                    $MatchingAA = Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $holidayCallFlow.Menu.MenuOptions.CallTarget.Id}

                        if ($MatchingAA) {

                            $holidayActionTargetTypeFriendly = "[Auto Attendant"
                            $holidayActionTargetName = "$($MatchingAA.Name)]"

                        }

                        else {

                            $MatchingCQ = Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $holidayCallFlow.Menu.MenuOptions.CallTarget.Id}

                            $holidayActionTargetTypeFriendly = "[Call Queue"
                            $holidayActionTargetName = "$($MatchingCQ.Name)]"

                        }

                    }
                
                }

                # Create mermaid code for the holiday action node based on the variables created in the switch statemenet
                $nodeElementHolidayAction = "elementAAHolidayAction$($HolidayCounter)($holidayAction) --> elementAAHolidayActionTargetType$($HolidayCounter)($holidayActionTargetTypeFriendly <br> $holidayActionTargetName)"

            }

            # Create subgraph per holiday call handling inside the Holidays subgraph
            $nodeElementHolidayDetails =@"

subgraph $($holidayCallFlow.Name)
direction LR
elementAAHoliday$($HolidayCounter)(Schedule <br> $($holidaySchedule.FixedSchedule.DateTimeRanges.Start) <br> $($holidaySchedule.FixedSchedule.DateTimeRanges.End)) --> elementAAHolidayGreeting$($HolidayCounter)>$holidayGreeting] --> $nodeElementHolidayAction
    end
"@

            # Increase the counter by 1
            $HolidayCounter ++

            # Add holiday call handling subgraph to holiday subgraph
            $mdSubGraphHolidays += $nodeElementHolidayDetails

        } # End of for-each loop

        # Create end for the holiday subgraph
        $mdSubGraphHolidaysEnd =@"

    end
"@
            
        # Add the end to the holiday subgraph mermaid code
        $mdSubGraphHolidays += $mdSubGraphHolidaysEnd

        # Mermaid node holiday check
        $nodeElementHolidayCheck = "elementHolidayCheck$($aaCounter){During Holiday?}"
    } # End if aa has holidays

    # Check if auto attendant has after hours and holidays
    if ($aaHasAfterHours) {

        # Get the business hours schedule and convert to csv for comparison with hard coded strings
        $aaBusinessHours = ($aa.Schedules | Where-Object {$_.name -match "after"}).WeeklyRecurrentSchedule | ConvertTo-Csv

        # Convert from csv to read the business hours per day
        $aaBusinessHoursFriendly = $aaBusinessHours | ConvertFrom-Csv

        # Monday
        # Check if Monday has business hours which are open 24 hours per day
        if ($aaBusinessHoursFriendly.DisplayMondayHours -eq "00:00:00-1.00:00:00") {
            $mondayHours = "Monday Hours: Open 24 hours"
        }
        # Check if Monday has business hours set different than 24 hours open per day
        elseif ($aaBusinessHoursFriendly.DisplayMondayHours) {
            $mondayHours = "Monday Hours: $($aaBusinessHoursFriendly.DisplayMondayHours)"
        }
        # Check if Monday has no business hours at all / is closed 24 hours per day
        else {
            $mondayHours = "Monday Hours: Closed"
        }

        # Tuesday
        if ($aaBusinessHoursFriendly.DisplayTuesdayHours -eq "00:00:00-1.00:00:00") {
            $TuesdayHours = "Tuesday Hours: Open 24 hours"
        }
        elseif ($aaBusinessHoursFriendly.DisplayTuesdayHours) {
            $TuesdayHours = "Tuesday Hours: $($aaBusinessHoursFriendly.DisplayTuesdayHours)"
        } 
        else {
            $TuesdayHours = "Tuesday Hours: Closed"
        }

        # Wednesday
        if ($aaBusinessHoursFriendly.DisplayWednesdayHours -eq "00:00:00-1.00:00:00") {
            $WednesdayHours = "Wednesday Hours: Open 24 hours"
        } 
        elseif ($aaBusinessHoursFriendly.DisplayWednesdayHours) {
            $WednesdayHours = "Wednesday Hours: $($aaBusinessHoursFriendly.DisplayWednesdayHours)"
        }
        else {
            $WednesdayHours = "Wednesday Hours: Closed"
        }

        # Thursday
        if ($aaBusinessHoursFriendly.DisplayThursdayHours -eq "00:00:00-1.00:00:00") {
            $ThursdayHours = "Thursday Hours: Open 24 hours"
        } 
        elseif ($aaBusinessHoursFriendly.DisplayThursdayHours) {
            $ThursdayHours = "Thursday Hours: $($aaBusinessHoursFriendly.DisplayThursdayHours)"
        }
        else {
            $ThursdayHours = "Thursday Hours: Closed"
        }

        # Friday
        if ($aaBusinessHoursFriendly.DisplayFridayHours -eq "00:00:00-1.00:00:00") {
            $FridayHours = "Friday Hours: Open 24 hours"
        } 
        elseif ($aaBusinessHoursFriendly.DisplayFridayHours) {
            $FridayHours = "Friday Hours: $($aaBusinessHoursFriendly.DisplayFridayHours)"
        }
        else {
            $FridayHours = "Friday Hours: Closed"
        }

        # Saturday
        if ($aaBusinessHoursFriendly.DisplaySaturdayHours -eq "00:00:00-1.00:00:00") {
            $SaturdayHours = "Saturday Hours: Open 24 hours"
        } 

        elseif ($aaBusinessHoursFriendly.DisplaySaturdayHours) {
            $SaturdayHours = "Saturday Hours: $($aaBusinessHoursFriendly.DisplaySaturdayHours)"
        }

        else {
            $SaturdayHours = "Saturday Hours: Closed"
        }

        # Sunday
        if ($aaBusinessHoursFriendly.DisplaySundayHours -eq "00:00:00-1.00:00:00") {
            $SundayHours = "Sunday Hours: Open 24 hours"
        }
        elseif ($aaBusinessHoursFriendly.DisplaySundayHours) {
            $SundayHours = "Sunday Hours: $($aaBusinessHoursFriendly.DisplaySundayHours)"
        }

        else {
            $SundayHours = "Sunday Hours: Closed"
        }

        # Create the mermaid node for business hours check including the actual business hours
        $nodeElementAfterHoursCheck = "elementAfterHoursCheck$($aaCounter){During Business Hours? <br> $mondayHours <br> $tuesdayHours  <br> $wednesdayHours  <br> $thursdayHours <br> $fridayHours <br> $saturdayHours <br> $sundayHours}"

    } # End if aa has after hours

    if ($aaHasHolidays -eq $true) {

        if ($aaHasAfterHours) {

            $mdHolidayAndAfterHoursCheck =@"
--> $nodeElementHolidayCheck
$nodeElementHolidayCheck -->|Yes| Holidays
$nodeElementHolidayCheck -->|No| $nodeElementAfterHoursCheck
$nodeElementAfterHoursCheck -->|Yes| $mdAutoAttendantDefaultCallFlow
$nodeElementAfterHoursCheck -->|No| $mdAutoAttendantAfterHoursCallFlow

$mdSubGraphHolidays

"@
        }

        else {
            $mdHolidayAndAfterHoursCheck =@"
--> $nodeElementHolidayCheck
$nodeElementHolidayCheck -->|Yes| Holidays
$nodeElementHolidayCheck -->|No| $mdAutoAttendantDefaultCallFlow

$mdSubGraphHolidays

"@
        }

    }

    
    # Check if auto attendant has no Holidays but after hours
    else {
    
        if ($aaHasAfterHours -eq $true) {

            $mdHolidayAndAfterHoursCheck =@"
--> $nodeElementAfterHoursCheckCheck
$nodeElementAfterHoursCheck -->|Yes| $mdAutoAttendantDefaultCallFlow
$nodeElementAfterHoursCheck -->|No| $mdAutoAttendantAfterHoursCallFlow

"@      
        }

        # Check if auto attendant has no after hours and no holidays
        else {

            $mdHolidayAndAfterHoursCheck =@"
--> $mdAutoAttendantDefaultCallFlow

"@
        }

    
    }

}

function Get-CallQueueCallFlow {
    param (
        [Parameter(Mandatory=$true)][String]$MatchingCQIdentity
    )

    if (!$cqCallFlowCounter) {
        $cqCallFlowCounter = 0
    }

    $cqCallFlowCounter ++

    $MatchingCQ = Get-CsCallQueue -Identity $MatchingCQIdentity

    # Store all neccessary call queue properties in variables
    $CqOverFlowThreshold = $MatchingCQ.OverflowThreshold
    $CqOverFlowAction = $MatchingCQ.OverflowAction.Value
    $CqTimeOut = $MatchingCQ.TimeoutThreshold
    $CqTimeoutAction = $MatchingCQ.TimeoutAction.Value
    $CqRoutingMethod = $MatchingCQ.RoutingMethod.Value
    $CqAgents = $MatchingCQ.Agents.ObjectId
    $CqAgentOptOut = $MatchingCQ.AllowOptOut
    $CqConferenceMode = $MatchingCQ.ConferenceMode
    $CqAgentAlertTime = $MatchingCQ.AgentAlertTime
    $CqPresenceBasedRouting = $MatchingCQ.PresenceBasedRouting
    $CqDistributionList = $MatchingCQ.DistributionLists
    $CqDefaultMusicOnHold = $MatchingCQ.UseDefaultMusicOnHold
    $CqWelcomeMusicFileName = $MatchingCQ.WelcomeMusicFileName

    # Check if call queue uses default music on hold
    if ($CqDefaultMusicOnHold -eq $true) {
        $CqMusicOnHold = "Default"
    }

    else {
        $CqMusicOnHold = "Custom"
    }

    # Check if call queue uses a greeting
    if (!$CqWelcomeMusicFileName) {
        $CqGreeting = "None"
    }

    else {
        $CqGreeting = "Audio File"

    }

    # Check if call queue useses users, group or teams channel as distribution list
    if (!$CqDistributionList) {

        $CqAgentListType = "Users"

    }

    else {

        if (!$MatchingCQ.ChannelId) {

            $CqAgentListType = "Group"

        }

        else {

            $CqAgentListType = "Teams Channel"

        }

    }

    # Switch through call queue overflow action target
    switch ($CqOverFlowAction) {
        DisconnectWithBusy {
            $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowCounter)((Disconnect Call))"
        }
        Forward {

            if ($MatchingCQ.OverflowActionTarget.Type -eq "User") {

                $MatchingOverFlowUser = (Get-MsolUser -ObjectId $MatchingCQ.OverflowActionTarget.Id).DisplayName

                $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqOverFlowActionTarget$($cqCallFlowCounter)(User <br> $MatchingOverFlowUser)"

            }

            elseif ($MatchingCQ.OverflowActionTarget.Type -eq "Phone") {

                $cqOverFlowPhoneNumber = ($MatchingCQ.OverflowActionTarget.Id).Replace("tel:","")

                $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqOverFlowActionTarget$($cqCallFlowCounter)(External Number <br> $cqOverFlowPhoneNumber)"
                
            }

            else {

                $MatchingOverFlowAA = (Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.OverflowActionTarget.Id}).Name

                if ($MatchingOverFlowAA) {

                    $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqOverFlowActionTarget$($cqCallFlowCounter)([Auto Attendant <br> $MatchingOverFlowAA])"

                }

                else {

                    $MatchingOverFlowCQ = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.OverflowActionTarget.Id})

                    $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqOverFlowActionTarget$($cqCallFlowCounter)([Call Queue <br> $($MatchingOverFlowCQ.Name)])"

                }

            }

        }
        SharedVoicemail {
            $MatchingOverFlowVoicemail = (Get-MsolGroup -ObjectId $MatchingCQ.OverflowActionTarget.Id).DisplayName

            if ($MatchingCQ.OverflowSharedVoicemailTextToSpeechPrompt) {

                $CqOverFlowVoicemailGreeting = "TextToSpeech"

            }

            else {

                $CqOverFlowVoicemailGreeting = "AudioFile"

            }

            $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqOverFlowVoicemailGreeting$($cqCallFlowCounter)>Greeting <br> $CqOverFlowVoicemailGreeting] --> cqOverFlowActionTarget$($cqCallFlowCounter)(Shared Voicemail <br> $MatchingOverFlowVoicemail)"

        }

    }

    # Switch through call queue timeout overflow action
    switch ($CqTimeoutAction) {
        Disconnect {
            $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowCounter)((Disconnect Call))"
        }
        Forward {
    
            if ($MatchingCQ.TimeoutActionTarget.Type -eq "User") {

                $MatchingTimeoutUser = (Get-MsolUser -ObjectId $MatchingCQ.TimeoutActionTarget.Id).DisplayName
    
                $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqTimeoutActionTarget$($cqCallFlowCounter)(User <br> $MatchingTimeoutUser)"
    
            }
    
            elseif ($MatchingCQ.TimeoutActionTarget.Type -eq "Phone") {
    
                $cqTimeoutPhoneNumber = ($MatchingCQ.TimeoutActionTarget.Id).Replace("tel:","")
    
                $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqTimeoutActionTarget$($cqCallFlowCounter)(External Number <br> $cqTimeoutPhoneNumber)"
                
            }
    
            else {
    
                $MatchingTimeoutAA = (Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.TimeoutActionTarget.Id}).Name
    
                if ($MatchingTimeoutAA) {
    
                    $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqTimeoutActionTarget$($cqCallFlowCounter)([Auto Attendant <br> $MatchingTimeoutAA])"
    
                }
    
                else {
    
                    $MatchingTimeoutCQ = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.TimeoutActionTarget.Id})

                    $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqTimeoutActionTarget$($cqCallFlowCounter)([Call Queue <br> $($MatchingTimeoutCQ.Name)])"
    
                }
    
            }
    
        }
        SharedVoicemail {
            $MatchingTimeoutVoicemail = (Get-MsolGroup -ObjectId $MatchingCQ.TimeoutActionTarget.Id).DisplayName
    
            if ($MatchingCQ.TimeoutSharedVoicemailTextToSpeechPrompt) {
    
                $CqTimeoutVoicemailGreeting = "TextToSpeech"
    
            }
    
            else {
    
                $CqTimeoutVoicemailGreeting = "AudioFile"
    
            }
    
            $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqTimeoutVoicemailGreeting$($cqCallFlowCounter)>Greeting <br> $CqTimeoutVoicemailGreeting] --> cqTimeoutActionTarget$($cqCallFlowCounter)(Shared Voicemail <br> $MatchingTimeoutVoicemail)"
    
        }
    
    }

    # Create empty mermaid element for agent list
    $mdCqAgentsDisplayNames = @"
"@

    # Define agent counter for unique mermaid element names
    $AgentCounter = 1

    # add each agent to the empty agents mermaid element
    foreach ($CqAgent in $CqAgents) {
        $AgentDisplayName = (Get-MsolUser -ObjectId $CqAgent).DisplayName

        $AgentDisplayNames = "agentListType$($cqCallFlowCounter) --> agent$($cqCallFlowCounter)$($AgentCounter)($AgentDisplayName) --> timeOut$($cqCallFlowCounter)`n"

        $mdCqAgentsDisplayNames += $AgentDisplayNames

        $AgentCounter ++
    }

    switch ($voiceAppType) {
        "Auto Attendant" {

            $voiceAppTypeSpecificCallFlow = "--> defaultCallFlow$($cqCallFlowCounter)($defaultCallFlowAction) --> defaultCallFlowAction$($cqCallFlowCounter)($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName) --> cqGreeting$($cqCallFlowCounter)>Greeting <br> $CqGreeting]"

        }
        "Call Queue" {

            $voiceAppTypeSpecificCallFlow = "--> cqGreeting$($cqCallFlowCounter)>Greeting <br> $CqGreeting]"

            $defaultCallFlowcCqIsTopLevel = $null

        }
    }

    # Create default callflow mermaid code

$mdCallQueueCallFlow =@"
$voiceAppTypeSpecificCallFlow
--> overFlow$($cqCallFlowCounter){More than $CqOverFlowThreshold <br> Active Calls}
overFlow$($cqCallFlowCounter) ---> |Yes| $CqOverFlowActionFriendly
overFlow$($cqCallFlowCounter) ---> |No| routingMethod$($cqCallFlowCounter)

$defaultCallFlowcCqIsTopLevel

subgraph Call Distribution
subgraph CQ Settings
routingMethod$($cqCallFlowCounter)[(Routing Method: $CqRoutingMethod)] --> agentAlertTime$($cqCallFlowCounter)
agentAlertTime$($cqCallFlowCounter)[(Agent Alert Time: $CqAgentAlertTime)] -.- cqMusicOnHold$($cqCallFlowCounter)
cqMusicOnHold$($cqCallFlowCounter)[(Music On Hold: $CqMusicOnHold)] -.- conferenceMode$($cqCallFlowCounter)
conferenceMode$($cqCallFlowCounter)[(Conference Mode Enabled: $CqConferenceMode)] -.- agentOptOut$($cqCallFlowCounter)
agentOptOut$($cqCallFlowCounter)[(Agent Opt Out Allowed: $CqAgentOptOut)] -.- presenceBasedRouting$($cqCallFlowCounter)
presenceBasedRouting$($cqCallFlowCounter)[(Presence Based Routing: $CqPresenceBasedRouting)] -.- timeOut$($cqCallFlowCounter)
timeOut$($cqCallFlowCounter)[(Timeout: $CqTimeOut Seconds)]
end
subgraph Agents $($MatchingCQ.Name)
agentAlertTime$($cqCallFlowCounter) --> agentListType$($cqCallFlowCounter)[(Agent List Type: $CqAgentListType)]
$mdCqAgentsDisplayNames
end
end

timeOut$($cqCallFlowCounter) --> cqResult$($cqCallFlowCounter){Call Connected?}
cqResult$($cqCallFlowCounter) --> |Yes| cqEnd$($cqCallFlowCounter)((Call Connected))
cqResult$($cqCallFlowCounter) --> |No| $CqTimeoutActionFriendly

"@
    
}

function Get-NestedCallQueueCallFlow {
    param (
        [Parameter(Mandatory=$true)][String]$MatchingCQIdentity
    )

    $nestedCallQueueCallFlow = $mdCallQueueCallFlow

    . Get-CallQueueCallFlow -MatchingCQIdentity $MatchingCQIdentity

    $nestedCallQueueCallFlow += $mdCallQueueCallFlow

    $mdCallQueueCallFlow = $nestedCallQueueCallFlow

}

function Get-AutoAttendantDefaultCallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceAppId
    )

    if (!$aaDefaultCallFlowCounter) {
        $aaDefaultCallFlowCounter = 0
    }

    $aaDefaultCallFlowCounter ++

    # Get the current auto attendants default call flow and default call flow action
    $defaultCallFlow = $aa.DefaultCallFlow
    $defaultCallFlowAction = $aa.DefaultCallFlow.Menu.MenuOptions.Action.Value

    # Get the current auto attentans default call flow greeting
    $defaultCallFlowGreeting = "Greeting <br> $($defaultCallFlow.Greetings.ActiveType.Value)"

    # Check if the default callflow action is transfer call to target
    if ($defaultCallFlowAction -eq "TransferCallToTarget") {

        # Get transfer target type
        $defaultCallFlowTargetType = $aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Type.Value

        # Switch through transfer target type and set variables accordingly
        switch ($defaultCallFlowTargetType) {
            User { 
                $defaultCallFlowTargetTypeFriendly = "User"
                $defaultCallFlowTargetName = (Get-MsolUser -ObjectId $($aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id)).DisplayName}
            ExternalPstn { 
                $defaultCallFlowTargetTypeFriendly = "External PSTN"
                $defaultCallFlowTargetName = ($aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id).Replace("tel:","")}
            ApplicationEndpoint {

                # Check if application endpoint is auto attendant or call queue
                $MatchingAA = Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id}

                if ($MatchingAA) {

                    $defaultCallFlowTargetTypeFriendly = "[Auto Attendant"
                    $defaultCallFlowTargetName = "$($MatchingAA.Name)]"

                }

                else {

                    $MatchingCQ = Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id}

                    $cqAssociatedApplicationInstance = Get-CsOnlineApplicationInstance -Identity $MatchingCQ.DisplayApplicationInstances

<#                     # check if call queue also has its own phone number
                    if ($cqAssociatedApplicationInstance.PhoneNumber) {
                        $secondTopLevelNumber = $($cqAssociatedApplicationInstance.PhoneNumber).Replace("tel:","")
                        $defaultCallFlowcCqIsTopLevel = "start200((Incoming Call at <br> $($secondTopLevelNumber))) -...-> defaultCallFlowAction"

                    }

                    else {
                        
                        $defaultCallFlowcCqIsTopLevel = $null
                    } #>

                    $defaultCallFlowTargetTypeFriendly = "[Call Queue"
                    $defaultCallFlowTargetName = "$($MatchingCQ.Name)]"

                }

            }
            SharedVoicemail {

                $defaultCallFlowTargetTypeFriendly = "Voicemail"
                $defaultCallFlowTargetName = (Get-MsolGroup -ObjectId $aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id).DisplayName

            }
        }

        # Check if transfer target type is call queue
        if ($defaultCallFlowTargetTypeFriendly -eq "[Call Queue") {

            $MatchingCQIdentity = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id}).Identity

            $aaDefaultCallFlowForwardsToCq = $true

            $mdAutoAttendantDefaultCallFlow = "defaultCallFlowGreeting$($aaDefaultCallFlowCounter)>$defaultCallFlowGreeting] --> defaultCallFlow$($aaDefaultCallFlowCounter)($defaultCallFlowAction) --> defaultCallFlowAction$($aaDefaultCallFlowCounter)($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName)"

            
        } # End if transfer target type is call queue

        # Check if default callflow action target is trasnfer call to target but something other than call queue
        else {

            $mdAutoAttendantDefaultCallFlow = "defaultCallFlowGreeting$($aaDefaultCallFlowCounter)>$defaultCallFlowGreeting] --> defaultCallFlow$($aaDefaultCallFlowCounter)($defaultCallFlowAction) --> defaultCallFlowAction$($aaDefaultCallFlowCounter)($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName)"

        }

    }

    # Check if default callflow action is disconnect call
    elseif ($defaultCallFlowAction -eq "DisconnectCall") {

        $mdAutoAttendantDefaultCallFlow = "defaultCallFlowGreeting$($aaDefaultCallFlowCounter)>$defaultCallFlowGreeting] --> defaultCallFlow$($aaDefaultCallFlowCounter)(($defaultCallFlowAction))"

    }
    
    
}

function Get-AutoAttendantAfterHoursCallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceAppId
    )

    if (!$aaAfterHoursCallFlowCounter) {
        $aaAfterHoursCallFlowCounter = 0
    }

    $aaAfterHoursCallFlowCounter ++

    # Get after hours call flow
    $afterHoursCallFlow = ($aa.CallFlows | Where-Object {$_.Name -Match "after hours"})
    $afterHoursCallFlowAction = ($aa.CallFlows | Where-Object {$_.Name -Match "after hours"}).Menu.MenuOptions.Action.Value

    # Get after hours greeting
    $afterHoursCallFlowGreeting = "Greeting <br> $($afterHoursCallFlow.Greetings.ActiveType.Value)"

    # Check if after hours action is transfer call to target
    if ($afterHoursCallFlowAction -eq "TransferCallToTarget") {

        $afterHoursCallFlowTargetType = $afterHoursCallFlow.Menu.MenuOptions.CallTarget.Type.Value

        # Switch through after hours call flow target type
        switch ($afterHoursCallFlowTargetType) {
            User { 
                $afterHoursCallFlowTargetTypeFriendly = "User"
                $afterHoursCallFlowTargetName = (Get-MsolUser -ObjectId $($afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id)).DisplayName}
            ExternalPstn { 
                $afterHoursCallFlowTargetTypeFriendly = "External PSTN"
                $afterHoursCallFlowTargetName = ($afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id).Replace("tel:","")}
            ApplicationEndpoint {

                # Check if application endpoint is an auto attendant or a call queue
                $MatchingAA = Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id}

                if ($MatchingAA) {

                    $afterHoursCallFlowTargetTypeFriendly = "[Auto Attendant"
                    $afterHoursCallFlowTargetName = "$($MatchingAA.Name)]"

                }

                else {

                    $MatchingCQ = Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id}

                    $afterHoursCallFlowTargetTypeFriendly = "[Call Queue"
                    $afterHoursCallFlowTargetName = "$($MatchingCQ.Name)]"

                }

            }
            SharedVoicemail {

                $afterHoursCallFlowTargetTypeFriendly = "Voicemail"
                $afterHoursCallFlowTargetName = (Get-MsolGroup -ObjectId $afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id).DisplayName

            }
        }

        # Check if transfer target type is call queue
        if ($afterHoursCallFlowTargetTypeFriendly -eq "[Call Queue") {

            $MatchingCQIdentity = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $aa.AfterHoursCallFlow.Menu.MenuOptions.CallTarget.Id}).Identity

            $aaAfterHoursCallFlowForwardsToCq = $true

            $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowCounter)>$AfterHoursCallFlowGreeting] --> AfterHoursCallFlow$($aaAfterHoursCallFlowCounter)($AfterHoursCallFlowAction) --> AfterHoursCallFlowAction$($aaAfterHoursCallFlowCounter)($AfterHoursCallFlowTargetTypeFriendly <br> $AfterHoursCallFlowTargetName)"

            
        } # End if transfer target type is call queue

        # Check if AfterHours callflow action target is trasnfer call to target but something other than call queue
        else {

            $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowCounter)>$AfterHoursCallFlowGreeting] --> AfterHoursCallFlow$($aaAfterHoursCallFlowCounter)($AfterHoursCallFlowAction) --> AfterHoursCallFlowAction$($aaAfterHoursCallFlowCounter)($AfterHoursCallFlowTargetTypeFriendly <br> $AfterHoursCallFlowTargetName)"

        }


        # Mermaid code for after hours call flow nodes
        $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowCounter)>$AfterHoursCallFlowGreeting] --> afterHoursCallFlow$($aaAfterHoursCallFlowCounter)($afterHoursCallFlowAction) --> afterHoursCallFlowAction$($aaAfterHoursCallFlowCounter)($afterHoursCallFlowTargetTypeFriendly <br> $afterHoursCallFlowTargetName)"

    }

    elseif ($afterHoursCallFlowAction -eq "DisconnectCall") {

        $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowCounter)>$AfterHoursCallFlowGreeting] --> afterHoursCallFlow$($aaAfterHoursCallFlowCounter)(($afterHoursCallFlowAction))"

    }
    

    

}

if ($PhoneNumber) {
    . Get-VoiceApp -PhoneNumber $PhoneNumber
}

else {
    . Get-VoiceApp
}


if ($voiceAppType -eq "Auto Attendant") {
    . Find-Holidays -VoiceAppId $VoiceApp.Identity
    . Find-AfterHours -VoiceAppId $VoiceApp.Identity

    if ($aaHasHolidays -eq $true -or $aaHasAfterHours -eq $true) {

        . Get-AutoAttendantDefaultCallFlow -VoiceAppId $VoiceApp.Identity

        . Get-AutoAttendantAfterHoursCallFlow -VoiceAppId $VoiceApp.Identity

        . Get-AutoAttendantHolidaysAndAfterHours -VoiceAppId $VoiceApp.Identity

    }

    else {

        . Get-AutoAttendantDefaultCallFlow -VoiceAppId $VoiceApp.Identity

        $mdHolidayAndAfterHoursCheck =@"
        --> $mdAutoAttendantDefaultCallFlow
        
"@

    }

<#     if ($aaDefaultCallFlowForwardsToCq -eq $true) {

        . Get-CallQueueCallFlow -MatchingCQIdentity $MatchingCQIdentity

    }

    else {

    }
 #>
}

elseif ($voiceAppType -eq "Call Queue") {
    . Get-CallQueueCallFlow -MatchingCQIdentity $VoiceApp.Identity
}

if ($ShowNestedDepth -and $MatchingTimeoutCQ) {

    . Get-NestedCallQueueCallFlow -MatchingCQIdentity $MatchingTimeoutCQ.Identity
}


. Set-Mermaid -docType $docType

$mermaidCode -Replace('```mermaid','') `
-Replace('```','') | Set-Clipboard

Write-Host "Mermaid code copied to clipboard." -ForegroundColor Cyan
