<#
    .SYNOPSIS
    Reads the configuration from a Microsoft 365 Phone System auto attendant and visualizes the call flow using mermaid-js.

    .DESCRIPTION
    Presents a selection of available auto attendants and then reads the config of that auto attendant and writes it into a mermaid-js flowchart file.

    Author:             Martin Heusser
    Version:            1.0.0
    Revision:
        20.10.2021:     Creation

    .PARAMETER Name
    -docType
        Mandatory=$false
        [String], allowed values: Markdown (default), Mermaid
            Markdown: Creates a markdown file (*.md) with mermaid code declaration
            Mermaid: Creates a mermaid file (*.mmd)

    .INPUTS
    - Auto Attendant Name (through Out-GridView)

    .OUTPUTS
    - Mermaid flowchart

    .EXAMPLE

    .LINK
    https://github.com/mozziemozz/M365CallFlowVisualizer
#>

#Requires -Modules MsOnline, MicrosoftTeams

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)][ValidateSet("Markdown","Mermaid")][String]$docType = "Markdown"
)

# Check if doctype is markdown or mermaid and create mermaid code and file extension accordingly
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

# Get resource account
$resourceAccount = Get-CsOnlineApplicationInstance | Where-Object {$_.PhoneNumber -notlike "" -and $_.ApplicationId -eq "ce933385-9390-45d1-9512-c8d228074e07"} | Select-Object DisplayName, PhoneNumber, ObjectId | Out-GridView -PassThru -Title "Choose an Auto Attendant from the list to visualize your call flow:"

# Create ps object to store properties from auto attendant and resource account
$aaProperties = New-Object -TypeName psobject

$aaProperties | Add-Member -MemberType NoteProperty -Name "PhoneNumber" -Value $($resourceAccount.PhoneNumber).Replace("tel:","")

# Get auto attendant
$aa = Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $resourceAccount.ObjectId}

$aaProperties | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $aa.Name

# Check if AA has holidays
if ($aa.CallHandlingAssociations.Type.Value -contains "Holiday") {

    $aaHasHolidays = $true

    $aaHolidays = $aa.CallHandlingAssociations | Where-Object {$_.Type -match "Holiday" -and $_.Enabled -eq $true}

    # Create empty mermaid subgraph for holidays
    $mdSubGraphHolidays =@"
subgraph Holidays
    direction LR
"@

# The counter is here so that each element is unique in Mermaid
    $HolidayCounter = 1
    
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

    }

    # Create end for the holiday subgraph
    $mdSubGraphHolidaysEnd =@"

end
"@

    # Add the end to the holiday subgraph mermaid code
    $mdSubGraphHolidays += $mdSubGraphHolidaysEnd

}

else {

    $aaHasHolidays = $false

}

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
    $nodeElementAfterHoursCheck = "elementAfterHoursCheck{During Business Hours? <br> $mondayHours <br> $tuesdayHours  <br> $wednesdayHours  <br> $thursdayHours <br> $fridayHours <br> $saturdayHours <br> $sundayHours}"

}

# Mermaid nodes start and Auto Attendant
$nodeStart = "start((Incoming Call at <br> $($aaProperties.PhoneNumber)))"
$nodeElementAA = "elementAA([Auto Attendant <br> $($aaProperties.DisplayName)])"

# Mermaid node holiday check
$nodeElementHolidayCheck = "elementHolidayCheck{During Holiday?}"

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

                # check if call queue also has its own phone number
                if ($cqAssociatedApplicationInstance.PhoneNumber) {

                    $defaultCallFlowcCqIsTopLevel = "start2((Incoming Call at <br> $($cqAssociatedApplicationInstance.PhoneNumber))) -...-> defaultCallFlowAction"

                }

                else {
                    
                    $defaultCallFlowcCqIsTopLevel = $null
                }

                $defaultCallFlowTargetTypeFriendly = "[Call Queue"
                $defaultCallFlowTargetName = "$($MatchingCQ.Name)]"

            }

        }
        SharedVoicemail {

            $defaultCallFlowTargetTypeFriendly = "Voicemail"
            $defaultCallFlowTargetName = (Get-MsolGroup -ObjectId $aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id).DisplayName

        }
    }

    # Check if transfer target type is call quueue
    if ($defaultCallFlowTargetTypeFriendly -eq "[Call Queue") {

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
                $CqOverFlowActionFriendly = "cqOverFlowAction((Disconnect Call))"
            }
            Forward {

                if ($MatchingCQ.OverflowActionTarget.Type -eq "User") {

                    $MatchingOverFlowUser = (Get-MsolUser -ObjectId $MatchingCQ.OverflowActionTarget.Id).DisplayName

                    $CqOverFlowActionFriendly = "cqOverFlowAction(TransferCallToTarget) --> cqOverFlowActionTarget(User <br> $MatchingOverFlowUser)"

                }

                elseif ($MatchingCQ.OverflowActionTarget.Type -eq "Phone") {

                    $cqOverFlowPhoneNumber = ($MatchingCQ.OverflowActionTarget.Id).Replace("tel:","")

                    $CqOverFlowActionFriendly = "cqOverFlowAction(TransferCallToTarget) --> cqOverFlowActionTarget(External Number <br> $cqOverFlowPhoneNumber)"
                    
                }

                else {

                    $MatchingOverFlowAA = (Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.OverflowActionTarget.Id}).Name

                    if ($MatchingOverFlowAA) {

                        $CqOverFlowActionFriendly = "cqOverFlowAction(TransferCallToTarget) --> cqOverFlowActionTarget([Auto Attendant <br> $MatchingOverFlowAA])"

                    }

                    else {

                        $MatchingOverFlowCQ = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.OverflowActionTarget.Id}).Name

                        $CqOverFlowActionFriendly = "cqOverFlowAction(TransferCallToTarget) --> cqOverFlowActionTarget([Call Queue <br> $MatchingOverFlowCQ])"

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

                $CqOverFlowActionFriendly = "cqOverFlowAction(TransferCallToTarget) --> cqOverFlowVoicemailGreeting>Greeting <br> $CqOverFlowVoicemailGreeting] --> cqOverFlowActionTarget(Shared Voicemail <br> $MatchingOverFlowVoicemail)"

            }

        }

        # Switch through call queue timeout overflow action
        switch ($CqTimeoutAction) {
            DisconnectWithBusy {
                $CqTimeoutActionFriendly = "cqTimeoutAction((Disconnect Call))"
            }
            Forward {
        
                if ($MatchingCQ.TimeoutActionTarget.Type -eq "User") {

                    $MatchingTimeoutUser = (Get-MsolUser -ObjectId $MatchingCQ.TimeoutActionTarget.Id).DisplayName
        
                    $CqTimeoutActionFriendly = "cqTimeoutAction(TransferCallToTarget) --> cqTimeoutActionTarget(User <br> $MatchingTimeoutUser)"
        
                }
        
                elseif ($MatchingCQ.TimeoutActionTarget.Type -eq "Phone") {
        
                    $cqTimeoutPhoneNumber = ($MatchingCQ.TimeoutActionTarget.Id).Replace("tel:","")
        
                    $CqTimeoutActionFriendly = "cqTimeoutAction(TransferCallToTarget) --> cqTimeoutActionTarget(External Number <br> $cqTimeoutPhoneNumber)"
                    
                }
        
                else {
        
                    $MatchingTimeoutAA = (Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.TimeoutActionTarget.Id}).Name
        
                    if ($MatchingTimeoutAA) {
        
                        $CqTimeoutActionFriendly = "cqTimeoutAction(TransferCallToTarget) --> cqTimeoutActionTarget([Auto Attendant <br> $MatchingTimeoutAA])"
        
                    }
        
                    else {
        
                        $MatchingTimeoutCQ = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.TimeoutActionTarget.Id}).Name
        
                        $CqTimeoutActionFriendly = "cqTimeoutAction(TransferCallToTarget) --> cqTimeoutActionTarget([Call Queue <br> $MatchingTimeoutCQ])"
        
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
        
                $CqTimeoutActionFriendly = "cqTimeoutAction(TransferCallToTarget) --> cqTimeoutVoicemailGreeting>Greeting <br> $CqTimeoutVoicemailGreeting] --> cqTimeoutActionTarget(Shared Voicemail <br> $MatchingTimeoutVoicemail)"
        
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

        $AgentDisplayNames = "agentListType --> agent$($AgentCounter)($AgentDisplayName) --> timeOut`n"

        $mdCqAgentsDisplayNames += $AgentDisplayNames

        $AgentCounter ++
    }

    # Create default callflow mermaid code
    $defaultCallFlowMarkDown =@"
--> defaultCallFlow($defaultCallFlowAction) --> defaultCallFlowAction($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName) --> cqGreeting>Greeting <br> $CqGreeting]
--> overFlow{More than $CqOverFlowThreshold <br> Active Calls}
overFlow --> |Yes| $CqOverFlowActionFriendly
overFlow --> |No| routingMethod

$defaultCallFlowcCqIsTopLevel

subgraph Call Distribution
    subgraph CQ Settings
    routingMethod[(Routing Method: $CqRoutingMethod)] --> agentAlertTime
    agentAlertTime[(Agent Alert Time: $CqAgentAlertTime)] -.- cqMusicOnHold
    cqMusicOnHold[(Music On Hold: $CqMusicOnHold)] -.- conferenceMode
    conferenceMode[(Conference Mode Enabled: $CqConferenceMode)] -.- agentOptOut
    agentOptOut[(Agent Opt Out Allowed: $CqAgentOptOut)] -.- presenceBasedRouting
    presenceBasedRouting[(Presence Based Routing: $CqPresenceBasedRouting)] -.- timeOut
    timeOut[(Timeout: $CqTimeOut Seconds)]
    end
    subgraph Agents
    agentAlertTime --> agentListType[(Agent List Type: $CqAgentListType)]
    $mdCqAgentsDisplayNames
    end
end

timeOut --> cqResult{Call Connected?}
    cqResult --> |Yes| cqEnd((Call Connected))
    cqResult --> |No| $CqTimeoutActionFriendly

"@
    
    }

    # Check if default callflow action target is trasnfer call to target but something other than call queue
    else {

        $defaultCallFlowMarkDown = "--> defaultCallFlow($defaultCallFlowAction) --> defaultCallFlowAction($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName)"

    }

}

# Check if default callflow action is disconnect call
elseif ($defaultCallFlowAction -eq "DisconnectCall") {

    $defaultCallFlowMarkDown = "--> defaultCallFlow(($defaultCallFlowAction))"

}

# Mermaid code for default call flow
$mdDefaultCallflow =@"
defaultCallFlowGreeting>$defaultCallFlowGreeting] $defaultCallFlowMarkDown
"@

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

    # Mermaid code for after hours call flow nodes
    $afterHoursCallFlowMarkDown = "--> afterHoursCallFlow($afterHoursCallFlowAction) --> afterHoursCallFlowAction($afterHoursCallFlowTargetTypeFriendly <br> $afterHoursCallFlowTargetName)"

}

elseif ($afterHoursCallFlowAction -eq "DisconnectCall") {

    $afterHoursCallFlowMarkDown = "--> afterHoursCallFlow(($afterHoursCallFlowAction))"

}

# Mermaid code for after hours call flow
$mdAfterHoursCallFlow =@"
afterHoursCallFlowGreeting>$afterHoursCallFlowGreeting] $afterHoursCallFlowMarkDown
"@

# Check if auto attendant has holidays and after hours
if ($aaHasHolidays) {

    # Check if auto attendant has after hours and holidays
    if ($aaHasAfterHours) {
    
    $mdContent =@"
$nodeStart --> $nodeElementAA --> 
$nodeElementHolidayCheck -->|Yes| Holidays
$nodeElementHolidayCheck -->|No| $nodeElementAfterHoursCheck
    $nodeElementAfterHoursCheck -->|Yes| $mdDefaultCallflow
    $nodeElementAfterHoursCheck -->|No| $mdAfterHoursCallFlow

$mdSubGraphHolidays

"@

    }

    # Check if auto attendant has holidays but no after hours
    else {

    $mdContent =@"
$nodeStart --> $nodeElementAA --> $nodeElementHolidayCheck
$nodeElementHolidayCheck -->|Yes| Holidays
$nodeElementHolidayCheck -->|No| $mdDefaultCallflow

$mdSubGraphHolidays

"@

    }

}

# Check if auto attendant has no Holidays but after hours
else {

    # Check if auto attendant has after hours
    if ($aaHasAfterHours) {

    $mdContent =@"
$nodeStart --> $nodeElementAA --> $nodeElementAfterHoursCheckCheck
$nodeElementAfterHoursCheck -->|Yes| $mdDefaultCallflow
$nodeElementAfterHoursCheck -->|No| $mdAfterHoursCallFlow

"@
    }

# Check if auto attendant has no after hours and no holidays
    else {

    $mdContent =@"
$nodeStart --> $nodeElementAA --> $mdDefaultCallflow

"@
    }

}


Set-Content -Path ".\$($aa.Name)_CallFlow$fileExtension" -Value $mdStart, $mdContent, $mdEnd -Encoding UTF8

