#Requires -Modules MsOnline, MicrosoftTeams

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)][ValidateSet("Markdown","Mermaid")][String]$DocType = "Mermaid",
    [Parameter(Mandatory=$false)][string]$PhoneNumber,
    [Parameter(Mandatory=$false)][Bool]$SetClipBoard = $true

)

function Set-Mermaid {
    param (
        [Parameter(Mandatory=$true)][String]$DocType
        )

    if ($DocType -eq "Markdown") {
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


        $mdIncomingCall = "start$($voiceAppCounter)((Incoming Call at <br> $($voiceAppProperties.PhoneNumber))) --> "

        if ($voiceAppCounter -eq 1) {

            $mermaidCode += $mdIncomingCall

        }

        $mdVoiceApp = "voiceApp$($voiceAppCounter)([$($voiceAppType) <br> $($voiceAppProperties.DisplayName)])"

        $nodeElementHolidayLink = $mdVoiceApp

        $mdNodeAdditionalNumbers = @()

        foreach ($ApplicationInstance in ($VoiceApp.ApplicationInstances | Where-Object {$_ -notcontains $resourceAccount.ObjectId})) {

            $voiceAppCounter ++

            $additionalResourceAccount = ((Get-CsOnlineApplicationInstance -Identity $ApplicationInstance).PhoneNumber) -replace ("tel:","")

            $mdNodeAdditionalNumber = "start$($voiceAppCounter)((Incoming Call at <br> $additionalResourceAccount)) -.-> voiceApp$($voiceAppCounter)"

            $mdNodeAdditionalNumbers += $mdNodeAdditionalNumber

        }

        $mermaidCode += $mdVoiceApp
        $mermaidCode += $mdNodeAdditionalNumbers

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
        [Parameter(Mandatory=$true)][String]$VoiceAppId,
        [Parameter(Mandatory=$false)][Bool]$InvokedByNesting = $false,
        [Parameter(Mandatory=$false)][String]$NestedAaCallFlowType
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
subgraph $($aaCounter)-Holidays
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

                $nodeElementHolidayAction = "elementAAHolidayAction$($aaCounter)-$($HolidayCounter)(($holidayAction))"

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
                $nodeElementHolidayAction = "elementAAHolidayAction$($aaCounter)-$($HolidayCounter)($holidayAction) --> elementAAHolidayActionTargetType$($aaCounter)-$($HolidayCounter)($holidayActionTargetTypeFriendly <br> $holidayActionTargetName)"

            }

            # Create subgraph per holiday call handling inside the Holidays subgraph
            $nodeElementHolidayDetails =@"

subgraph $($holidayCallFlow.Name)
direction LR
elementAAHoliday$($aaCounter)-$($HolidayCounter)(Schedule <br> $($holidaySchedule.FixedSchedule.DateTimeRanges.Start) <br> $($holidaySchedule.FixedSchedule.DateTimeRanges.End)) --> elementAAHolidayGreeting$($aaCounter)-$($HolidayCounter)>$holidayGreeting] --> $nodeElementHolidayAction
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

        $aaTimeZone = $aa.TimeZoneId

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
        $nodeElementAfterHoursCheck = "elementAfterHoursCheck$($aaCounter){During Business Hours? <br> Time Zone: $aaTimeZone <br> $mondayHours <br> $tuesdayHours  <br> $wednesdayHours  <br> $thursdayHours <br> $fridayHours <br> $saturdayHours <br> $sundayHours}"

    } # End if aa has after hours

    $aaCounterDifference = $aaCounter -1
    $CallFlowLinkCounter = $aaCounter - $aaCounterDifference

    if ($InvokedByNesting -eq $true -and $NestedAaCallFlowType -eq "AfterHours") {

        $nodeElementHolidayLink = "afterHoursCallFlowAction$($CallFlowLinkCounter)"
        
        Write-Host "Proof this works too!" -ForegroundColor DarkYellow
    }

    if ($InvokedByNesting -eq $true -and $NestedAaCallFlowType -eq "Default") {

        $nodeElementHolidayLink = "defaultCallFlowAction$($CallFlowLinkCounter)"

        Write-Host "Proof this works too!" -ForegroundColor DarkYellow
    }


    if ($aaHasHolidays -eq $true) {

        if ($aaHasAfterHours) {

            $mdHolidayAndAfterHoursCheck =@"
$nodeElementHolidayLink --> $nodeElementHolidayCheck
$nodeElementHolidayCheck -->|Yes| $($aaCounter)-Holidays
$nodeElementHolidayCheck -->|No| $nodeElementAfterHoursCheck
$nodeElementAfterHoursCheck -->|No| $mdAutoAttendantAfterHoursCallFlow
$nodeElementAfterHoursCheck -->|Yes| $mdAutoAttendantDefaultCallFlow

$mdSubGraphHolidays

"@
        }

        else {
            $mdHolidayAndAfterHoursCheck =@"
$nodeElementHolidayLink --> $nodeElementHolidayCheck
$nodeElementHolidayCheck -->|Yes| $($aaCounter)-Holidays
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

    $mermaidCode += $mdHolidayAndAfterHoursCheck

}

function Get-AutoAttendantDefaultCallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceAppId,
        [Parameter(Mandatory=$false)][Bool]$InvokedByNesting = $false,
        [Parameter(Mandatory=$false)][String]$NestedAaCallFlowType
    )

    if (!$aaDefaultCallFlowCounter) {
        $aaDefaultCallFlowCounter = 0
    }

    $aaDefaultCallFlowCounter ++

    # Get the current auto attendants default call flow and default call flow action
    $defaultCallFlow = $aa.DefaultCallFlow
    $defaultCallFlowAction = $aa.DefaultCallFlow.Menu.MenuOptions.Action.Value

    # Get the current auto attentans default call flow greeting
    if (!$defaultCallFlow.Greetings.ActiveType.Value){
        $defaultCallFlowGreeting = "Greeting <br> None"
    }

    else {
        $defaultCallFlowGreeting = "Greeting <br> $($defaultCallFlow.Greetings.ActiveType.Value)"
    }

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

                    if ($InvokedByNesting -eq $false) {
                        $aaDefaultCallFlowForwardsToAa = $true
                        $aaDefaultCallFlowNestedAaIdentity = $MatchingAA.Identity

                    }


                }

                else {

                    $MatchingCqAaDefaultCallFlow = Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id}

                    $defaultCallFlowTargetTypeFriendly = "[Call Queue"
                    $defaultCallFlowTargetName = "$($MatchingCqAaDefaultCallFlow.Name)]"

                    if ($InvokedByNesting -eq $false) {
                        $aaDefaultCallFlowForwardsToCq = $true
                    }

                    else {
                        $aaNestedDefaultCallFlowForwardsToCq = $true
                    }

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

            $lastCallFlowAction = "defaultCallFlowAction$($aaDefaultCallFlowCounter)"

            $mdAutoAttendantDefaultCallFlow = "defaultCallFlowGreeting$($aaDefaultCallFlowCounter)>$defaultCallFlowGreeting] --> defaultCallFlow$($aaDefaultCallFlowCounter)($defaultCallFlowAction) --> $lastCallFlowAction($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName)"
            
            . Get-CallQueueCallFlow -MatchingCQIdentity $MatchingCQIdentity

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
        [Parameter(Mandatory=$false)][String]$VoiceAppId,
        [Parameter(Mandatory=$false)][Bool]$InvokedByNesting = $false,
        [Parameter(Mandatory=$false)][String]$NestedAaCallFlowType
    )

    if (!$aaAfterHoursCallFlowCounter) {
        $aaAfterHoursCallFlowCounter = 0
    }

    $aaAfterHoursCallFlowCounter ++

    # Get after hours call flow
    $afterHoursCallFlow = ($aa.CallFlows | Where-Object {$_.Name -Match "after hours"})
    $afterHoursCallFlowAction = ($aa.CallFlows | Where-Object {$_.Name -Match "after hours"}).Menu.MenuOptions.Action.Value

    # Get after hours greeting
    if (!$afterHoursCallFlow.Greetings.ActiveType.Value){
        $afterHoursCallFlowGreeting = "Greeting <br> None"
    }

    else {
        $afterHoursCallFlowGreeting = "Greeting <br> $($afterHoursCallFlow.Greetings.ActiveType.Value)"
    }

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

                    if ($InvokedByNesting -eq $false) {
                        $aaAfterHoursCallFlowForwardsToAa = $true
                        $aaAfterHoursCallFlowNestedAaIdentity = $MatchingAA.Identity

                    }

                }

                else {

                    $MatchingCqAaAfterHoursCallFlow = Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id}

                    $afterHoursCallFlowTargetTypeFriendly = "[Call Queue"
                    $afterHoursCallFlowTargetName = "$($MatchingCqAaAfterHoursCallFlow.Name)]"

                    if ($InvokedByNesting -eq $false) {
                        $aaAfterHoursCallFlowForwardsToCq = $true
                    }

                    else {
                        $aaNestedAfterHoursCallFlowForwardsToCq = $true
                    }

                }

            }
            SharedVoicemail {

                $afterHoursCallFlowTargetTypeFriendly = "Voicemail"
                $afterHoursCallFlowTargetName = (Get-MsolGroup -ObjectId $afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id).DisplayName

            }
        }

        # Check if transfer target type is call queue
        if ($afterHoursCallFlowTargetTypeFriendly -eq "[Call Queue") {

            $MatchingCQIdentity = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq ($aa.CallFlows.Menu | Where-Object {$_.Name -match "After hours"}).MenuOptions.CallTarget.Id}).Identity

            $lastCallFlowAction = "afterHoursCallFlowAction$($aaDefaultCallFlowCounter)"

            $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowCounter)>$AfterHoursCallFlowGreeting] --> AfterHoursCallFlow$($aaAfterHoursCallFlowCounter)($AfterHoursCallFlowAction) --> $lastCallFlowAction($AfterHoursCallFlowTargetTypeFriendly <br> $AfterHoursCallFlowTargetName)"

            . Get-CallQueueCallFlow -MatchingCQIdentity $MatchingCQIdentity
            
        } # End if transfer target type is call queue

        # Check if AfterHours callflow action target is trasnfer call to target but something other than call queue
        else {

            $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowCounter)>$AfterHoursCallFlowGreeting] --> AfterHoursCallFlow$($aaAfterHoursCallFlowCounter)($AfterHoursCallFlowAction) --> afterHoursCallFlowAction$($aaAfterHoursCallFlowCounter)($AfterHoursCallFlowTargetTypeFriendly <br> $AfterHoursCallFlowTargetName)"

        }


        # Mermaid code for after hours call flow nodes
        $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowCounter)>$AfterHoursCallFlowGreeting] --> afterHoursCallFlow$($aaAfterHoursCallFlowCounter)($afterHoursCallFlowAction) --> afterHoursCallFlowAction$($aaAfterHoursCallFlowCounter)($afterHoursCallFlowTargetTypeFriendly <br> $afterHoursCallFlowTargetName)"

    }

    elseif ($afterHoursCallFlowAction -eq "DisconnectCall") {

        $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowCounter)>$AfterHoursCallFlowGreeting] --> afterHoursCallFlow$($aaAfterHoursCallFlowCounter)(($afterHoursCallFlowAction))"

    }
    

    

}


function Get-CallQueueCallFlow {
    param (
        [Parameter(Mandatory=$true)][String]$MatchingCQIdentity,
        [Parameter(Mandatory=$false)][Bool]$InvokedByNesting,
        [Parameter(Mandatory=$false)][String]$NestedCQType

    )

    if (!$cqCallFlowCounter) {
        $cqCallFlowCounter = 0
    }

    $cqCallFlowCounter ++

    $MatchingCQ = Get-CsCallQueue -Identity $MatchingCQIdentity

    Write-Host "Function running for $($MatchingCQ.Name)" -ForegroundColor Magenta
    Write-Host "Function run number: $cqCallFlowCounter" -ForegroundColor Magenta
    Write-Host "Function running nested: $InvokedByNesting" -ForegroundColor Magenta

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

            $MatchingOverFlowCQ = $null

        }
        Forward {

            if ($MatchingCQ.OverflowActionTarget.Type -eq "User") {

                $MatchingOverFlowUser = (Get-MsolUser -ObjectId $MatchingCQ.OverflowActionTarget.Id).DisplayName

                $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqOverFlowActionTarget$($cqCallFlowCounter)(User <br> $MatchingOverFlowUser)"

                $MatchingOverFlowCQ = $null

            }

            elseif ($MatchingCQ.OverflowActionTarget.Type -eq "Phone") {

                $cqOverFlowPhoneNumber = ($MatchingCQ.OverflowActionTarget.Id).Replace("tel:","")

                $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqOverFlowActionTarget$($cqCallFlowCounter)(External Number <br> $cqOverFlowPhoneNumber)"

                $MatchingOverFlowCQ = $null
                
            }

            else {

                $MatchingOverFlowAA = (Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.OverflowActionTarget.Id}).Name

                if ($MatchingOverFlowAA) {

                    $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqOverFlowActionTarget$($cqCallFlowCounter)([Auto Attendant <br> $MatchingOverFlowAA])"

                    $MatchingOverFlowCQ = $null

                }

                else {

                    $MatchingOverFlowCQ = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.OverflowActionTarget.Id})

                    $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqOverFlowActionTarget$($cqCallFlowCounter)([Call Queue <br> $($MatchingOverFlowCQ.Name)])"

                }

            }

        }
        SharedVoicemail {
            $MatchingOverFlowVoicemail = (Get-MsolGroup -ObjectId $MatchingCQ.OverflowActionTarget.Id).DisplayName

            $MatchingOverFlowCQ = $null

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

            $MatchingTimeoutCQ = $null

        }
        Forward {
    
            if ($MatchingCQ.TimeoutActionTarget.Type -eq "User") {

                $MatchingTimeoutUser = (Get-MsolUser -ObjectId $MatchingCQ.TimeoutActionTarget.Id).DisplayName
    
                $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqTimeoutActionTarget$($cqCallFlowCounter)(User <br> $MatchingTimeoutUser)"

                $MatchingTimeoutCQ = $null
    
            }
    
            elseif ($MatchingCQ.TimeoutActionTarget.Type -eq "Phone") {
    
                $cqTimeoutPhoneNumber = ($MatchingCQ.TimeoutActionTarget.Id).Replace("tel:","")
    
                $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqTimeoutActionTarget$($cqCallFlowCounter)(External Number <br> $cqTimeoutPhoneNumber)"

                $MatchingTimeoutCQ = $null
                
            }
    
            else {
    
                $MatchingTimeoutAA = (Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.TimeoutActionTarget.Id}).Name
    
                if ($MatchingTimeoutAA) {
    
                    $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqTimeoutActionTarget$($cqCallFlowCounter)([Auto Attendant <br> $MatchingTimeoutAA])"

                    $MatchingTimeoutCQ = $null
    
                }
    
                else {
    
                    $MatchingTimeoutCQ = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.TimeoutActionTarget.Id})

                    Write-Host "Matching Time Out CQ Name: $($MatchingTimeoutCQ.Name)"

                    $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowCounter)(TransferCallToTarget) --> cqTimeoutActionTarget$($cqCallFlowCounter)([Call Queue <br> $($MatchingTimeoutCQ.Name)])"
    
                }
    
            }
    
        }
        SharedVoicemail {
            $MatchingTimeoutVoicemail = (Get-MsolGroup -ObjectId $MatchingCQ.TimeoutActionTarget.Id).DisplayName

            $MatchingTimeoutCQ = $null
    
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


    if ($InvokedByNesting -eq $true -and $NestedCQType -eq "OverFlow") {

        Write-Host "proof that this is working... " -ForegroundColor Green

        if ($cqCallFlowCounter -eq 3) {
            $lastCallFlowAction = "cqOverFlowActionTarget$($cqCallFlowCounter -2)"
        }
        else {
            $lastCallFlowAction = "cqOverFlowActionTarget$($cqCallFlowCounter -1)"
        }


    }

    if ($InvokedByNesting -eq $true -and $NestedCQType -eq "TimeOut") {

        Write-Host "proof that this is working... " -ForegroundColor Green

        if ($cqCallFlowCounter -eq 3) {
            $lastCallFlowAction = "cqTimeoutActionTarget$($cqCallFlowCounter -2)"
        }
        else {
            $lastCallFlowAction = "cqTimeoutActionTarget$($cqCallFlowCounter -1)"
        }


    }

    else {

        #$lastCallFlowAction = $null

    }

    
    # Create default callflow mermaid code

$mdCallQueueCallFlow =@"
$lastCallFlowAction --> cqGreeting$($cqCallFlowCounter)>Greeting <br> $CqGreeting] --> overFlow$($cqCallFlowCounter){More than $CqOverFlowThreshold <br> Active Calls}
overFlow$($cqCallFlowCounter) ---> |Yes| $CqOverFlowActionFriendly
overFlow$($cqCallFlowCounter) ---> |No| routingMethod$($cqCallFlowCounter)

$nestedCallQueueTopLevelNumbers

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

$mermaidCode += $mdCallQueueCallFlow

if ($InvokedByNesting -eq $false) {

    $InitialMatchingCQ = $MatchingCQ

    Write-Host "Initial Function Run for $($InitialMatchingCQ.Name)" -ForegroundColor DarkGreen

    if ($MatchingOverFlowCQ -or $MatchingTimeoutCQ) {

        Write-Host "Initial CQ time out or forward action is another CQ" -ForegroundColor DarkGreen

        $InitialMatchingCqTimeoutActionTarget = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.TimeoutActionTarget.Id}).Identity
        $InitialMatchingCqOverFlowActionTarget = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -contains $MatchingCQ.OverFlowActionTarget.Id}).Identity  

        Write-Host "Getting Over Flow CQ..."

        . Get-CallQueueCallFlow -MatchingCQIdentity $InitialMatchingCqOverFlowActionTarget -InvokedByNesting $true -NestedCQType "OverFlow"

        Write-Host "Getting Time Out CQ..."

        . Get-CallQueueCallFlow -MatchingCQIdentity $InitialMatchingCqTimeoutActionTarget -InvokedByNesting $true -NestedCQType "TimeOut"

   
    }


}
  
}


. Set-Mermaid -DocType $DocType

function Get-CallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceAppId
    )
    
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
            $nodeElementHolidayLink --> $mdAutoAttendantDefaultCallFlow
            
"@

            $mermaidCode += $mdHolidayAndAfterHoursCheck
    
        }
    
    
    }
    
    elseif ($voiceAppType -eq "Call Queue") {
        . Get-CallQueueCallFlow -MatchingCQIdentity $VoiceApp.Identity
    }

    
}

# Get First Call Flow
. Get-CallFlow

function Get-AutoAttendantCallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceAppId,
        [Parameter(Mandatory=$false)][Bool]$InvokedByNesting,
        [Parameter(Mandatory=$false)][String]$NestedAaCallFlowType
    )
    
    . Find-Holidays -VoiceAppId $VoiceAppId
    . Find-AfterHours -VoiceAppId $VoiceAppId

    if ($aaHasHolidays -eq $true -or $aaHasAfterHours -eq $true) {

        . Get-AutoAttendantDefaultCallFlow -VoiceAppId $VoiceAppId -InvokedByNesting $true -NestedAaCallFlowType $NestedAaCallFlowType

        . Get-AutoAttendantAfterHoursCallFlow -VoiceAppId $VoiceAppId -InvokedByNesting $true -NestedAaCallFlowType $NestedAaCallFlowType

        . Get-AutoAttendantHolidaysAndAfterHours -VoiceAppId $VoiceAppId -InvokedByNesting $true -NestedAaCallFlowType $NestedAaCallFlowType

    }

    else {

        . Get-AutoAttendantDefaultCallFlow -VoiceAppId $VoiceAppId -InvokedByNesting $true -NestedAaCallFlowType $NestedAaCallFlowType

        if ($InvokedByNesting -eq $true -and $NestedAaCallFlowType -eq "AfterHours") {
            $nodeElementHolidayLink = "afterHoursCallFlowAction$($aaCounter)"

            Write-Host "Proof this works too!" -ForegroundColor DarkYellow
        }

        if ($InvokedByNesting -eq $true -and $NestedAaCallFlowType -eq "Default") {
            $nodeElementHolidayLink = "defaultCallFlowAction$($aaCounter)"

            Write-Host "Proof this works too!" -ForegroundColor DarkYellow
        }

        $mdHolidayAndAfterHoursCheck =@"
        $nodeElementHolidayLink --> $mdAutoAttendantDefaultCallFlow
        
"@

        $mermaidCode += $mdHolidayAndAfterHoursCheck

    }


}

if ($aaAfterHoursCallFlowForwardsToAa) {

    . Get-AutoAttendantCallFlow -VoiceAppId $aaAfterHoursCallFlowNestedAaIdentity -InvokedByNesting $true -NestedAaCallFlowType "AfterHours"

}

if ($aaDefaultCallFlowForwardsToAa) {

    . Get-AutoAttendantCallFlow -VoiceAppId $aaDefaultCallFlowNestedAaIdentity -InvokedByNesting $true -NestedAaCallFlowType "Default"

}


Set-Content -Path ".\$(($VoiceApp.Name).Replace(" ","_"))_CallFlow$fileExtension" -Value $mermaidCode -Encoding UTF8

if ($SetClipBoard -eq $true) {
    $mermaidCode -Replace('```mermaid','') `
    -Replace('```','') | Set-Clipboard

    Write-Host "Mermaid code copied to clipboard. Paste it on https://mermaid.live" -ForegroundColor Cyan
}
