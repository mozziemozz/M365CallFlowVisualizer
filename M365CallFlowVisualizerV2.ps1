<#
    .SYNOPSIS
    Reads the configuration from a Microsoft 365 Phone System auto attendant or call queue and visualizes the call flow using mermaid-js.

    .DESCRIPTION
    Reads the configuration from a Microsoft 365 Phone System auto attendant or call queue by either specifying the voice app name and type, unique identity of the voice app or presents a selection of available auto attendants or call queues if none of the identifiers are supplied.
    The call flow is then written into either a mermaid (*.mmd) or a markdown (*.md) file containing the mermaid syntax.

    Author:             Martin Heusser
    Version:            2.1.0
    Revision:
        20.10.2021:     Creation
        21.10.2021:     Add comments and streamline code, add longer arrow links for default call flow desicion node
        21.10.2021:     Add support for top level call queues (besides auto attendants)
        21.10.2021:     Move call queue specific operations into a function
        24.10.2021:     Fixed a bug where Disconnect Call was not reflected in mermaid correctly when CQ timeout action was disconnect call
        30.10.2021:     V2: most of the script logic was moved into functions. Added parameters for specifig resource account (specified by phone number), added support for nested queues, added support to display only 1 queue if timeout and overflow go to the same queue.
        01.11.2021:     Add support to display call queues for an after hours call flow of an auto attendant
        01.11.2021:     Fix issue where additional entry point numbers were not shown on after hours call flow call queues
        02.11.2021:     Add support for nested Auto Attendants
        04.01.2022:     V2.1 more or less a complete rewrite of the script logic to make it really dynamic and support indefinite chaning/nesting of voice apps
                        Add support to disable rendering of nested voice apps
                        Add support for voice app name and type parameters
                        Fixed a bug where some phone numbers which contained extensions including a ";" were not rendered in mermaid. (replace ";" with ",")

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

    -DisplayNestedCallFlows
        Specifies whether or not to also display the call flows of nested call queues or auto attendants. If set to false, only the name of nested voice apps will be rendered. Nested call flows won't be expanded.
        Required:           false
        Type:               boolean
        Default value:      true   

    -DocType
        Specifies the document type.
        Required:           false
        Type:               string
        Accepted values:    Markdown, Mermaid
        Default value:      Mermaid
    
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

    .EXAMPLE
        .\M365CallFlowVisualizerV2.ps1 -Identity "6fb84b40-f045-45e8-8c1a-8fc18188e46x"

    .EXAMPLE
        .\M365CallFlowVisualizerV2.ps1 -VoiceAppName "PS Test AA" -VoiceAppType "Auto Attendant"

    .EXAMPLE
        .\M365CallFlowVisualizerV2.ps1 -VoiceAppName "PS Test CQ" -VoiceAppType "Call Queue"

    .EXAMPLE
        .\M365CallFlowVisualizerV2.ps1 -DocType Markdown -SetClipBoard $false

    .LINK
    https://github.com/mozziemozz/M365CallFlowVisualizer
    
#>

#Requires -Modules MsOnline, MicrosoftTeams

[CmdletBinding(DefaultParametersetName="None")]
param(
    [Parameter(Mandatory=$false)][String]$Identity,
    [Parameter(Mandatory=$false)][Bool]$SetClipBoard = $true,
    [Parameter(Mandatory=$false)][Bool]$DisplayNestedCallFlows = $true,
    [Parameter(Mandatory=$false)][ValidateSet("Markdown","Mermaid")][String]$DocType = "Mermaid",
    [Parameter(ParameterSetName="VoiceAppProperties",Mandatory=$false)][String]$VoiceAppName,
    [Parameter(ParameterSetName="VoiceAppProperties",Mandatory=$true)][ValidateSet("Auto Attendant","Call Queue")][String]$VoiceAppType
)

$processedVoiceApps = @()

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

    Write-Host "Getting call flow for: $($aa.Name)" -ForegroundColor Magenta
    Write-Host "Voice App Id: $($aa.Identity)" -ForegroundColor Magenta

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
        [Parameter(Mandatory=$true)][String]$VoiceAppId
    )

    $aaObjectId = $aa.Identity

    $holidaySubgraphName = "Holiday-Table-$($aa.Name.Replace(" ","-"))"

    if ($aaHasHolidays -eq $true) {

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

                $holidayGreeting = "Greeting <br> $($holidayCallFlow.Greetings.ActiveType.Value)"

            }

            $holidayAction = $holidayCallFlow.Menu.MenuOptions.Action.Value

            # Check if holiday call handling is disconnect call
            if ($holidayAction -eq "DisconnectCall") {

                $nodeElementHolidayAction = "elementAAHolidayAction$($aaObjectId)-$($HolidayCounter)(($holidayAction))"

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
                $nodeElementHolidayAction = "elementAAHolidayAction$($aaObjectId)-$($HolidayCounter)($holidayAction) --> elementAAHolidayActionTargetType$($aaObjectId)-$($HolidayCounter)($holidayActionTargetTypeFriendly <br> $holidayActionTargetName)"

            }

            # Create subgraph per holiday call handling inside the Holidays subgraph
            $nodeElementHolidayDetails =@"

subgraph $($holidayCallFlow.Name)
direction LR
elementAAHoliday$($aaObjectId)-$($HolidayCounter)(Schedule <br> $($holidaySchedule.FixedSchedule.DateTimeRanges.Start) <br> $($holidaySchedule.FixedSchedule.DateTimeRanges.End)) --> elementAAHolidayGreeting$($aaObjectId)-$($HolidayCounter)>$holidayGreeting] --> $nodeElementHolidayAction
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
        $nodeElementHolidayCheck = "elementHolidayCheck$($aaObjectId){During Holiday?}"
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
        $nodeElementAfterHoursCheck = "elementAfterHoursCheck$($aaObjectId){During Business Hours? <br> Time Zone: $aaTimeZone <br> $mondayHours <br> $tuesdayHours  <br> $wednesdayHours  <br> $thursdayHours <br> $fridayHours <br> $saturdayHours <br> $sundayHours}"

    } # End if aa has after hours

    $nodeElementHolidayLink = "$($aa.Identity)([Auto Attendant <br> $($aa.Name)])"

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
                $defaultCallFlowTargetUser = (Get-MsolUser -ObjectId $($aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id))
                $defaultCallFlowTargetName = $defaultCallFlowTargetUser.DisplayName
                $defaultCallFlowTargetIdentity = $defaultCallFlowTargetUser.ObjectId
            }
            ExternalPstn { 
                $defaultCallFlowTargetTypeFriendly = "External PSTN"
                $defaultCallFlowTargetName = ($aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id).Replace("tel:","")
                $defaultCallFlowTargetIdentity = $defaultCallFlowTargetName
            }
            ApplicationEndpoint {

                # Check if application endpoint is auto attendant or call queue
                $MatchingAaDefaultCallFlowAa = Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id}

                if ($MatchingAaDefaultCallFlowAa) {

                    $defaultCallFlowTargetTypeFriendly = "[Auto Attendant"
                    $defaultCallFlowTargetName = "$($MatchingAaDefaultCallFlowAa.Name)]"

                }

                else {

                    $MatchingCqAaDefaultCallFlow = Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id}

                    $defaultCallFlowTargetTypeFriendly = "[Call Queue"
                    $defaultCallFlowTargetName = "$($MatchingCqAaDefaultCallFlow.Name)]"

                }

            }
            SharedVoicemail {

                $defaultCallFlowTargetTypeFriendly = "Voicemail"
                $defaultCallFlowTargetGroup = (Get-MsolGroup -ObjectId $aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id)
                $defaultCallFlowTargetName = $defaultCallFlowTargetGroup.DisplayName
                $defaultCallFlowTargetIdentity = $defaultCallFlowTargetGroup.ObjectId

            }
        }

        ### move this part to switch statement!

        # Check if transfer target type is call queue
        if ($defaultCallFlowTargetTypeFriendly -eq "[Call Queue") {

            $MatchingCQIdentity = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $aa.DefaultCallFlow.Menu.MenuOptions.CallTarget.Id}).Identity

            $mdAutoAttendantDefaultCallFlow = "defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowGreeting] --> defaultCallFlow$($aaDefaultCallFlowAaObjectId)($defaultCallFlowAction) --> $($MatchingCQIdentity)($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName)"
            
            if ($nestedVoiceApps -notcontains $MatchingCQIdentity) {

                $nestedVoiceApps += $MatchingCQIdentity

            }

        
        } # End if transfer target type is call queue

        elseif ($defaultCallFlowTargetTypeFriendly -eq "[Auto Attendant") {

            $mdAutoAttendantDefaultCallFlow = "defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowGreeting] --> defaultCallFlow$($aaDefaultCallFlowAaObjectId)($defaultCallFlowAction) --> $($MatchingAaDefaultCallFlowAa.Identity)($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName)"

            if ($nestedVoiceApps -notcontains $MatchingAaDefaultCallFlowAa.Identity) {

                $nestedVoiceApps += $MatchingAaDefaultCallFlowAa.Identity

            }

        }

        # Check if default callflow action target is trasnfer call to target but something other than call queue
        else {

            $mdAutoAttendantDefaultCallFlow = "defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowGreeting] --> defaultCallFlow$($aaDefaultCallFlowAaObjectId)($defaultCallFlowAction) --> $($defaultCallFlowTargetIdentity)($defaultCallFlowTargetTypeFriendly <br> $defaultCallFlowTargetName)"

        }

    }

    # Check if default callflow action is disconnect call
    elseif ($defaultCallFlowAction -eq "DisconnectCall") {

        $mdAutoAttendantDefaultCallFlow = "defaultCallFlowGreeting$($aaDefaultCallFlowAaObjectId)>$defaultCallFlowGreeting] --> defaultCallFlow$($aaDefaultCallFlowAaObjectId)(($defaultCallFlowAction))"

    }
    
    
}

function Get-AutoAttendantAfterHoursCallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceAppId
    )

    $aaAfterHoursCallFlowAaObjectId = $aa.Identity

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
                $afterHoursCallFlowTargetUser = (Get-MsolUser -ObjectId $($afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id))
                $afterHoursCallFlowTargetName = $afterHoursCallFlowTargetUser.DisplayName
                $afterHoursCallFlowTargetIdentity = $afterHoursCallFlowTargetUser.ObjectId            
            }
            ExternalPstn { 
                $afterHoursCallFlowTargetTypeFriendly = "External PSTN"
                $afterHoursCallFlowTargetName = ($afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id).Replace("tel:","")
                $afterHoursCallFlowTargetIdentity = $afterHoursCallFlowTargetName
            }
            ApplicationEndpoint {

                # Check if application endpoint is an auto attendant or a call queue
                $MatchingAaAfterHoursCallFlowAa = Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -match $afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id}

                if ($MatchingAaAfterHoursCallFlowAa) {

                    $afterHoursCallFlowTargetTypeFriendly = "[Auto Attendant"
                    $afterHoursCallFlowTargetName = "$($MatchingAaAfterHoursCallFlowAa.Name)]"

                }

                else {

                    $MatchingCqAaAfterHoursCallFlow = Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id}

                    $afterHoursCallFlowTargetTypeFriendly = "[Call Queue"
                    $afterHoursCallFlowTargetName = "$($MatchingCqAaAfterHoursCallFlow.Name)]"

                }

            }
            SharedVoicemail {

                $afterHoursCallFlowTargetTypeFriendly = "Voicemail"
                $afterHoursCallFlowTargetGroup = (Get-MsolGroup -ObjectId $afterHoursCallFlow.Menu.MenuOptions.CallTarget.Id)
                $afterHoursCallFlowTargetName = $afterHoursCallFlowTargetGroup.DisplayName
                $afterHoursCallFlowTargetIdentity = $afterHoursCallFlowTargetGroup.ObjectId

            }
        }

        # Check if transfer target type is call queue
        if ($afterHoursCallFlowTargetTypeFriendly -eq "[Call Queue") {

            $MatchingCQIdentity = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq ($aa.CallFlows.Menu | Where-Object {$_.Name -match "After hours"}).MenuOptions.CallTarget.Id}).Identity

            $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId)>$AfterHoursCallFlowGreeting] --> AfterHoursCallFlow$($aaAfterHoursCallFlowAaObjectId)($AfterHoursCallFlowAction) --> $($MatchingCQIdentity)($AfterHoursCallFlowTargetTypeFriendly <br> $AfterHoursCallFlowTargetName)"

            if ($nestedVoiceApps -notcontains $MatchingCQIdentity) {

                $nestedVoiceApps += $MatchingCQIdentity

            }

            
        } # End if transfer target type is call queue

        elseif ($afterHoursCallFlowTargetTypeFriendly -eq "[Auto Attendant") {

            $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId)>$AfterHoursCallFlowGreeting] --> AfterHoursCallFlow$($aaAfterHoursCallFlowAaObjectId)($AfterHoursCallFlowAction) --> $($MatchingAaAfterHoursCallFlowAa.Identity)($AfterHoursCallFlowTargetTypeFriendly <br> $AfterHoursCallFlowTargetName)"

            if ($nestedVoiceApps -notcontains $MatchingAaAfterHoursCallFlowAa.Identity) {

                $nestedVoiceApps += $MatchingAaAfterHoursCallFlowAa.Identity

            }

        }

        # Check if AfterHours callflow action target is trasnfer call to target but something other than call queue
        else {

            $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId)>$AfterHoursCallFlowGreeting] --> AfterHoursCallFlow$($aaAfterHoursCallFlowAaObjectId)($AfterHoursCallFlowAction) --> $($afterHoursCallFlowTargetIdentity)($AfterHoursCallFlowTargetTypeFriendly <br> $AfterHoursCallFlowTargetName)"
            
        }

    }

    #Maybe combine all disconnect calls into one node...
    elseif ($afterHoursCallFlowAction -eq "DisconnectCall") {

        $mdAutoAttendantAfterHoursCallFlow = "afterHoursCallFlowGreeting$($aaAfterHoursCallFlowAaObjectId)>$AfterHoursCallFlowGreeting] --> afterHoursCallFlow$($aaAfterHoursCallFlowAaObjectId)(($afterHoursCallFlowAction))"

    }
    

    

}


function Get-CallQueueCallFlow {
    param (
        [Parameter(Mandatory=$true)][String]$MatchingCQIdentity
    )

    $MatchingCQ = Get-CsCallQueue -Identity $MatchingCQIdentity

    $cqCallFlowObjectId = $MatchingCQ.Identity

    Write-Host "Getting call flow for: $($MatchingCQ.Name)" -ForegroundColor Magenta
    Write-Host "Voice App Id: $cqCallFlowObjectId" -ForegroundColor Magenta

    # Store all neccessary call queue properties in variables
    $CqName = $MatchingCQ.Name
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
            $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)((Disconnect Call))"

        }
        Forward {

            if ($MatchingCQ.OverflowActionTarget.Type -eq "User") {

                $MatchingOverFlowUserProperties = (Get-MsolUser -ObjectId $MatchingCQ.OverflowActionTarget.Id)
                $MatchingOverFlowUser = $MatchingOverFlowUserProperties.DisplayName
                $MatchingOverFlowIdentity = $MatchingOverFlowUserProperties.ObjectId

                $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingOverFlowIdentity)(User <br> $MatchingOverFlowUser)"

            }

            elseif ($MatchingCQ.OverflowActionTarget.Type -eq "Phone") {

                $cqOverFlowPhoneNumber = ($MatchingCQ.OverflowActionTarget.Id).Replace("tel:","")

                $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($cqOverFlowPhoneNumber)(External Number <br> $cqOverFlowPhoneNumber)"
                
            }

            else {

                $MatchingOverFlowAA = (Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.OverflowActionTarget.Id})

                if ($MatchingOverFlowAA) {

                    $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingOverFlowAA.Identity)([Auto Attendant <br> $($MatchingOverFlowAA.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingOverFlowAA.Identity) {

                        $nestedVoiceApps += $MatchingOverFlowAA.Identity
        
                    }
        

                }

                else {

                    $MatchingOverFlowCQ = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.OverflowActionTarget.Id})

                    $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingOverFlowCQ.Identity)([Call Queue <br> $($MatchingOverFlowCQ.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingOverFlowCQ.Identity) {

                        $nestedVoiceApps += $MatchingOverFlowCQ.Identity
        
                    }

                }

            }

        }
        SharedVoicemail {
            $MatchingOverFlowVoicemailProperties = (Get-MsolGroup -ObjectId $MatchingCQ.OverflowActionTarget.Id)
            $MatchingOverFlowVoicemail = $MatchingOverFlowVoicemailProperties.DisplayName
            $MatchingOverFlowIdentity = $MatchingOverFlowVoicemailProperties.ObjectId

            if ($MatchingCQ.OverflowSharedVoicemailTextToSpeechPrompt) {

                $CqOverFlowVoicemailGreeting = "TextToSpeech"

            }

            else {

                $CqOverFlowVoicemailGreeting = "AudioFile"

            }

            $CqOverFlowActionFriendly = "cqOverFlowAction$($cqCallFlowObjectId)(TransferCallToTarget) --> cqOverFlowVoicemailGreeting$($cqCallFlowObjectId)>Greeting <br> $CqOverFlowVoicemailGreeting] --> $($MatchingOverFlowIdentity)(Shared Voicemail <br> $MatchingOverFlowVoicemail)"

        }

    }

    # Switch through call queue timeout overflow action
    switch ($CqTimeoutAction) {
        Disconnect {
            $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)((Disconnect Call))"

        }
        Forward {
    
            if ($MatchingCQ.TimeoutActionTarget.Type -eq "User") {

                $MatchingTimeoutUserProperties = (Get-MsolUser -ObjectId $MatchingCQ.TimeoutActionTarget.Id)
                $MatchingTimeoutUser = $MatchingTimeoutUserProperties.DisplayName
                $MatchingTimeoutIdentity = $MatchingTimeoutUserProperties.ObjectId
    
                $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingTimeoutIdentity)(User <br> $MatchingTimeoutUser)"
    
            }
    
            elseif ($MatchingCQ.TimeoutActionTarget.Type -eq "Phone") {
    
                $cqTimeoutPhoneNumber = ($MatchingCQ.TimeoutActionTarget.Id).Replace("tel:","")
    
                $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($cqTimeoutPhoneNumber)(External Number <br> $cqTimeoutPhoneNumber)"
                
            }
    
            else {
    
                $MatchingTimeoutAA = (Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.TimeoutActionTarget.Id})
    
                if ($MatchingTimeoutAA) {
    
                    $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingTimeoutAA.Identity)([Auto Attendant <br> $($MatchingTimeoutAA.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingTimeoutAA.Identity) {

                        $nestedVoiceApps += $MatchingTimeoutAA.Identity
        
                    }

    
                }
    
                else {
    
                    $MatchingTimeoutCQ = (Get-CsCallQueue | Where-Object {$_.ApplicationInstances -eq $MatchingCQ.TimeoutActionTarget.Id})

                    Write-Host "Matching Time Out CQ Name: $($MatchingTimeoutCQ.Name)"

                    $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget) --> $($MatchingTimeoutCQ.Identity)([Call Queue <br> $($MatchingTimeoutCQ.Name)])"

                    if ($nestedVoiceApps -notcontains $MatchingTimeoutCQ.Identity) {

                        $nestedVoiceApps += $MatchingTimeoutCQ.Identity
        
                    }
    
                }
    
            }
    
        }
        SharedVoicemail {
            $MatchingTimeoutVoicemailProperties = (Get-MsolGroup -ObjectId $MatchingCQ.TimeoutActionTarget.Id)
            $MatchingTimeoutVoicemail = $MatchingTimeoutVoicemailProperties.DisplayName
            $MatchingTimeoutIdentity = $MatchingTimeoutVoicemailProperties.ObjectId
    
            if ($MatchingCQ.TimeoutSharedVoicemailTextToSpeechPrompt) {
    
                $CqTimeoutVoicemailGreeting = "TextToSpeech"
    
            }
    
            else {
    
                $CqTimeoutVoicemailGreeting = "AudioFile"
    
            }
    
            $CqTimeoutActionFriendly = "cqTimeoutAction$($cqCallFlowObjectId)(TransferCallToTarget) --> cqTimeoutVoicemailGreeting$($cqCallFlowObjectId)>Greeting <br> $CqTimeoutVoicemailGreeting] --> $($MatchingTimeoutIdentity)(Shared Voicemail <br> $MatchingTimeoutVoicemail)"
    
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

        $AgentDisplayNames = "agentListType$($cqCallFlowObjectId) --> agent$($cqCallFlowObjectId)$($AgentCounter)($AgentDisplayName) --> timeOut$($cqCallFlowObjectId)`n"

        $mdCqAgentsDisplayNames += $AgentDisplayNames

        $AgentCounter ++
    }

    $lastCallFlowAction = "$($MatchingCQIdentity)([Call Queue <br> $($CqName)])"

    
    # Create default callflow mermaid code

$mdCallQueueCallFlow =@"
$lastCallFlowAction --> cqGreeting$($cqCallFlowObjectId)>Greeting <br> $CqGreeting] --> overFlow$($cqCallFlowObjectId){More than $CqOverFlowThreshold <br> Active Calls?}
overFlow$($cqCallFlowObjectId) ---> |Yes| $CqOverFlowActionFriendly
overFlow$($cqCallFlowObjectId) ---> |No| routingMethod$($cqCallFlowObjectId)

$nestedCallQueueTopLevelNumbers

subgraph Call Distribution
subgraph CQ Settings
routingMethod$($cqCallFlowObjectId)[(Routing Method: $CqRoutingMethod)] --> agentAlertTime$($cqCallFlowObjectId)
agentAlertTime$($cqCallFlowObjectId)[(Agent Alert Time: $CqAgentAlertTime)] -.- cqMusicOnHold$($cqCallFlowObjectId)
cqMusicOnHold$($cqCallFlowObjectId)[(Music On Hold: $CqMusicOnHold)] -.- conferenceMode$($cqCallFlowObjectId)
conferenceMode$($cqCallFlowObjectId)[(Conference Mode Enabled: $CqConferenceMode)] -.- agentOptOut$($cqCallFlowObjectId)
agentOptOut$($cqCallFlowObjectId)[(Agent Opt Out Allowed: $CqAgentOptOut)] -.- presenceBasedRouting$($cqCallFlowObjectId)
presenceBasedRouting$($cqCallFlowObjectId)[(Presence Based Routing: $CqPresenceBasedRouting)] -.- timeOut$($cqCallFlowObjectId)
timeOut$($cqCallFlowObjectId)[(Timeout: $CqTimeOut Seconds)]
end
subgraph Agents $($MatchingCQ.Name)
agentAlertTime$($cqCallFlowObjectId) --> agentListType$($cqCallFlowObjectId)[(Agent List Type: $CqAgentListType)]
$mdCqAgentsDisplayNames
end
end

timeOut$($cqCallFlowObjectId) --> cqResult$($cqCallFlowObjectId){Call Connected?}
cqResult$($cqCallFlowObjectId) --> |Yes| cqEnd$($cqCallFlowObjectId)((Call Connected))
cqResult$($cqCallFlowObjectId) --> |No| $CqTimeoutActionFriendly

"@

if ($mermaidCode -notcontains $mdCallQueueCallFlow) {

    $mermaidCode += $mdCallQueueCallFlow

}
  
}


. Set-Mermaid -DocType $DocType

#This is needed to determine if the Get-CallFlow function is running for the first time or not.
$mdNodePhoneNumbersCounter = 0

#This array stores information about the voice app's forwading targets.
$nestedVoiceApps = @()

function Get-CallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceAppId,
        [Parameter(Mandatory=$false)][String]$VoiceAppName,
        [Parameter(Mandatory=$false)][String]$voiceAppType
    )
    
    if (!$VoiceAppName -and !$voiceAppType -and !$VoiceAppId) {
        
        $VoiceApps = @()

        $VoiceAppAas = Get-CsAutoAttendant
        $VoiceAppCqs = Get-CsCallQueue

        foreach ($VoiceApp in $VoiceAppAas) {

            $VoiceAppProperties = New-Object -TypeName psobject

            $VoiceAppProperties | Add-Member -MemberType NoteProperty -Name "Name" -Value $VoiceApp.Name
            $VoiceAppProperties | Add-Member -MemberType NoteProperty -Name "Type" -Value "Auto Attendant"

            $VoiceApps += $VoiceAppProperties

        }

        foreach ($VoiceApp in $VoiceAppCqs) {

            $VoiceAppProperties = New-Object -TypeName psobject

            $VoiceAppProperties | Add-Member -MemberType NoteProperty -Name "Name" -Value $VoiceApp.Name
            $VoiceAppProperties | Add-Member -MemberType NoteProperty -Name "Type" -Value "Call Queue"

            $VoiceApps += $VoiceAppProperties

        }

        $VoiceAppSelection = $VoiceApps | Out-GridView -Title "Choose an Auto Attendant or Call Queue from the list." -PassThru

        if ($VoiceAppSelection.Type -eq "Auto Attendant") {

            $VoiceApp = Get-CsAutoAttendant | Where-Object {$_.Name -eq $VoiceAppSelection.Name}
            $voiceAppType = "Auto Attendant"

        }

        else {

            $VoiceApp = Get-CsCallQueue | Where-Object {$_.Name -eq $VoiceAppSelection.Name}
            $voiceAppType = "Call Queue"

        }


    }

    elseif ($VoiceAppId) {

        try {
            $VoiceApp = Get-CsAutoAttendant -Identity $VoiceAppId
            $voiceAppType = "Auto Attendant"
        }
        catch {
            $VoiceApp = Get-CsCallQueue -Identity $VoiceAppId
            $voiceAppType = "Call Queue"
        }

    }

    else {

        if ($voiceAppType -eq "Auto Attendant") {

            $VoiceApp = Get-CsAutoAttendant | Where-Object {$_.Name -eq $VoiceAppName}

        }

        else {

            $VoiceApp = Get-CsCallQueue | Where-Object {$_.Name -eq $VoiceAppName}

        }

    }

    $mdNodePhoneNumbers = @()

    foreach ($ApplicationInstance in ($VoiceApp.ApplicationInstances)) {

        if ($mdNodePhoneNumbersCounter -eq 0) {

            $mdPhoneNumberLinkType = "-->"

        }

        else {

            $mdPhoneNumberLinkType = "-.->"

        }

        $ApplicationInstancePhoneNumber = ((Get-CsOnlineApplicationInstance -Identity $ApplicationInstance).PhoneNumber) -replace ("tel:","")

        if ($ApplicationInstancePhoneNumber) {

            $mdNodeNumber = "start$($ApplicationInstancePhoneNumber)((Incoming Call at <br> $ApplicationInstancePhoneNumber)) $mdPhoneNumberLinkType $($VoiceApp.Identity)([$($voiceAppType) <br> $($VoiceApp.Name)])"

            $mdNodePhoneNumbers += $mdNodeNumber
    
            $mdNodePhoneNumbersCounter ++

        }


    }

    if ($mermaidCode -notcontains $mdNodePhoneNumbers) {

        $mermaidCode += $mdNodePhoneNumbers

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

            $nodeElementHolidayLink = "$($aa.Identity)([Auto Attendant <br> $($aa.Name)])"
    
            $mdHolidayAndAfterHoursCheck =@"
            $nodeElementHolidayLink --> $mdAutoAttendantDefaultCallFlow
            
"@

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

            . Get-CallFlow -VoiceAppId $nestedVoiceApp

        }

    }

    if (Compare-Object -ReferenceObject $nestedVoiceApps -DifferenceObject $processedVoiceApps) {

        . Get-NestedCallFlow

    }

}

if ($DisplayNestedCallFlows -eq $true) {

    . Get-NestedCallFlow

}

else {
    
    if ($nestedVoiceApps) {

        Write-Host "Your call flow contains nested call queues or auto attendants. They won't be expanded because 'DisplayNestedCallFlows' is set to false." -ForegroundColor Yellow
        Write-Host "Nested Voice App Ids:" -ForegroundColor Yellow
        $nestedVoiceApps

    }

}


#Remove invalid characters from mermaid syntax
$mermaidCode = $mermaidCode.Replace(";",",")

Set-Content -Path ".\$(($VoiceApp.Name).Replace(" ","_"))_CallFlow$fileExtension" -Value $mermaidCode -Encoding UTF8

if ($SetClipBoard -eq $true) {
    $mermaidCode -Replace('```mermaid','') `
    -Replace('```','') | Set-Clipboard

    Write-Host "Mermaid code copied to clipboard. Paste it on https://mermaid.live" -ForegroundColor Cyan
}