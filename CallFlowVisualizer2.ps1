#Requires -Modules MsOnline, MicrosoftTeams

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)][ValidateSet("Markdown","Mermaid")][String]$docType = "Markdown",
    [Parameter(Mandatory=$false)][Int32]$ShowNestedDepth = 1,
    [Parameter(Mandatory=$false)][Switch]$SubSequentRun,
    [Parameter(Mandatory=$false)][string]$PhoneNumber

)

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
    $mermaidCode += $mdHolidayCheck
    $mermaidCode += $mdEnd
    
}

function Get-VoiceApp {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceApp
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

        $voiceAppCounter = 0
        $voiceAppCounter ++
        $resourceAccountCounter = 1

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

function Find-BusinessHours {
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

function Get-AutoAttendantCallFlow {
    param (
    )

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

    # Check if auto attendant has after hours and holidays
    if ($aaHasAfterHours) {

        # Mermaid node holiday check
        $nodeElementHolidayCheck = "elementHolidayCheck{During Holiday?}"

    
        $mdHolidayCheck =@"
--> $nodeElementHolidayCheck
$nodeElementHolidayCheck -->|Yes| Holidays
$nodeElementHolidayCheck -->|No| $nodeElementAfterHoursCheck
    $nodeElementAfterHoursCheck -->|Yes| $mdDefaultCallflow
    $nodeElementAfterHoursCheck -->|No| $mdAfterHoursCallFlow

$mdSubGraphHolidays

"@

    }

    # Check if auto attendant has holidays but no after hours
    else {

        $mdHolidayCheck =@"
--> $nodeElementHolidayCheck
$nodeElementHolidayCheck -->|Yes| Holidays
$nodeElementHolidayCheck -->|No| nextStep(next step)

$mdSubGraphHolidays

"@

    }
    
}

. Get-VoiceApp
. Find-Holidays -VoiceAppId $VoiceApp.Identity
. Find-BusinessHours -VoiceAppId $VoiceApp.Identity

if ($aaHasHolidays -eq $true) {
    . Get-AutoAttendantCallFlow
}

else {
    $mdSubGraphHolidays = $null
}

. Set-Mermaid -docType $docType



