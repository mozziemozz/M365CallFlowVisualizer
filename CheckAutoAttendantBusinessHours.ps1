$localTimeZone = (Get-TimeZone).Id

$holidayLinkedVoiceApps = @()
$businessHoursLinkedVoiceApps = @()

# . .\M365CallFlowVisualizerV2.ps1 -CheckCallFlowRouting -SaveToFile $false -CacheResults $false -Identity "edde8e13-1a73-434b-a724-6b6d805d3cf7"
. .\M365CallFlowVisualizerV2.ps1 -CheckCallFlowRouting -SaveToFile $false -CacheResults $false -Identity "9516a748-95e5-4024-aae1-5f11fad27a52"

$holidayLinkedVoiceApps

$holidayException = $false

foreach ($linkedHoliday in $holidayLinkedVoiceApps) {

    # Local date time
    $localDateTime = Get-Date #-Date "03.11.2023 07:59:59"

    # Time zone configured on Auto Attendant
    $toTimeZone = $linkedHoliday.TimeZone

    # Convert strings from Auto Attendant holiday to date time object
    $linkedHolidayStart = Get-Date $linkedHoliday.StartDate
    $linkedHolidayEnd = Get-Date $linkedHoliday.EndDate

    # Convert Auto Attendant holiday to time zone configured on Auto Attendant
    # $autoAttendantDateTimeStart = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($linkedHolidayStart, $localTimeZone, $toTimeZone)
    # $autoAttendantDateTimeEnd = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($linkedHolidayEnd, $localTimeZone, $toTimeZone)

    # Convert local time to time zone configured on Auto Attendant
    $convertedDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($localDateTime, $localTimeZone, $toTimeZone)

    Write-Host "Local Time Zone: $localTimeZone" -ForegroundColor Yellow
    Write-Host "Local Date Time: $localDateTime" -ForegroundColor Yellow
    Write-Host "Auto Attendant Time Zone: $toTimeZone" -ForegroundColor Yellow
    Write-Host "Time in Auto Attendant Time Zone: $convertedDateTime" -ForegroundColor Yellow

    if ($VoiceAppFileName -eq $aa.Name) {

        $consoleOutput = "The top-level Auto Attendant '$($linkedHoliday.VoiceAppName)' is closed because of a holiday."

    }

    else {

        $consoleOutput = "The Auto Attendant '$($linkedHoliday.VoiceAppName)' which is a nested in the call flow of the top-level Auto Attendant '$VoiceAppFileName' is closed because of a holiday."

    }

    if ($convertedDateTime -ge $linkedHolidayStart -and $localDateTime -le $linkedHolidayEnd) {

        Write-Host "$consoleOutput`nHoliday Call Flow Name: $($linkedHoliday.HolidayCallFlowName)`nHoliday Schedule Name: $($linkedHoliday.HolidayScheduleName)" -ForegroundColor Red

        $holidayException = $true
        exit

    }

    else {

        Write-Host "It's not a holiday right now! $($linkedHoliday.VoiceAppName) is open." -ForegroundColor Green

    }

}

$localDayOfWeek = $convertedDateTime.DayOfWeek

$businessHoursLinkedVoiceApps