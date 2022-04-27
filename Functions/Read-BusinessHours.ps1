<#
    .SYNOPSIS
    Reads Business Hours from an auto Attendant in a human readable format.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function Read-BusinessHours {
    param (
    )

    $weekDays = @(

        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday"
    )

    $aaEffectiveScheduleProperties = New-Object -TypeName psobject

    foreach ($day in $weekDays) {

        [string]$currentDayHoursFriendly = ""

        $currentDayHoursEntryCount = $aaAfterHoursScheduleProperties."$($day)Hours".Count

        foreach ($entry in $aaAfterHoursScheduleProperties."$($day)Hours") {

            $start = $entry.Start
            $end = $entry.End

            if ($currentDayHoursEntryCount -le 1) {
                $comma = $null
            }

            else {
                $comma = ", "
            }

            $currentDayHoursFriendly += "$start-$end$comma"

            $currentDayHoursEntryCount --

        }

        $aaEffectiveScheduleProperties | Add-Member -MemberType NoteProperty -Name "Display$($day)Hours" -Value $currentDayHoursFriendly

    }
    
}