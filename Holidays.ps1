switch ($DateFormat) {
    EU {
        $dateFormat = "dd.MM.yyyy HH:mm"
    }
    US {
        $dateFormat = "MM.dd.yyyy HH:mm"
    }
    Default {
        $dateFormat = "dd.MM.yyyy HH:mm"
    }
}

$holidayScheduleSorted = $holidaySchedule.FixedSchedule.DateTimeRanges | Sort-Object Start

$holidayDates = ""
$holidayScheduleCounter = 0

foreach ($holidayDate in $holidayScheduleSorted) {

    $holidayDates += "$($holidayDate.Start.ToString($dateFormat))<br>$($holidayDate.End.ToString($dateFormat))"

    $holidayScheduleCounter ++

    if ($holidayScheduleCounter -lt $holidayScheduleSorted.Count) {

        $holidayDates += "<br><br>"

    }

}

$nodeElementHolidayDetails =@"

subgraph subgraph$($HolidayCallHandling.CallFlowId)[$mermaidFriendlyHolidayName]
direction LR
elementAAHoliday$($aaObjectId)-$($HolidayCounter)(Dates<br>$holidayDates) --> elementAAHolidayGreeting$($aaObjectId)-$($HolidayCounter)>$holidayGreeting] --> $holidayVoicemailSystemGreeting $nodeElementHolidayAction
    end
"@

