<#
    .SYNOPSIS
    Creates a new PSObject containing information about auto attendants which have holiday call flows.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function New-HolidayLinkProperties {
    param (
        [Parameter(Mandatory=$false)][String]$HolidayLinkVoiceAppType,
        [Parameter(Mandatory=$false)][String]$HolidayLinkVoiceAppName,
        [Parameter(Mandatory=$false)][String]$HolidayLinkVoiceAppId,
        [Parameter(Mandatory = $false)][String]$HolidayLinkStartDate,
        [Parameter(Mandatory = $false)][String]$HolidayLinkEndDate,
        [Parameter(Mandatory = $false)][String]$HolidayLinkTimeZone,
        [Parameter(Mandatory = $false)][String]$HolidayLinkScheduleName,
        [Parameter(Mandatory = $false)][String]$HolidayLinkCallFlowName

    )

    $holidayLinkProperties = New-Object -TypeName psobject

    $holidayLinkProperties | Add-Member -MemberType NoteProperty -Name "VoiceAppName" -Value $HolidayLinkVoiceAppName
    $holidayLinkProperties | Add-Member -MemberType NoteProperty -Name "VoiceAppType" -Value $HolidayLinkVoiceAppType
    $holidayLinkProperties | Add-Member -MemberType NoteProperty -Name "VoiceAppId" -Value $HolidayLinkVoiceAppId
    $holidayLinkProperties | Add-Member -MemberType NoteProperty -Name "StartDate" -Value $HolidayLinkStartDate
    $holidayLinkProperties | Add-Member -MemberType NoteProperty -Name "EndDate" -Value $HolidayLinkEndDate
    $holidayLinkProperties | Add-Member -MemberType NoteProperty -Name "TimeZone" -Value $HolidayLinkTimeZone
    $holidayLinkProperties | Add-Member -MemberType NoteProperty -Name "HolidayScheduleName" -Value $HolidayLinkScheduleName
    $holidayLinkProperties | Add-Member -MemberType NoteProperty -Name "HolidayCallFlowName" -Value $HolidayLinkCallFlowName

    $holidayLinkedVoiceApps += $holidayLinkProperties

}