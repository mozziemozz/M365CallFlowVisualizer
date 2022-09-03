<#
    .SYNOPSIS
    Creates a new PSObject containing the User Id, Name, Voice App Id, Name, Type and Call Flow Action.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function New-VoiceAppUserLinkProperties {
    param (
        [Parameter(Mandatory=$false)][String]$userLinkUserId,
        [Parameter(Mandatory=$false)][String]$userLinkUserName,
        [Parameter(Mandatory=$false)][String]$userLinkVoiceAppType,
        [Parameter(Mandatory=$false)][String]$userLinkVoiceAppActionType,
        [Parameter(Mandatory=$false)][String]$userLinkVoiceAppName,
        [Parameter(Mandatory=$false)][String]$userLinkVoiceAppId
    )

    $userLinkProperties = New-Object -TypeName psobject

    $userLinkProperties | Add-Member -MemberType NoteProperty -Name "UserId" -Value $userLinkUserId
    $userLinkProperties | Add-Member -MemberType NoteProperty -Name "UserName" -Value $userLinkUserName
    $userLinkProperties | Add-Member -MemberType NoteProperty -Name "VoiceAppName" -Value $userLinkVoiceAppName
    $userLinkProperties | Add-Member -MemberType NoteProperty -Name "VoiceAppType" -Value $userLinkVoiceAppType
    $userLinkProperties | Add-Member -MemberType NoteProperty -Name "VoiceAppActionType" -Value $userLinkVoiceAppActionType
    $userLinkProperties | Add-Member -MemberType NoteProperty -Name "VoiceAppId" -Value $userLinkVoiceAppId

    $userLinkVoiceApps += $userLinkProperties

}