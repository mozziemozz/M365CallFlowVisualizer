<#
    .SYNOPSIS
    Used the M365 Call Flow Visualizer Script to create a seperate Markdown file for each top level voice app and creates a Markdown file, an htm file and exported audio file and TTS greetings in a subfolder of the voice app Id.

    .DESCRIPTION
    This script can only be used when it's called from another script in another repo. Do not use it manually.

    Author:             Martin Heusser
    Version:            1.0.6

#>

#Requires -Modules @{ ModuleName = "MicrosoftTeams"; ModuleVersion = "5.0.0" }, "Microsoft.Graph.Users", "Microsoft.Graph.Groups"

[CmdletBinding(DefaultParametersetName="None")]
param(
    [Parameter(Mandatory=$true)][String]$ArticlesRelativePath,
    [Parameter(Mandatory=$false)][Switch]$IncludeCallFlowDocs
)

# Import external functions via dot sourcing
. .\Functions\Connect-M365CFV.ps1
. .\Functions\Get-AllVoiceAppsAndResourceAccounts.ps1

. Connect-M365CFV

$VoiceApps = @()

# Get all voice apps and resource accounts from external function
. Get-AllVoiceAppsAndResourceAccounts

foreach ($AutoAttendant in $AllAutoAttendants) {

    $ResourceAccounts = $AutoAttendant.ApplicationInstances

    foreach ($ResourceAccount in $ResourceAccounts) {

        $ResourceAccountCheck = Get-CsOnlineApplicationInstance -Identity $ResourceAccount

        if ($ResourceAccountCheck.PhoneNumber) {

            $VoiceApps += $AutoAttendant.Identity

        }

    }
    
}

foreach ($CallQueue in $AllCallQueues) {

    $ResourceAccounts = $CallQueue.ApplicationInstances

    foreach ($ResourceAccount in $ResourceAccounts) {

        $ResourceAccountCheck = Get-CsOnlineApplicationInstance -Identity $ResourceAccount

        if ($ResourceAccountCheck.PhoneNumber) {

            $VoiceApps += $CallQueue.Identity

        }

    }
    
}

Set-Content -Path "$localRepoPath\Docs\articles\$ArticlesRelativePath\call-flows.md" -Value "# Call Flows"

if ($IncludeCallFlowDocs) {

    Add-Content -Path "$localRepoPath\Docs\articles\$ArticlesRelativePath\call-flows.md" -Value "`n[!INCLUDE[](call_flow_docs.md)]`n"

}

$VoiceApps = $VoiceApps | Sort-Object -Unique

foreach ($VoiceAppIdentity in $VoiceApps) {

    . .\M365CallFlowVisualizerV2.ps1 -Identity $VoiceAppIdentity -Theme dark -CustomFilePath "$localRepoPath\Docs\articles\$ArticlesRelativePath\$voiceAppIdentity" -ShowCqAgentPhoneNumbers -ExportAudioFiles -ExportTTSGreetings -ShowAudioFileName -ShowTTSGreetingText -ExportPng $true -CacheResults $true -ExportHtml $true -DocFxMode -ShowCqAgentOptInStatus -ShowSharedVoicemailGroupMembers $true

    $markdownInclude = "&nbsp;`n[!include[$($VoiceAppFileName)]($voiceAppIdentity/$(($VoiceAppFileName).Replace(" ","_"))_CallFlow$fileExtension)]`n- [Enlarge View]($voiceAppIdentity/$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.htm)`n- [PNG Download]($voiceAppIdentity/$(($VoiceAppFileName).Replace(" ","_"))_CallFlow.png) `n&nbsp;`n"

    Add-Content -Path "$localRepoPath\Docs\articles\$ArticlesRelativePath\call-flows.md" -Value $markdownInclude

}