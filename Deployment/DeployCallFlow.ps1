<#
    .SYNOPSIS
    Deploys simple call flows from A-Z by data specified in a CSV file.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

#Requires -Modules @{ ModuleName = "MicrosoftTeams"; ModuleVersion = "4.8.0" }, "Microsoft.Graph.Users", "Microsoft.Graph.Groups", "Microsoft.Graph.Identity.DirectoryManagement", "ExchangeOnlineManagement"

. .\Functions\Connect-M365CFV.ps1

. Connect-M365CFV
. Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Domain.ReadWrite.All","Organization.ReadWrite.All","Directory.ReadWrite.All" -TenantId $msTeamsTenantId

. Connect-ExchangeOnline

$supportedTimeZones = Get-CsAutoAttendantSupportedTimeZone
$supportedLanguages = Get-CsAutoAttendantSupportedLanguage

function Deploy-CallFlow {
    [CmdletBinding(DefaultParametersetName="None")]
    param(
        [Parameter(Mandatory=$true)][String]$departmentName,
        [Parameter(Mandatory=$false)][String]$topLevelNumber,
        [Parameter(Mandatory=$false)][ValidateSet("CallingPlan","OperatorConnect","DirectRouting",$null)][String]$numberType,
        [Parameter(Mandatory=$false)][String]$voiceRoutingPolicyName,
        [Parameter(Mandatory=$true)][String]$usageLocation,
        [Parameter(Mandatory=$true)][single]$agentAlertTime,
        [Parameter(Mandatory=$true)][single]$timeoutThreshold,
        [Parameter(Mandatory=$true)][String]$promptLanguage,
        [Parameter(Mandatory=$true)][String]$timeoutSharedVoicemailPrompt,
        [Parameter(Mandatory=$false)][String]$afterHoursDisconnectPrompt,
        [Parameter(Mandatory=$true)][String]$timeZone,
        [Parameter(Mandatory=$false)][String]$BusinessHoursStart1,
        [Parameter(Mandatory=$false)][String]$BusinessHoursEnd1,
        [Parameter(Mandatory=$false)][String]$BusinessHoursStart2,
        [Parameter(Mandatory=$false)][String]$BusinessHoursEnd2
    )

    if ($callFlow.DomainSuffix) {

        $tenantDomains = (Get-MgDomain).Id

        if ($tenantDomains -contains $callFlow.DomainSuffix) {

            Write-Warning "The default domain '$defaultDomain' will be replaced with $($callFlow.DomainSuffix)"

            $defaultDomain = $callFlow.DomainSuffix    

        }

        else {

            Write-Error -Message "The domain '$($callFlow.DomainSuffix)' cannot be found in your tenant. Aborting."
            exit

        }

    }

    else {

        $defaultDomain = (Get-MgDomain | Where-Object {$_."IsDefault" -eq $true}).Id

    }

    if ($supportedLanguages.Id -notcontains $promptLanguage) {

        Write-Warning "The specified prompt language '$promptLanguage' is invalid."

        $promptLanguage = ($supportedLanguages | Out-GridView -Title "Please choose a valid language from the list" -PassThru).Id

    }

    if ($supportedTimeZones.Id -notcontains $timeZone) {

        Write-Warning "The specified time zone '$timeZone' is invalid."

        $timeZone = ($supportedTimeZones | Out-GridView -Title "Please choose a valid time zone from the list" -PassThru).Id

    }

    $mailNickName = $departmentName.Replace(" ","_")

    Write-Host "Deploying call flow for $departmentName..." -ForegroundColor Magenta

    $resourceAccounts = (@{
        1 = [PSCustomObject]@{Name=$departmentName;Type="AutoAttendant";TypeShort="AA";PhoneNumber=$topLevelNumber;AppId="ce933385-9390-45d1-9512-c8d228074e07";UpnPrefix="ra_aa_"}
        2 = [PSCustomObject]@{Name=$departmentName;Type="CallQueue";TypeShort="CQ";PhoneNumber="";AppId="11cd3e2e-fccb-42ad-ad00-878b93575e07";UpnPrefix="ra_cq_"}    
    }).Values | Sort-Object -Descending

    $channelName = "$departmentName CQ"

    # Start Logging
    $ScriptName = $MyInvocation.MyCommand.Name
    $Scriptversion = "1.4"
    $CurrentDate = $((Get-Date).ToString('MM-dd-yyyy_hh-mm-ss'))
    $ScriptLogPath = ".\$ScriptName-V$Scriptversion-$defaultDomain-$env:USERNAME-$CurrentDate.Log"
    #Start-Transcript -Path $ScriptLogPath -NoClobber

    $MCOVU = (Get-MgSubscribedSku -All | Where-Object {$_.SkuPartNumber -eq "PHONESYSTEM_VIRTUALUSER"}).SkuId

    # Check for alternative license name if none are found
    if (!$MCOVU) {
        $MCOVU = (Get-MgSubscribedSku -All | Where-Object {$_.SkuPartNumber -eq "MCOEV_VIRTUALUSER"}).SkuId
    }

    # Create Team

    if (!(Get-Team -MailNickName $mailNickName)) {

        $newTeam = New-Team -DisplayName "$departmentName $($resourceAccounts[-1].TypeShort)" -Description "Call Queue Team for $departmentName" -MailNickName "cq_$($mailNickName)"  -Visibility Private
        $newChannel = New-TeamChannel -GroupId $newTeam.GroupId -DisplayName $channelName -Description "Channel for calls to $departmentName" -MembershipType Standard
        $newChannelOwner = (Get-TeamChannelUser -GroupId $newTeam.GroupId -DisplayName $channelName | Where-Object {$_.Role -eq "Owner"}[0]).UserId

    }

    else {
        
        $newTeam = Get-Team -MailNickName $mailNickName

        if (Get-TeamChannel -GroupId $newTeam.GroupId | Where-Object {$_.DisplayName -eq $channelName}) {

            $newChannel = Get-TeamChannel -GroupId $newTeam.GroupId | Where-Object {$_.DisplayName -eq $channelName}
            $newChannelOwner = (Get-TeamChannelUser -GroupId $newTeam.GroupId -DisplayName $channelName | Where-Object {$_.Role -eq "Owner"}[0]).UserId

        }

        else {

            $newChannel = New-TeamChannel -GroupId $newTeam.GroupId -DisplayName $channelName -Description "Channel for calls to $departmentName" -MembershipType Standard
            $newChannelOwner = (Get-TeamChannelUser -GroupId $newTeam.GroupId -DisplayName $channelName | Where-Object {$_.Role -eq "Owner"}[0]).UserId    

        }

    }

    Set-UnifiedGroup -Identity $newTeam.GroupId -HiddenFromExchangeClientsEnabled:$false -AutoSubscribeNewMembers:$true

    # Create resource accounts
    foreach ($resourceAccount in $resourceAccounts) {

        $upn = $resourceAccount.UpnPrefix + $resourceAccount.Name.Replace(" ","_") + "@" + $defaultDomain

        # Create new resource account
        Write-Host "Creating resource account for '$upn'..." -ForegroundColor Cyan

        New-CsOnlineApplicationInstance -UserPrincipalName $upn -ApplicationId $resourceAccount.AppId -DisplayName "$($resourceAccount.Name) $($resourceAccount.TypeShort)"

        # Add generated upn and displayname to psobject
        $resourceAccount | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $upn
        $resourceAccount | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "$($resourceAccount.Name) $($resourceAccount.TypeShort)"

        # Wait for resource account to appear in Azure AD
        Start-Sleep 15
        do {

            Write-Host "Checking if the user object is already available in AAD... next try in 15s..." -ForegroundColor Cyan
            Start-Sleep 15
            $resourceAccountObject = Get-MgUser -UserId $upn

        } until(
            $resourceAccountObject.UserPrincipalName -eq $upn
        )

        $resourceAccount | Add-Member -MemberType NoteProperty -Name "ObjectId" -Value $resourceAccountObject.Id

        # License resource account
        Write-Host "Setting usage location for resource account '$upn'..." -ForegroundColor Cyan
        Update-MgUser -UserId $upn -UsageLocation $usageLocation
        Start-Sleep 15
        Write-Host "Assigning license for resource account '$upn'..." -ForegroundColor Cyan
        do {
            $error.Clear()
            Start-Sleep 15
            Set-MgUserLicense -UserId $upn -AddLicenses @(@{SkuId = $MCOVU}) -RemoveLicenses @()
        } until (
            !$error
        )

        # Checking for Teams user
        Start-Sleep 30
        do {
            Write-Host "Checking if the resource account '$upn' already exists in Teams... next try in 30s" -ForegroundColor Cyan
            Start-Sleep 30
            $checkTeamsUser = Get-CsOnlineUser -Identity $upn
        } until (
            $checkTeamsUser
        )

        Write-Host "Resource account '$upn' is ready in Teams. Continuing..." -ForegroundColor Cyan

        if ($voiceRoutingPolicyName) {

            # Assign online voice routing policy
            Write-Host "Assigning voice routing policy for resource account '$upn'..." -ForegroundColor Cyan
            Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $voiceRoutingPolicyName

        }

        # Assign phone number to resource account, if applicable
        if ($resourceAccount.PhoneNumber) {

            do {
                $error.Clear()
                Write-Host "Trying to assign phone number for resource account '$upn'... next try in 30s..." -ForegroundColor Cyan
                Start-Sleep 30
                Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $resourceAccount.PhoneNumber -PhoneNumberType $numberType
            } until (
                !$error
            )

        }


    }

    # Create call queue
    $callQueueResourceAccount = $resourceAccounts | Where-Object {$_.Type -eq "CallQueue"}

    $newCallQueue = New-CsCallQueue -Name $callQueueResourceAccount.DisplayName `
    -LanguageId $promptLanguage `
    -UseDefaultMusicOnHold $true `
    -AgentAlertTime $agentAlertTime `
    -AllowOptOut $true `
    -ConferenceMode $true `
    -RoutingMethod Attendant `
    -PresenceBasedRouting $false `
    -TimeoutThreshold $timeoutThreshold `
    -TimeoutAction SharedVoicemail `
    -TimeoutActionTarget $newTeam.GroupId `
    -TimeoutSharedVoicemailTextToSpeechPrompt $timeoutSharedVoicemailPrompt `
    -EnableTimeoutSharedVoicemailSystemPromptSuppression $true `
    -DistributionLists @($newTeam.GroupId) `
    -ChannelId $newChannel.Id `
    -ChannelUserObjectId $newChannelOwner

    do {
                    
        $verifyCallQueueDeployment = Get-CsCallQueue -WarningAction SilentlyContinue -Identity $newCallQueue.Identity
        Write-Host "Checking if the Call Queue is already deployed... next try in 10s..." -ForegroundColor Cyan
        Start-Sleep 10

    } until ($verifyCallQueueDeployment)

    New-CsOnlineApplicationInstanceAssociation -Identities @($callQueueResourceAccount.ObjectId) -ConfigurationId $newCallQueue.Identity -ConfigurationType CallQueue

    # Create auto attendant

    $autoAttendantResourceAccount = $resourceAccounts | Where-Object {$_.Type -eq "AutoAttendant"}

    $newCallableEntitiy = New-CsAutoAttendantCallableEntity -Identity $callQueueResourceAccount.ObjectId -Type ApplicationEndpoint
    $defaultMenuOptions = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -CallTarget $newCallableEntitiy -DtmfResponse Automatic
    $defaultMenu = New-CsAutoAttendantMenu -Name "Default menu" -MenuOptions @($defaultMenuOptions)
    $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default call flow" -Menu $defaultMenu

    if ($BusinessHoursStart1 -and $BusinessHoursEnd1) {

        $afterHoursGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $afterHoursDisconnectPrompt
        $automaticMenuOption = New-CsAutoAttendantMenuOption -Action Disconnect -DtmfResponse Automatic
        $afterHoursMenu = New-CsAutoAttendantMenu -Name "After Hours menu" -MenuOptions @($automaticMenuOption)
        $afterHoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours call flow" -Greetings @($afterHoursGreetingPrompt) -Menu $afterHoursMenu

        if (!$BusinessHoursStart2 -and !$BusinessHoursEnd2) {

            $tr1 = New-CsOnlineTimeRange -Start $BusinessHoursStart1 -End $BusinessHoursEnd1

            $afterHoursSchedule = New-CsOnlineSchedule -Name "After Hours" -WeeklyRecurrentSchedule -MondayHours @($tr1) -TuesdayHours @($tr1) -WednesdayHours @($tr1) -ThursdayHours @($tr1) -FridayHours @($tr1) -Complement

        }

        else {

            $tr1 = New-CsOnlineTimeRange -Start $BusinessHoursStart1 -End $BusinessHoursEnd1
            $tr2 = New-CsOnlineTimeRange -Start $BusinessHoursStart2 -End $BusinessHoursEnd2

            $afterHoursSchedule = New-CsOnlineSchedule -Name "After Hours" -WeeklyRecurrentSchedule -MondayHours @($tr1,$tr2) -TuesdayHours @($tr1,$tr2) -WednesdayHours @($tr1,$tr2) -ThursdayHours @($tr1,$tr2) -FridayHours @($tr1,$tr2) -Complement

        }

        $afterHoursCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type AfterHours -ScheduleId $afterHoursSchedule.Id -CallFlowId $afterHoursCallFlow.Id

        $newAutoAttendant = New-CsAutoAttendant -Name $autoAttendantResourceAccount.DisplayName -DefaultCallFlow $defaultCallFlow -Language $promptLanguage -TimeZoneId $timeZone -CallFlows @($afterHoursCallFlow) -CallHandlingAssociations @($afterHoursCallHandlingAssociation)
    
    }

    else {

        $newAutoAttendant = New-CsAutoAttendant -Name $autoAttendantResourceAccount.DisplayName -DefaultCallFlow $defaultCallFlow -Language $promptLanguage -TimeZoneId $timeZone

    }

    do {
        
        $verifyAutoAttendantDeployment = Get-CsAutoAttendant -Identity $newAutoAttendant.Identity
        Write-Host "Checking if the Auto Attendant is already deployed... next try in 10s..." -ForegroundColor Cyan
        Start-Sleep 10

    } until ($verifyAutoAttendantDeployment)

    New-CsOnlineApplicationInstanceAssociation -Identities @($autoAttendantResourceAccount.ObjectId) -ConfigurationId $newAutoAttendant.Identity -ConfigurationType AutoAttendant

    #Stop-Transcript

}

$callFlows = Import-Csv -Path .\Deployment\VoiceAppList.csv -Delimiter ";" -Encoding UTF8

foreach ($callFlow in $callFlows) {

    $defaultDomain = (Get-MgDomain | Where-Object {$_."IsDefault" -eq $true}).Id

    . Deploy-CallFlow `
    -departmentName $callFlow.DepartmentName `
    -toplevelnumber $callFlow.TopLevelNumber `
    -NumberType $callFlow.NumberType`
    -VoiceRoutingPolicyName $callFlow.VoiceRoutingPolicyName `
    -UsageLocation $callFlow.UsageLocation `
    -AgentAlertTime $callFlow.AgentAlertTime `
    -timeoutThreshold $callFlow.TimeoutThreshold `
    -promptlanguage $callFlow.PromptLanguage `
    -timeoutSharedVoicemailPrompt $callFlow.TimeoutSharedVoiceMailPrompt `
    -afterhoursdisconnectprompt $callFlow.AfterHoursDisconnectPrompt `
    -timeZone $callFlow.TimeZone `
    -BusinessHoursStart1 $callFlow.BusinessHoursStart1 `
    -BusinessHoursEnd1 $callFlow.BusinessHoursEnd1 `
    -BusinessHoursStart2 $callFlow.BusinessHoursStart2 `
    -BusinessHoursEnd2 $callFlow.BusinessHoursEnd2 `

    . .\M365CallFlowVisualizerV2.ps1 -Identity $newAutoAttendant.Identity -PreviewHtml -ShowTTSGreetingText -TruncateGreetings 80

}