#Requires -Modules @{ ModuleName = "MicrosoftTeams"; ModuleVersion = "4.8.0" }, "Microsoft.Graph.Users", "Microsoft.Graph.Groups", "Microsoft.Graph.Identity.DirectoryManagement", "ExchangeOnlineManagement"

[CmdletBinding(DefaultParametersetName="None")]
param(
    [Parameter(Mandatory=$true)][String]$tenantName,
    [Parameter(Mandatory=$true)][String]$departmentName,
    [Parameter(Mandatory=$false)][String]$topLevelNumber,
    [Parameter(Mandatory=$false)][ValidateSet("CallingPlan","OperatorConnect","DirectRouting")][String]$numberType,
    [Parameter(Mandatory=$false)][String]$voiceRoutingPolicyName,
    [Parameter(Mandatory=$true)][String]$usageLocation,
    [Parameter(Mandatory=$false)][single]$agentAlertTime = 30,
    [Parameter(Mandatory=$false)][single]$timeoutThreshold = 30,
    [Parameter(Mandatory=$false)][String]$promptLanguage = "en-GB",
    [Parameter(Mandatory=$true)][String]$timeoutSharedVoicemailPrompt,
    [Parameter(Mandatory=$true)][String]$afterHoursDisconnectPrompt,
    [Parameter(Mandatory=$false)][String]$timeZone = "W. Europe Standard Time"

)

. .\Functions\Connect-M365CFV.ps1

. Connect-M365CFV
. Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Domain.ReadWrite.All","Organization.ReadWrite.All","Directory.ReadWrite.All" -TenantId $msTeamsTenantId

. Connect-ExchangeOnline

# Auto Attendant Details
$tr1 = New-CsOnlineTimeRange -Start 08:15 -End 12:00
$tr2 = New-CsOnlineTimeRange -Start 13:30 -End 17:45

$resourceAccounts = (@{
    1 = [PSCustomObject]@{Name=$departmentName;Type="AutoAttendant";TypeShort="AA";PhoneNumber=$topLevelNumber;AppId="ce933385-9390-45d1-9512-c8d228074e07";UpnPrefix="ra_aa_";TimeOutTargetNumber=""}
    2 = [PSCustomObject]@{Name=$departmentName;Type="CallQueue";TypeShort="CQ";PhoneNumber="";AppId="11cd3e2e-fccb-42ad-ad00-878b93575e07";UpnPrefix="ra_cq_";TimeOutTargetNumber=$timeOutTargetNumber}    
}).Values | Sort-Object -Descending

$channelName = "$departmentName CQ"

# Start Logging
$ScriptName = $MyInvocation.MyCommand.Name
$Scriptversion = "1.4"
$CurrentDate = $((Get-Date).ToString('MM-dd-yyyy_hh-mm-ss'))
$ScriptLogPath = ".\$ScriptName-V$Scriptversion-$TenantName-$env:USERNAME-$CurrentDate.Log"
#Start-Transcript -Path $ScriptLogPath -NoClobber

$defaultDomain = (Get-MgDomain | Where-Object {$_."IsDefault" -eq $true}).Id

$MCOVU = (Get-MgSubscribedSku -All | Where-Object {$_.SkuPartNumber -eq "PHONESYSTEM_VIRTUALUSER"}).SkuId

# Check for alternative license name if none are found
if (!$MCOVU) {
    $MCOVU = (Get-MgSubscribedSku -All | Where-Object {$_.SkuPartNumber -eq "MCOEV_VIRTUALUSER"}).SkuId
}

# Create Team

$newTeam = New-Team -DisplayName "$departmentName $($resourceAccounts[-1].TypeShort)" -Description "Call Queue Team for $departmentName" -MailNickName "cq_$($mailNickName)"  -Visibility Private
$newChannel = New-TeamChannel -GroupId $newTeam.GroupId -DisplayName $channelName -Description "Channel for calls to $departmentName" -MembershipType Standard
$newChannelOwner = (Get-TeamChannelUser -GroupId $newTeam.GroupId -DisplayName $channelName | Where-Object {$_.Role -eq "Owner"}[0]).UserId
Set-UnifiedGroup -Identity $newTeam.GroupId -HiddenFromExchangeClientsEnabled:$false -AutoSubscribeNewMembers:$true

# Create resource accounts
foreach ($resourceAccount in $resourceAccounts) {

    $upn = $resourceAccount.UpnPrefix + $resourceAccount.Name.Replace(" ","_") + "@" + $defaultDomain
    $mailNickName = $departmentName.Replace(" ","_")

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
            Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $resourceAccount.PhoneNumber -PhoneNumberType "DirectRouting"
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

$afterHoursGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $afterHoursDisconnectPrompt
$automaticMenuOption = New-CsAutoAttendantMenuOption -Action Disconnect -DtmfResponse Automatic
$afterHoursMenu = New-CsAutoAttendantMenu -Name "After Hours menu" -MenuOptions @($automaticMenuOption)
$afterHoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours call flow" -Greetings @($afterHoursGreetingPrompt) -Menu $afterHoursMenu

$afterHoursSchedule = New-CsOnlineSchedule -Name "After Hours" -WeeklyRecurrentSchedule -MondayHours @($tr1,$tr2) -TuesdayHours @($tr1,$tr2) -WednesdayHours @($tr1,$tr2) -ThursdayHours @($tr1,$tr2) -FridayHours @($tr1,$tr2) -Complement

$afterHoursCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type AfterHours -ScheduleId $afterHoursSchedule.Id -CallFlowId $afterHoursCallFlow.Id

$newAutoAttendant = New-CsAutoAttendant -Name $autoAttendantResourceAccount.DisplayName -DefaultCallFlow $defaultCallFlow -Language $promptLanguage -TimeZoneId $timeZone -CallFlows @($afterHoursCallFlow) -CallHandlingAssociations @($afterHoursCallHandlingAssociation)

do {
    
    $verifyAutoAttendantDeployment = Get-CsAutoAttendant -Identity $newAutoAttendant.Identity
    Write-Host "Checking if the Auto Attendant is already deployed... next try in 10s..." -ForegroundColor Cyan
    Start-Sleep 10

} until ($verifyAutoAttendantDeployment)

New-CsOnlineApplicationInstanceAssociation -Identities @($autoAttendantResourceAccount.ObjectId) -ConfigurationId $newAutoAttendant.Identity -ConfigurationType AutoAttendant

#Stop-Transcript