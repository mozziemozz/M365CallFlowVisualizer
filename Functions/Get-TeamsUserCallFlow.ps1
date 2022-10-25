<#
    .SYNOPSIS
    Reads the user calling settings of a Teams user and visualizes the call flow using mermaid-js.

    .DESCRIPTION
    Reads the user calling settings of a Teams user and outputs them in an easy to understand SVG diagram. See the script "ExportAllUserCallFlowsToSVG.ps1" in the root of this repo for an example on how to generate a diagram for each user in a tenant.

    Author:             Martin Heusser
    Version:            1.0.4
    Changelog:          Moved to repository at .\Changelog.md

    .PARAMETER Name
    -UserID
        Specifies the identity of the user by an AAD Object Id
        Required:           false
        Type:               string
        Accepted values:    any string
        Default value:      none

    -UserPrincipalName
        Specifies the identity of the user by a upn
        Required:           false
        Type:               string
        Accepted values:    any string
        Default value:      none

    -SetClipBoard
        Specifies if the mermaid code should be copied to the clipboard after the function has been executed.
        Required:           false
        Type:               boolean
        Default value:      false

    -CustomFilePath
        Specifies the file path for the output file.
        Required:           false
        Type:               string
        Accepted values:    file paths e.g. "C:\Temp"
        Default value:      ".\Output\UserCallingSettings"

    -StandAlone
        Specifies if the function is running in standalone mode. This means that the first node (user object ID) will be drawn including the users name. The value $false will only be used when this function is implemented to M365CallFlowVisualizerV2.ps1
        Required:           false
        Type:               bool
        Accepted values:    true, false
        Default value:      true

    -ExportSvg
        Specifies if the function should export the diagram as an SVG image leveraging the mermaid.ink service
        Required:           false
        Type:               bool
        Accepted values:    true, false
        Default value:      true

    -PreviewSvg
        Specifies if the generated diagram should open mermaid.ink in the default browser
        Required:           false
        Type:               bool
        Accepted values:    true, false
        Default value:      true

    -ObfuscatePhoneNumbers
        Specifies if phone numbers in call flows should be obfuscated for sharing / example reasons. This will replace the last 4 digits in numbers with an asterisk (*) character. Warning: This will only obfuscate phone numbers in node descriptions. The complete phone number will still be included in Markdown, Mermaid and HTML output!
        Required:           false
        Type:               bool
        Default value:      false

    -ExportSvg
        Specifies if the function should export the diagram as a markdown file (*.md) NOT YET IMPLEMENTED
        Required:           false
        Type:               bool
        Accepted values:    true, false
        Default value:      false

    .INPUTS
        None.

    .OUTPUTS
        Files:
            - *.svg
            
    .EXAMPLE
        .\Functions\Get-TeamsUserCallFlow.ps1 -UserPrincipalName user@domain.com

    .LINK
    https://github.com/mozziemozz/M365CallFlowVisualizer
    
#>

function Get-TeamsUserCallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$UserId,
        [Parameter(Mandatory=$false)][String]$UserPrincipalName,
        [Parameter(Mandatory=$false)][string]$CustomFilePath = ".\Output\UserCallingSettings",
        [Parameter(Mandatory=$false)][bool]$StandAlone = $true,
        [Parameter(Mandatory=$false)][bool]$ExportMarkdown = $false,
        [Parameter(Mandatory=$false)][bool]$PreviewSvg = $true,
        [Parameter(Mandatory=$false)][bool]$SetClipBoard = $true,
        [Parameter(Mandatory=$false)][Bool]$ObfuscatePhoneNumbers = $false,
        [Parameter(Mandatory=$false)][bool]$ExportSvg = $true
    )

    . .\Functions\Connect-M365CFV.ps1

    . Connect-M365CFV

    if ($CustomFilePath) {

        $filePath = $CustomFilePath

    }

    else {

        $filePath = ".\"

    }

    if ($UserPrincipalName) {

        $UserId = (Get-CsOnlineUser -Identity $UserPrincipalName).Identity
        $UserId
    }

    $teamsUser = Get-CsOnlineUser -Identity $UserId

    Write-Host "Reading user calling settings for: $($teamsUser.DisplayName)" -ForegroundColor Magenta
    Write-Host "User Id: $UserId" -ForegroundColor Magenta

    $userCallingSettings = Get-CsUserCallingSettings -Identity $UserId

    #$userCallingSettings

    [int]$userUnansweredTimeoutMinutes = ($userCallingSettings.UnansweredDelay).Split(":")[1]
    [int]$userUnansweredTimeoutSeconds = ($userCallingSettings.UnansweredDelay).Split(":")[-1]

    if ($StandAlone) {

        $mdFlowChart = "flowchart TB`n"

        $userNode = "$UserId(User<br> $($teamsUser.DisplayName))"

    }

    else {

        $mdFlowChart = ""

        $userNode = $UserId

    }


    if ($userUnansweredTimeoutMinutes -eq 1) {

        $userUnansweredTimeout = "60 Seconds"

    }

    else {

        $userUnansweredTimeout = "$userUnansweredTimeoutSeconds Seconds"

    }


    # user is neither forwarding or unanswered enabled
    if (!$userCallingSettings.IsForwardingEnabled -and !$userCallingSettings.IsUnansweredEnabled) {

        #Write-Host "User is neither forwaring or unanswered enabled"

        $mdUserCallingSettings = @"
        $userNode
"@

    }

    # user is immediate forwarding enabled
    elseif ($userCallingSettings.ForwardingType -eq "Immediate") {

        #Write-Host "User is immediate forwarding enabled."

        switch ($userCallingSettings.ForwardingTargetType) {
            MyDelegates {

                $mdUserCallingSettings = @"
$userNode --> userForwarding$UserId(Immediate Forwarding)
subgraph subgraphSettings$UserId[ ]
userForwarding$UserId --> subgraphDelegates$UserId

"@

$allMermaidNodes += @("userForwarding$UserId","userForwarding$UserId","")
$allSubgraphs += "subgraphSettings$UserId"

                $mdSubgraphDelegates = @"
subgraph subgraphDelegates$UserId[Delegates of $($teamsUser.DisplayName)]
direction LR
ringType$UserId[(Simultaneous Ring)]

"@

$allMermaidNodes += "ringType$UserId"
$allSubgraphs += "subgraphDelegates$UserId"

                $delegateCounter = 1

                foreach ($delegate in $userCallingSettings.Delegates) {

                    $delegateUserObject = (Get-CsOnlineUser -Identity $delegate.Id)

                    $delegateRing = "                ringType$UserId -.-> delegate$($delegateUserObject.Identity)$delegateCounter($($delegateUserObject.DisplayName))`n"

                    $mdSubgraphDelegates += $delegateRing

                    $allMermaidNodes += "delegate$($delegateUserObject.Identity)$delegateCounter"

                    $delegateCounter ++

                }

                $mdUserCallingSettings += $mdSubgraphDelegates

                switch ($userCallingSettings.UnansweredTargetType) {
                    Voicemail {
                        $mdUnansweredTarget = "--> userVoicemail$UserId(Voicemail<br> $($teamsUser.DisplayName))"
                        $subgraphUnansweredSettings = $null

                        $allMermaidNodes += "userVoicemail$UserId"
                    }
                    Group {

                        switch ($userCallingSettings.CallGroupOrder) {
                            InOrder {
                                $ringOrder = "Serial"
                            }
                            Simultaneous {
                                $ringOrder = "Simultaneous"
                            }
                            Default {}
                        }

                        $subgraphUnansweredSettings = @"
subgraph subgraphCallGroups$UserId[Call Group of $($teamsUser.DisplayName)]
direction LR
callGroupRingType$UserId[($ringOrder Ring)]

"@

$allMermaidNodes += "callGroupRingType$UserId"
$allSubgraphs += "subgraphCallGroups$UserId"

                        $callGroupMemberCounter = 1

                        foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {

                            $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)

                            if ($ringOrder -eq "Serial") {

                                $linkNumber = " |$callGroupMemberCounter|"

                            }

                            else {

                                $linkNumber = $null

                            }

                            $callGroupRing = "callGroupRingType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"

                            $subgraphUnansweredSettings += $callGroupRing

                            $allMermaidNodes += "callGroupMember$($callGroupUserObject.Identity)"

                            $callGroupMemberCounter ++
                            
                        }

                        $subgraphUnansweredSettings += "`nend"

                        $mdUnansweredTarget = "--> subgraphCallGroups$UserId"

    
                    }
                    SingleTarget {

                        if ($userCallingSettings.UnansweredTarget -match "sip:" -or $userCallingSettings.UnansweredTarget -notmatch "\+") {

                            $userForwardingTarget = (Get-CsOnlineUser -Identity $userCallingSettings.UnansweredTarget).DisplayName
                    
                            if (Get-CsOnlineUser -AccountType ResourceAccount |  Where-Object {$_.SipAddress -eq $userCallingSettings.UnansweredTarget}) {
                    
                                $checkUserAccountType = Get-CsOnlineApplicationInstance -Identity $($userCallingSettings.UnansweredTarget).Replace("sip:","")
                                
                            }
                    
                            else {
                    
                                $checkUserAccountType = $null
                    
                            }
                    
                            if ($checkUserAccountType) {
                    
                                switch ($checkUserAccountType.ApplicationId) {
                                    # Call Queue
                                    11cd3e2e-fccb-42ad-ad00-878b93575e07 {
                                        $forwardingTargetType = "Call Queue"
                                    }
                                    # Auto Attendant
                                    ce933385-9390-45d1-9512-c8d228074e07 {
                                        $forwardingTargetType = "Auto Attendant"
                                    }
                                    Default {}
                                }
                    
                                if ($StandAlone -eq $false) {
                    
                                    switch ($forwardingTargetType) {
                                        "Auto Attendant" {
                    
                                            $unansweredUserTargetVoiceAppId = ($allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity
                    
                                        }
                                        "Call Queue" {
                    
                                            $unansweredUserTargetVoiceAppId = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity
                    
                                        }
                                        Default {}
                                    }
                    
                                    $mdUnansweredTarget = "--> $unansweredUserTargetVoiceAppId([$forwardingTargetType<br> $userForwardingTarget])"
                    
                                    if ($nestedVoiceApps -notcontains $unansweredUserTargetVoiceAppId) {
                    
                                        $nestedVoiceApps += $unansweredUserTargetVoiceAppId
                    
                                    }
                    
                                }
                    
                                else {
                    
                                    $mdUnansweredTarget = "--> userUnansweredTarget$unansweredUserTargetVoiceAppId([$forwardingTargetType<br> $userForwardingTarget])"
                    
                                }
                    
                    
                            }
                    
                            else {
                                
                                $forwardingTargetType = "User"
                    
                                if ($null -eq $userForwardingTarget) {
                    
                                    $userForwardingTarget = "External Tenant"
                                    $forwardingTargetType = "Federated User"
                    
                                }
                    
                                if ($StandAlone -eq $false -and $forwardingTargetType -eq "User") {
                    
                                    $unansweredUserTargetUserId = (Get-CsOnlineUser -Identity $userCallingSettings.UnansweredTarget).Identity
                    
                                    $mdUnansweredTarget = "-->$unansweredUserTargetUserId($forwardingTargetType<br> $userForwardingTarget)"

                                    $allMermaidNodes += $unansweredUserTargetUserId
                    
                                    if ($nestedVoiceApps -notcontains $unansweredUserTargetUserId) {
                    
                                        $nestedVoiceApps += $unansweredUserTargetUserId
                    
                                    }
                    
                                }
                    
                                else {
                    
                                    $mdUnansweredTarget = "--> userUnansweredTarget$UserId($forwardingTargetType<br> $userForwardingTarget)"
                    
                                }
                    
                    
                            }
                    
                        }
                    
                        else {
                    
                            $userForwardingTarget = $userCallingSettings.UnansweredTarget
                            $forwardingTargetType = "External Number"

                            if ($ObfuscatePhoneNumbers -eq $true) {

                                $userForwardingTarget = $userForwardingTarget.Remove(($userForwardingTarget.Length -4)) + "****"
            
                            }            
                    
                            $mdUnansweredTarget = "--> userUnansweredTarget$UserId($forwardingTargetType<br> $userForwardingTarget)"
                    
                        }
                    
                        $subgraphUnansweredSettings = $null
                        
                        $allMermaidNodes += "userUnansweredTarget$UserId"
                    
                    }
                    Default {}
                }

                $mdUserCallingSettingsAddition = @"
end
userForwardingResult$UserId --> |No| userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)]
subgraphDelegates$UserId --> userForwardingResult$UserId{Call Answered?}
$subgraphUnansweredSettings
end
userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)] $mdUnansweredTarget
userForwardingResult$UserId --> |Yes| userForwardingConnected$UserId((Call Connected))

"@

                $allMermaidNodes += @("userForwardingResult$UserId","userForwardingTimeout$UserId","userForwardingConnected$UserId")

                $mdUserCallingSettings += $mdUserCallingSettingsAddition

            }
            Voicemail {

                $mdUserCallingSettings = @"
$userNode --> userForwarding$UserId(Immediate Forwarding)
subgraph subgraphSettings$UserId[ ]
userForwarding$UserId --> voicemail$UserId(Voicemail<br> $($teamsUser.DisplayName))
end

"@

$allMermaidNodes += @("userForwarding$UserId","voicemail$UserId")
$allSubgraphs += "subgraphSettings$UserId"

            }
            Group {

                switch ($userCallingSettings.CallGroupOrder) {
                    InOrder {
                        $ringOrder = "Serial"
                    }
                    Simultaneous {
                        $ringOrder = "Simultaneous"
                    }
                    Default {}
                }

                $mdUserCallingSettings = @"
$userNode --> userForwarding$UserId(Immediate Forwarding)
subgraph subgraphSettings$UserId[ ]
userForwarding$UserId --> subgraphCallGroups$UserId

"@

$allMermaidNodes += "userForwarding$UserId"
$allSubgraphs += "subgraphSettings$UserId"

                $mdSubgraphcallGroups = @"
subgraph subgraphCallGroups$UserId[Call Group of $($teamsUser.DisplayName)]
direction LR
ringType$UserId[($ringOrder Ring)]

"@

$allMermaidNodes += "ringType$UserId"
$allSubgraphs += "subgraphCallGroups$UserId"

                $callGroupMemberCounter = 1

                foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {

                    $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)

                    if ($ringOrder -eq "Serial") {

                        $linkNumber = " |$callGroupMemberCounter|"

                    }

                    else {

                        $linkNumber = $null

                    }

                    $callGroupRing = "ringType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"

                    $mdSubgraphcallGroups += $callGroupRing
                    
                    $allMermaidNodes += "callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter"

                    $callGroupMemberCounter ++

                }

                $mdUserCallingSettings += $mdSubgraphcallGroups

                $mdUserCallingSettingsAddition = @"
end
end

"@

                $mdUserCallingSettings += $mdUserCallingSettingsAddition




            }
            SingleTarget {

                if ($userCallingSettings.ForwardingTarget -match "sip:" -or $userCallingSettings.ForwardingTarget -notmatch "\+") {

                    $userForwardingTarget = (Get-CsOnlineUser -Identity $userCallingSettings.ForwardingTarget).DisplayName

                    if (Get-CsOnlineUser -AccountType ResourceAccount |  Where-Object {$_.SipAddress -eq $userCallingSettings.ForwardingTarget}) {

                        $checkUserAccountType = Get-CsOnlineApplicationInstance -Identity $($userCallingSettings.ForwardingTarget).Replace("sip:","")
                        
                    }
            
                    else {
            
                        $checkUserAccountType = $null
            
                    }

                    if ($checkUserAccountType) {

                        switch ($checkUserAccountType.ApplicationId) {
                            # Call Queue
                            11cd3e2e-fccb-42ad-ad00-878b93575e07 {
                                $forwardingTargetType = "Call Queue"
                            }
                            # Auto Attendant
                            ce933385-9390-45d1-9512-c8d228074e07 {
                                $forwardingTargetType = "Auto Attendant"
                            }
                            Default {}
                        }
            
                        if ($StandAlone -eq $false) {

                            switch ($forwardingTargetType) {
                                "Auto Attendant" {
            
                                    $immediateForwardingUserTargetVoiceAppId = ($allAutoAttendants| Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity
            
                                }
                                "Call Queue" {
            
                                    $immediateForwardingUserTargetVoiceAppId = ($allCallQueues| Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity
            
                                }
                                Default {}
                            }

                            $mdImmediateForwardingTarget = "--> $immediateForwardingUserTargetVoiceAppId([$forwardingTargetType<br>$userForwardingTarget])"
            
                            if ($nestedVoiceApps -notcontains $immediateForwardingUserTargetVoiceAppId) {

                                $nestedVoiceApps += $immediateForwardingUserTargetVoiceAppId
            
                            }            

                        }

                        else {

                            $mdImmediateForwardingTarget = "--> userImmediateForwardingTarget$immediateForwardingUserTargetVoiceAppId([$forwardingTargetType<br>$userForwardingTarget])"

                        }


                    }

                    else {

                        $forwardingTargetType = "User"

                        if ($null -eq $userForwardingTarget) {
    
                            $userForwardingTarget = "External Tenant"
                            $forwardingTargetType = "Federated User"
    
                        }

                        if ($StandAlone -eq $false -and $forwardingTargetType -eq "User") {

                            $immediateForwardingUserTargetUserId = (Get-CsOnlineUser -Identity $userCallingSettings.ForwardingTarget).Identity

                            $mdImmediateForwardingTarget = "-->$immediateForwardingUserTargetUserId($forwardingTargetType<br> $userForwardingTarget)"

                            $allMermaidNodes += $immediateForwardingUserTargetUserId
            
                            if ($nestedVoiceApps -notcontains $immediateForwardingUserTargetUserId) {
            
                                $nestedVoiceApps += $immediateForwardingUserTargetUserId
            
                            }
            
                        }

                        else {

                            $mdImmediateForwardingTarget = "--> userImmediateForwardingTarget$UserId($forwardingTargetType<br>$userForwardingTarget)"

                        }


                    }
            

                }

                else {

                    $userForwardingTarget = $userCallingSettings.ForwardingTarget
                    $forwardingTargetType = "External Number"

                    if ($ObfuscatePhoneNumbers -eq $true) {

                        $userForwardingTarget = $userForwardingTarget.Remove(($userForwardingTarget.Length -4)) + "****"
    
                    }            

                    $mdImmediateForwardingTarget = "--> userImmediateForwardingTarget$UserId($forwardingTargetType<br>$userForwardingTarget)"

                    $allMermaidNodes += "userImmediateForwardingTarget$UserId"

                }


                $mdUserCallingSettings = @"
$userNode --> userForwarding$UserId(Immediate Forwarding)
userForwarding$UserId $mdImmediateForwardingTarget

"@

$allMermaidNodes += @("userForwarding$UserId","userForwardingTarget$UserId")
#$allSubgraphs += "subgraphSettings$UserId"

            }
            Default {}
        }

    }

    # user is either forwarding or unanswered enabled
    else {

        # user is forwarding and unanswered enabled
        if ($userCallingSettings.IsForwardingEnabled -and $userCallingSettings.IsUnansweredEnabled) {

            #Write-Host "User is forwaring and unanswered enabled"

            switch ($userCallingSettings.UnansweredTargetType) {
                MyDelegates {

                    $ringOrder = "Simultaneous"

                    $subgraphUnansweredSettings = @"
subgraph subgraphdelegates$UserId[Delegates of $($teamsUser.DisplayName)]
direction LR
delegateRingType$UserId[($ringOrder Ring)]

"@

$allMermaidNodes += "delegateRingType$UserId"
$allSubgraphs += "subgraphdelegates$UserId"

                    $delegateCounter = 1

                    foreach ($delegate in $userCallingSettings.Delegates) {

                        $delegateUserObject = (Get-CsOnlineUser -Identity $delegate.Id)

                        if ($ringOrder -eq "Serial") {

                            $linkNumber = " |$delegateCounter|"

                        }

                        else {

                            $linkNumber = $null

                        }

                        $delegateRing = "delegateRingType$UserId -.->$linkNumber delegateMember$($delegateUserObject.Identity)$delegateCounter($($delegateUserObject.DisplayName))`n"

                        $subgraphUnansweredSettings += $delegateRing

                        $allMermaidNodes += "delegateMember$($delegateUserObject.Identity)$delegateCounter"

                        $delegateCounter ++

                    }

                    $subgraphUnansweredSettings += "`nend"

                    $mdUnansweredTarget = "--> subgraphdelegates$UserId"


                }

                Voicemail {
                    $mdUnansweredTarget = "--> userVoicemail$UserId(Voicemail<br> $($teamsUser.DisplayName))"
                    $subgraphUnansweredSettings = $null

                    $allMermaidNodes += "userVoicemail$UserId"
                }

                Group {

                    if ($userCallingSettings.ForwardingTargetType -eq "Group") {

                        $subgraphUnansweredSettings = $null

                        $mdUnansweredTarget = "--> subgraphCallGroups$UserId"

                    }

                    else {

                        switch ($userCallingSettings.CallGroupOrder) {
                            InOrder {
                                $ringOrder = "Serial"
                            }
                            Simultaneous {
                                $ringOrder = "Simultaneous"
                            }
                            Default {}
                        }
                
                        $subgraphUnansweredSettings = @"
subgraph subgraphCallGroups$UserId[Call Group of $($teamsUser.DisplayName)]
direction LR
callGroupRingType$UserId[($ringOrder Ring)]

"@

$allMermaidNodes += "callGroupRingType$UserId"
$allSubgraphs += "subgraphCallGroups$UserId"
    
                        $callGroupMemberCounter = 1
    
                        foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {
    
                            $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)
    
                            if ($ringOrder -eq "Serial") {
    
                                $linkNumber = " |$callGroupMemberCounter|"
    
                            }
    
                            else {
    
                                $linkNumber = $null
    
                            }
    
                            $callGroupRing = "callGroupRingType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"
    
                            $subgraphUnansweredSettings += $callGroupRing

                            $allMermaidNodes += "callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter"
    
                            $callGroupCounter ++

                        }
    
                        $subgraphUnansweredSettings += "`nend"
    
                        $mdUnansweredTarget = "--> subgraphCallGroups$UserId"
    
    
                    }
    

                }
                SingleTarget {

                    if ($userCallingSettings.UnansweredTarget -match "sip:" -or $userCallingSettings.UnansweredTarget -notmatch "\+") {

                        $userForwardingTarget = (Get-CsOnlineUser -Identity $userCallingSettings.UnansweredTarget).DisplayName

                        if (Get-CsOnlineUser -AccountType ResourceAccount |  Where-Object {$_.SipAddress -eq $userCallingSettings.UnansweredTarget}) {

                            $checkUserAccountType = Get-CsOnlineApplicationInstance -Identity $($userCallingSettings.UnansweredTarget).Replace("sip:","")
                            
                        }

                        else {

                            $checkUserAccountType = $null

                        }

                        if ($checkUserAccountType) {

                            switch ($checkUserAccountType.ApplicationId) {
                                # Call Queue
                                11cd3e2e-fccb-42ad-ad00-878b93575e07 {
                                    $forwardingTargetType = "Call Queue"
                                }
                                # Auto Attendant
                                ce933385-9390-45d1-9512-c8d228074e07 {
                                    $forwardingTargetType = "Auto Attendant"
                                }
                                Default {}
                            }

                            if ($StandAlone -eq $false) {

                                switch ($forwardingTargetType) {
                                    "Auto Attendant" {

                                        $unansweredUserTargetVoiceAppId = ($allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity

                                    }
                                    "Call Queue" {

                                        $unansweredUserTargetVoiceAppId = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity

                                    }
                                    Default {}
                                }

                                $mdUnansweredTarget = "--> $unansweredUserTargetVoiceAppId([$forwardingTargetType<br> $userForwardingTarget])"

                                if ($nestedVoiceApps -notcontains $unansweredUserTargetVoiceAppId) {

                                    $nestedVoiceApps += $unansweredUserTargetVoiceAppId
                
                                }

                            }

                            else {

                                $mdUnansweredTarget = "--> userUnansweredTarget$unansweredUserTargetVoiceAppId([$forwardingTargetType<br> $userForwardingTarget])"

                            }


                        }

                        else {
                            
                            $forwardingTargetType = "User"
        
                            if ($null -eq $userForwardingTarget) {
        
                                $userForwardingTarget = "External Tenant"
                                $forwardingTargetType = "Federated User"
        
                            }

                            if ($StandAlone -eq $false -and $forwardingTargetType -eq "User") {

                                $unansweredUserTargetUserId = (Get-CsOnlineUser -Identity $userCallingSettings.UnansweredTarget).Identity

                                $mdUnansweredTarget = "-->$unansweredUserTargetUserId($forwardingTargetType<br> $userForwardingTarget)"

                                $allMermaidNodes += $unansweredUserTargetUserId

                                if ($nestedVoiceApps -notcontains $unansweredUserTargetUserId) {
    
                                    $nestedVoiceApps += $unansweredUserTargetUserId
                
                                }
    
                            }

                            else {

                                $mdUnansweredTarget = "--> userUnansweredTarget$UserId($forwardingTargetType<br> $userForwardingTarget)"

                            }

    
                        }
    
                    }
    
                    else {
    
                        $userForwardingTarget = $userCallingSettings.UnansweredTarget
                        $forwardingTargetType = "External Number"

                        if ($ObfuscatePhoneNumbers -eq $true) {

                            $userForwardingTarget = $userForwardingTarget.Remove(($userForwardingTarget.Length -4)) + "****"
        
                        }

                        $mdUnansweredTarget = "--> userUnansweredTarget$UserId($forwardingTargetType<br> $userForwardingTarget)"
    
                    }

                    $subgraphUnansweredSettings = $null
                    
                    $allMermaidNodes += "userUnansweredTarget$UserId"

                }
                Default {}
            }


            switch ($userCallingSettings.ForwardingTargetType) {
                MyDelegates {
    
                    $mdUserCallingSettings = @"
$userNode --> userForwarding$UserId(Also Ring)
userForwarding$UserId -.-> userParallelRing$userId(Teams Clients<br> $($teamsUser.DisplayName)) ---> userForwardingResult$UserId
userForwarding$UserId -.-> subgraphDelegates$UserId
subgraph subgraphSettings$UserId[ ]

"@

$allMermaidNodes += @("userForwarding$UserId","userParallelRing$userId","userForwardingResult$UserId")
$allSubgraphs += "subgraphSettings$UserId"
    
                    $mdSubgraphDelegates = @"
subgraph subgraphDelegates$UserId[Delegates of $($teamsUser.DisplayName)]
direction LR
ringType$UserId[(Simultaneous Ring)]

"@

$allMermaidNodes += "ringType$UserId"
$allSubgraphs += "subgraphDelegates$UserId"
    
                    $delegateCounter = 1
    
                    foreach ($delegate in $userCallingSettings.Delegates) {
    
                        $delegateUserObject = (Get-CsOnlineUser -Identity $delegate.Id)
    
                        $delegateRing = "                ringType$UserId -.-> delegate$($delegateUserObject.Identity)$delegateCounter($($delegateUserObject.DisplayName))`n"
    
                        $mdSubgraphDelegates += $delegateRing

                        $allMermaidNodes += "delegate$($delegateUserObject.Identity)$delegateCounter"
    
                        $delegateCounter ++
                    }
    
                    $mdUserCallingSettings += $mdSubgraphDelegates
        
                    $mdUserCallingSettingsAddition = @"
end
userForwardingResult$UserId --> |No| userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)]
subgraphDelegates$UserId --> userForwardingResult$UserId{Call Answered?}
$subgraphUnansweredSettings
end
userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)] $mdUnansweredTarget
userForwardingResult$UserId --> |Yes| userForwardingConnected$UserId((Call Connected))

"@
    
                    $allMermaidNodes += @("userForwardingResult$UserId","userForwardingTimeout$UserId","userForwardingConnected$UserId")

                    $mdUserCallingSettings += $mdUserCallingSettingsAddition
    
                }
                Voicemail {
    
                    $mdUserCallingSettings = @"
$userNode --> userForwarding$UserId(Immediate Forwarding)
subgraph subgraphSettings$UserId[ ]
userForwarding$UserId --> voicemail$UserId(Voicemail<br> $($teamsUser.DisplayName))
end

"@

$allMermaidNodes += @("userForwarding$UserId","voicemail$UserId")
$allSubgraphs += "subgraphSettings$UserId"

                }
                Group {
    
                    switch ($userCallingSettings.CallGroupOrder) {
                        InOrder {
                            $ringOrder = "Serial"
                        }
                        Simultaneous {
                            $ringOrder = "Simultaneous"
                        }
                        Default {}
                    }
    
                    $mdUserCallingSettings = @"
$userNode --> userForwarding$UserId(Also Ring)
userForwarding$UserId -.-> userParallelRing$userId(Teams Clients<br> $($teamsUser.DisplayName)) ---> userForwardingResult$UserId
userForwarding$UserId -.-> subgraphCallGroups$UserId
subgraph subgraphSettings$UserId[ ]

"@

$allMermaidNodes += @("userForwarding$UserId","userParallelRing$userId","userForwardingResult$UserId")
$allSubgraphs += "subgraphSettings$UserId"
    
                    $mdSubgraphcallGroups = @"
subgraph subgraphCallGroups$UserId[Call Group of $($teamsUser.DisplayName)]
direction LR
ringType$UserId[($ringOrder Ring)]

"@

$allMermaidNodes += "ringType$UserId"
$allSubgraphs += "subgraphCallGroups$UserId"

                    $callGroupMemberCounter = 1
    
                    foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {
    
                        $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)
    
                        if ($ringOrder -eq "Serial") {
    
                            $linkNumber = " |$callGroupMemberCounter|"
    
                        }
    
                        else {
    
                            $linkNumber = $null
    
                        }
    
                        $callGroupRing = "ringType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"
    
                        $mdSubgraphcallGroups += $callGroupRing

                        $allMermaidNodes += "callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter"
    
                        $callGroupMemberCounter ++
                    }
    
                    $mdUserCallingSettings += $mdSubgraphcallGroups
    
                    $mdUserCallingSettingsAddition = @"
end
subgraphCallGroups$UserId --> userForwardingResult$UserId{Call Answered?}
$subgraphUnansweredSettings
userForwardingResult$UserId --> |No| userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)]
end
userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)] $mdUnansweredTarget
userForwardingResult$UserId --> |Yes| userForwardingConnected$UserId((Call Connected))

"@
    
                    $allMermaidNodes += @("userForwardingResult$UserId","userForwardingTimeout$UserId","userForwardingConnected$UserId")

                    $mdUserCallingSettings += $mdUserCallingSettingsAddition

                }
                SingleTarget {
    
                    if ($userCallingSettings.ForwardingTarget -match "sip:" -or $userCallingSettings.ForwardingTarget -notmatch "\+") {
    
                        $userForwardingTarget = (Get-CsOnlineUser -Identity $userCallingSettings.ForwardingTarget).DisplayName

                        if (Get-CsOnlineUser -AccountType ResourceAccount |  Where-Object {$_.SipAddress -eq $userCallingSettings.ForwardingTarget}) {

                            $checkUserAccountType = Get-CsOnlineApplicationInstance -Identity $($userCallingSettings.ForwardingTarget).Replace("sip:","")
                            
                        }
                
                        else {
                
                            $checkUserAccountType = $null
                
                        }

                        if ($checkUserAccountType) {

                            switch ($checkUserAccountType.ApplicationId) {
                                # Call Queue
                                11cd3e2e-fccb-42ad-ad00-878b93575e07 {
                                    $forwardingTargetType = "Call Queue"
                                }
                                # Auto Attendant
                                ce933385-9390-45d1-9512-c8d228074e07 {
                                    $forwardingTargetType = "Auto Attendant"
                                }
                                Default {}
                            }
                
                            if ($StandAlone -eq $false) {
                
                                switch ($forwardingTargetType) {
                                    "Auto Attendant" {
                
                                        $alsoRingUserTargetVoiceAppId = ($allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity
                                        $mdAlsoRingWarningMessage = "--> warning$alsoRingUserTargetVoiceAppId[[Warning: Also Ring is enabled<br> $forwardingTargetType will answer automatically!]] --> "
                                    }
                                    "Call Queue" {
                
                                        $alsoRingUserTargetVoiceAppId = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity
                                        $checkAlsoRingTargetCqOverflowThreshold = ($allCallQueues | Where-Object {$_.Identity -eq $alsoRingUserTargetVoiceAppId}).OverflowThreshold

                                        if ($checkAlsoRingTargetCqOverflowThreshold -eq 0) {
                                            $mdAlsoRingWarningMessage = "-.-> "
                                        }

                                        else {
                                            $mdAlsoRingWarningMessage = "--> warning$alsoRingUserTargetVoiceAppId[[Warning: Also Ring is enabled<br> $forwardingTargetType will answer automatically!]] --> "
                                        }
                                    }
                                    Default {}

                                }

                                $allMermaidNodes += "warning$alsoRingUserTargetVoiceAppId"
                
                                #2 variables without whitespace!
                                $mdAlsoRingTarget = "$mdAlsoRingWarningMessage$alsoRingUserTargetVoiceAppId([$forwardingTargetType<br>$userForwardingTarget])"
                
                                if ($nestedVoiceApps -notcontains $alsoRingUserTargetVoiceAppId) {
                
                                    $nestedVoiceApps += $alsoRingUserTargetVoiceAppId
                
                                }            
                
                            }
                
                            else {
                
                                $mdAlsoRingTarget = "-.-> userForwardingTarget$alsoRingUserTargetVoiceAppId([$forwardingTargetType<br>$userForwardingTarget])"
                
                            }
                
                
                        }
                
                        else {

                            $forwardingTargetType = "User"
    
                            if ($null -eq $userForwardingTarget) {
        
                                $userForwardingTarget = "External Tenant"
                                $forwardingTargetType = "Federated User"
        
                            }

                            if ($StandAlone -eq $false -and $forwardingTargetType -eq "User") {

                                $alsoRingUserTargetUserId = (Get-CsOnlineUser -Identity $userCallingSettings.ForwardingTarget).Identity
                
                                $mdAlsoRingTarget = "-.-> $alsoRingUserTargetUserId($forwardingTargetType<br> $userForwardingTarget)"

                                $allMermaidNodes += $alsoRingUserTargetUserId
                
                                if ($nestedVoiceApps -notcontains $alsoRingUserTargetUserId) {
                
                                    $nestedVoiceApps += $alsoRingUserTargetUserId
                
                                }
                
                            }
                
                            else {
                
                                $mdAlsoRingTarget = "-.-> userForwardingTarget$UserId($forwardingTargetType<br>$userForwardingTarget)"

                                $mdSubGraphUnansweredSettingsAddition1 = "userForwardingTarget$UserId($forwardingTargetType<br> $userForwardingTarget)"
                                $mdSubGraphUnansweredSettingsAddition2 = "userForwardingTarget$UserId --> "
                
                            }
                
                        }

    
                    }
    
                    else {
    
                        $userForwardingTarget = $userCallingSettings.ForwardingTarget
                        $forwardingTargetType = "External Number"

                        if ($ObfuscatePhoneNumbers -eq $true) {

                            $userForwardingTarget = $userForwardingTarget.Remove(($userForwardingTarget.Length -4)) + "****"
        
                        }            

                        $mdAlsoRingTarget = "-.-> userForwardingTarget$UserId($forwardingTargetType<br>$userForwardingTarget)"
    
                    }
    
    
                    $mdUserCallingSettings = @"
$userNode --> userForwarding$UserId(Also Ring)
userForwarding$UserId -.-> userParallelRing$userId(Teams Clients<br> $($teamsUser.DisplayName)) ---> userForwardingResult$UserId
userForwarding$UserId $mdAlsoRingTarget
subgraph subgraphSettings$UserId[ ]
$mdSubGraphUnansweredSettingsAddition1

"@

$allMermaidNodes += @("userForwarding$UserId","userParallelRing$userId","userForwardingResult$UserId","userForwardingTarget$UserId")
$allSubgraphs += "subgraphSettings$UserId"

                    $mdUserCallingSettingsAddition = @"
$mdSubGraphUnansweredSettingsAddition2 userForwardingResult$UserId{Call Answered?}
$subgraphUnansweredSettings
userForwardingResult$UserId --> |No| userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)]
end
userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)] $mdUnansweredTarget
userForwardingResult$UserId --> |Yes| userForwardingConnected$UserId((Call Connected))

"@

                    $allMermaidNodes += @("userForwardingTarget$UserId","userForwardingResult$UserId","userForwardingTimeout$UserId","userForwardingConnected$UserId")

                    $mdUserCallingSettings += $mdUserCallingSettingsAddition

    
                }

            }

        }

        # user is forwarding enabled but not unanswered enabled
        elseif ($userCallingSettings.IsForwardingEnabled -and !$userCallingSettings.IsUnansweredEnabled) {

            #Write-Host "User is forwarding enabled but not unanswered enabled"

            switch ($userCallingSettings.ForwardingTargetType) {
                Group {
    
                    switch ($userCallingSettings.CallGroupOrder) {
                        InOrder {
                            $ringOrder = "Serial"
                        }
                        Simultaneous {
                            $ringOrder = "Simultaneous"
                        }
                        Default {}
                    }
    
                    $mdUserCallingSettings = @"
$userNode --> userForwarding$UserId(Also Ring)
userForwarding$UserId -.-> userParallelRing$userId(Teams Clients<br> $($teamsUser.DisplayName))
userForwarding$UserId -.-> subgraphCallGroups$UserId
subgraph subgraphSettings$UserId[ ]

"@

$allMermaidNodes += @("userForwarding$UserId","userParallelRing$userId")
$allSubgraphs += "subgraphSettings$UserId"
    
                    $mdSubgraphcallGroups = @"
subgraph subgraphCallGroups$UserId[Call Group of $($teamsUser.DisplayName)]
direction LR
ringType$UserId[($ringOrder Ring)]
 
"@

$allMermaidNodes += "ringType$UserId"
$allSubgraphs += "subgraphCallGroups$UserId"
    
                    $callGroupMemberCounter = 1
    
                    foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {
    
                        $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)
    
                        if ($ringOrder -eq "Serial") {
    
                            $linkNumber = " |$callGroupMemberCounter|"
    
                        }
    
                        else {
    
                            $linkNumber = $null
    
                        }
    
                        $callGroupRing = "ringType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"
    
                        $mdSubgraphcallGroups += $callGroupRing

                        $allMermaidNodes += "callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter"
    
                        $callGroupMemberCounter ++
                    }
    
                    $mdUserCallingSettings += $mdSubgraphcallGroups
    
                    $mdUserCallingSettingsAddition = @"
    end
    end

"@
    
                    $mdUserCallingSettings += $mdUserCallingSettingsAddition
    
    
    
    
                }
                SingleTarget {
    
                    if ($userCallingSettings.ForwardingTarget -match "sip:" -or $userCallingSettings.ForwardingTarget -notmatch "\+") {
    
                        $userForwardingTarget = (Get-CsOnlineUser -Identity $userCallingSettings.ForwardingTarget).DisplayName

                        if (Get-CsOnlineUser -AccountType ResourceAccount |  Where-Object {$_.SipAddress -eq $userCallingSettings.ForwardingTarget}) {

                            $checkUserAccountType = Get-CsOnlineApplicationInstance -Identity $($userCallingSettings.ForwardingTarget).Replace("sip:","")
                            
                        }
                
                        else {
                
                            $checkUserAccountType = $null
                
                        }

                        if ($checkUserAccountType) {

                            switch ($checkUserAccountType.ApplicationId) {
                                # Call Queue
                                11cd3e2e-fccb-42ad-ad00-878b93575e07 {
                                    $forwardingTargetType = "Call Queue"
                                }
                                # Auto Attendant
                                ce933385-9390-45d1-9512-c8d228074e07 {
                                    $forwardingTargetType = "Auto Attendant"
                                }
                                Default {}
                            }
                
                            if ($StandAlone -eq $false) {
                
                                switch ($forwardingTargetType) {
                                    "Auto Attendant" {
                
                                        $alsoRingUserTargetVoiceAppId = ($allAutoAttendants| Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity
                                        $mdAlsoRingWarningMessage = "--> warning$alsoRingUserTargetVoiceAppId[[Warning: Also Ring is enabled<br> $forwardingTargetType will answer automatically!]] --> "
                                    }
                                    "Call Queue" {
                
                                        $alsoRingUserTargetVoiceAppId = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity
                                        $checkAlsoRingTargetCqOverflowThreshold = ($allCallQueues | Where-Object {$_.Identity -eq $alsoRingUserTargetVoiceAppId}).OverflowThreshold

                                        if ($checkAlsoRingTargetCqOverflowThreshold -eq 0) {
                                            $mdAlsoRingWarningMessage = "-.-> "
                                        }

                                        else {
                                            $mdAlsoRingWarningMessage = "--> warning$alsoRingUserTargetVoiceAppId[[Warning: Also Ring is enabled<br> $forwardingTargetType will answer automatically!]] --> "
                                        }
                                    }
                                    Default {}

                                }

                                $allMermaidNodes += "warning$alsoRingUserTargetVoiceAppId"
                
                                #2 variables without whitespace!
                                $mdAlsoRingTarget = "$mdAlsoRingWarningMessage$alsoRingUserTargetVoiceAppId([$forwardingTargetType<br>$userForwardingTarget])"
                
                                if ($nestedVoiceApps -notcontains $alsoRingUserTargetVoiceAppId) {
                
                                    $nestedVoiceApps += $alsoRingUserTargetVoiceAppId
                
                                }            
                
                            }
                
                            else {
                
                                $mdAlsoRingTarget = "-.-> userForwardingTarget$alsoRingUserTargetVoiceAppId([$forwardingTargetType<br>$userForwardingTarget])"
                
                            }
                
                
                        }
                
                        else {

                            $forwardingTargetType = "User"
    
                            if ($null -eq $userForwardingTarget) {
        
                                $userForwardingTarget = "External Tenant"
                                $forwardingTargetType = "Federated User"
        
                            }

                            if ($StandAlone -eq $false -and $forwardingTargetType -eq "User") {

                                $alsoRingUserTargetUserId = (Get-CsOnlineUser -Identity $userCallingSettings.ForwardingTarget).Identity
                
                                $mdAlsoRingTarget = "-.-> $alsoRingUserTargetUserId($forwardingTargetType<br> $userForwardingTarget)"

                                $allMermaidNodes += $alsoRingUserTargetUserId
                
                                if ($nestedVoiceApps -notcontains $alsoRingUserTargetUserId) {
                
                                    $nestedVoiceApps += $alsoRingUserTargetUserId
                
                                }
                
                            }
                
                            else {
                
                                $mdAlsoRingTarget = "-.-> userForwardingTarget$UserId($forwardingTargetType<br>$userForwardingTarget)"
                
                            }
                
                        }

    
                    }
    
                    else {
    
                        $userForwardingTarget = $userCallingSettings.ForwardingTarget
                        $forwardingTargetType = "External Number"

                        if ($ObfuscatePhoneNumbers -eq $true) {

                            $userForwardingTarget = $userForwardingTarget.Remove(($userForwardingTarget.Length -4)) + "****"
        
                        }            

                        $mdAlsoRingTarget = "-.-> userForwardingTarget$UserId($forwardingTargetType<br>$userForwardingTarget)"
    
                    }
    
    
                    $mdUserCallingSettings = @"
$userNode --> userForwarding$UserId(Also Ring)
userForwarding$UserId -.-> userParallelRing$userId(Teams Clients<br> $($teamsUser.DisplayName))
userForwarding$UserId $mdAlsoRingTarget
"@

$allMermaidNodes += @("userForwarding$UserId","userParallelRing$userId","userForwardingTarget$UserId")
    
                }

            }

        }

        # user is unanswered enabled but not forwarding enabled
        elseif ($userCallingSettings.IsUnansweredEnabled -and !$userCallingSettings.IsForwardingEnabled) {

            #Write-Host "User is unanswered enabled but not forwarding enabled"

            switch ($userCallingSettings.UnansweredTargetType) {
                MyDelegates {

                    $ringOrder = "Simultaneous"

                    $subgraphUnansweredSettings = @"
subgraph subgraphdelegates$UserId[Delegates of $($teamsUser.DisplayName)]
direction LR
delegateRingType$UserId[($ringOrder Ring)]
    
"@

$allMermaidNodes += "delegateRingType$UserId"
$allSubgraphs += "subgraphdelegates$UserId"

                    $delegateCounter = 1

                    foreach ($delegate in $userCallingSettings.Delegates) {

                        $delegateUserObject = (Get-CsOnlineUser -Identity $delegate.Id)

                        if ($ringOrder -eq "Serial") {

                            $linkNumber = " |$delegateCounter|"

                        }

                        else {

                            $linkNumber = $null

                        }

                        $delegateRing = "delegateRingType$UserId -.->$linkNumber delegateMember$($delegateUserObject.Identity)$delegateCounter($($delegateUserObject.DisplayName))`n"

                        $subgraphUnansweredSettings += $delegateRing

                        $allMermaidNodes += "delegateMember$($delegateUserObject.Identity)$delegateCounter"

                        $delegateCounter ++
                    }

                    $subgraphUnansweredSettings += "`nend"

                    $mdUnansweredTarget = "--> subgraphdelegates$UserId"


                }

                Voicemail {
                    $mdUnansweredTarget = "--> userVoicemail$UserId(Voicemail<br> $($teamsUser.DisplayName))"
                    $subgraphUnansweredSettings = $null

                    $allMermaidNodes += "userVoicemail$UserId"
                }

                Group {

                    switch ($userCallingSettings.CallGroupOrder) {
                        InOrder {
                            $ringOrder = "Serial"
                        }
                        Simultaneous {
                            $ringOrder = "Simultaneous"
                        }
                        Default {}
                    }
            
                    $subgraphUnansweredSettings = @"
subgraph subgraphCallGroups$UserId[Call Group of $($teamsUser.DisplayName)]
direction LR
callGroupRingType$UserId[($ringOrder Ring)]
    
"@

$allMermaidNodes += "callGroupRingType$UserId"
$allSubgraphs += "subgraphCallGroups$UserId"

                    $callGroupMemberCounter = 1

                    foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {

                        $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)

                        if ($ringOrder -eq "Serial") {

                            $linkNumber = " |$callGroupMemberCounter|"

                        }

                        else {

                            $linkNumber = $null

                        }

                        $callGroupRing = "callGroupRingType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"

                        $subgraphUnansweredSettings += $callGroupRing

                        $allMermaidNodes += "callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter"

                        $callGroupCounter ++
                    }

                    $subgraphUnansweredSettings += "`nend"

                    $mdUnansweredTarget = "--> subgraphCallGroups$UserId"

                }
                SingleTarget {

                    if ($userCallingSettings.UnansweredTarget -match "sip:" -or $userCallingSettings.UnansweredTarget -notmatch "\+") {

                        $userForwardingTarget = (Get-CsOnlineUser -Identity $userCallingSettings.UnansweredTarget).DisplayName

                        if (Get-CsOnlineUser -AccountType ResourceAccount |  Where-Object {$_.SipAddress -eq $userCallingSettings.UnansweredTarget}) {

                            $checkUserAccountType = Get-CsOnlineApplicationInstance -Identity $($userCallingSettings.UnansweredTarget).Replace("sip:","")
                            
                        }

                        else {

                            $checkUserAccountType = $null

                        }

                        if ($checkUserAccountType) {

                            switch ($checkUserAccountType.ApplicationId) {
                                # Call Queue
                                11cd3e2e-fccb-42ad-ad00-878b93575e07 {
                                    $forwardingTargetType = "Call Queue"
                                }
                                # Auto Attendant
                                ce933385-9390-45d1-9512-c8d228074e07 {
                                    $forwardingTargetType = "Auto Attendant"
                                }
                                Default {}
                            }

                            if ($StandAlone -eq $false) {

                                switch ($forwardingTargetType) {
                                    "Auto Attendant" {

                                        $unansweredUserTargetVoiceAppId = ($allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity

                                    }
                                    "Call Queue" {

                                        $unansweredUserTargetVoiceAppId = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $checkUserAccountType.ObjectId}).Identity

                                    }
                                    Default {}
                                }

                                $mdUnansweredTarget = "--> $unansweredUserTargetVoiceAppId([$forwardingTargetType<br> $userForwardingTarget])"

                                if ($nestedVoiceApps -notcontains $unansweredUserTargetVoiceAppId) {

                                    $nestedVoiceApps += $unansweredUserTargetVoiceAppId
                
                                }

                            }

                            else {

                                $mdUnansweredTarget = "--> userUnansweredTarget$unansweredUserTargetVoiceAppId([$forwardingTargetType<br> $userForwardingTarget])"

                            }


                        }

                        else {
                            
                            $forwardingTargetType = "User"
        
                            if ($null -eq $userForwardingTarget) {
        
                                $userForwardingTarget = "External Tenant"
                                $forwardingTargetType = "Federated User"
        
                            }

                            if ($StandAlone -eq $false -and $forwardingTargetType -eq "User") {

                                $unansweredUserTargetUserId = (Get-CsOnlineUser -Identity $userCallingSettings.UnansweredTarget).Identity

                                $mdUnansweredTarget = "-->$unansweredUserTargetUserId($forwardingTargetType<br> $userForwardingTarget)"

                                $allMermaidNodes += $unansweredUserTargetUserId

                                if ($nestedVoiceApps -notcontains $unansweredUserTargetUserId) {
    
                                    $nestedVoiceApps += $unansweredUserTargetUserId
                
                                }
    
                            }

                            else {

                                $mdUnansweredTarget = "--> userUnansweredTarget$UserId($forwardingTargetType<br> $userForwardingTarget)"

                            }

    
                        }
    
                    }
    
                    else {
    
                        $userForwardingTarget = $userCallingSettings.UnansweredTarget
                        $forwardingTargetType = "External Number"

                        if ($ObfuscatePhoneNumbers -eq $true) {

                            $userForwardingTarget = $userForwardingTarget.Remove(($userForwardingTarget.Length -4)) + "****"
        
                        }            

                        $mdUnansweredTarget = "--> userUnansweredTarget$UserId($forwardingTargetType<br> $userForwardingTarget)"
    
                    }

                    $subgraphUnansweredSettings = $null
                    
                    $allMermaidNodes += "userUnansweredTarget$UserId"

                }
                Default {}
            }

            $mdUserCallingSettings = @"
    
$userNode --> userParallelRing$userId(Teams Clients<br> $($teamsUser.DisplayName))
subgraph subgraphSettings$UserId[ ]
userParallelRing$userId --> userForwardingResult$UserId{Call Answered?}

"@

$allMermaidNodes += @("userParallelRing$userId","userForwardingResult$UserId")
$allSubgraphs += "subgraphSettings$UserId"

            $mdUserCallingSettingsAddition = @"
userForwardingResult$UserId --> |No| userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)]
$subgraphUnansweredSettings
end
userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)] $mdUnansweredTarget
userForwardingResult$UserId --> |Yes| userForwardingConnected$UserId((Call Connected))

"@

$allMermaidNodes += @("userForwardingResult$UserId","userForwardingTimeout$UserId","userForwardingConnected$UserId")
$allSubgraphs += "subgraphSettings$UserId"

            $mdUserCallingSettings += $mdUserCallingSettingsAddition

            
        }

    }

    $mdFlowChart += $mdUserCallingSettings

    if ($SetClipBoard) {

        $mdFlowChart | Set-Clipboard

    }

    if ($ExportSvg -or $PreviewSvg) {

        $mdFlowChart = $mdFlowChart.Trim()

        $base64FriendlyFlowChart = @"
$mdFlowChart

"@

        $flowChartBytes = [System.Text.Encoding]::ASCII.GetBytes($base64FriendlyFlowChart)
        $encodedUrl =[Convert]::ToBase64String($flowChartBytes)

        $url = "https://mermaid.ink/svg/$encodedUrl"

    }

    if ($ExportSvg) {

        if (!(Test-Path -Path "$filePath")) {

            New-Item -Path $filePath -ItemType Directory

        }

        (Invoke-WebRequest -Uri $url).Content > "$filePath\UserCallingSettings_$($teamsUser.DisplayName).svg"

    }

    if ($PreviewSvg) {

        Start-Process $url

    }

}

