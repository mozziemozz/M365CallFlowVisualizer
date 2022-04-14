function Get-TeamsUserCallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$UserId,
        [Parameter(Mandatory=$false)][bool]$standAlone = $true

    )

    $teamsUser = Get-CsOnlineUser -Identity $UserId

    $userCallingSettings = Get-CsUserCallingSettings -Identity $UserId

    $userCallingSettings

    [int]$userUnansweredTimeoutMinutes = ($userCallingSettings.UnansweredDelay).Split(":")[1]
    [int]$userUnansweredTimeoutSeconds = ($userCallingSettings.UnansweredDelay).Split(":")[-1]

    $mdFlowChart = "flowchart TB`n"

    if ($standAlone) {

        $userNode = "$UserId(User<br> $($teamsUser.DisplayName))"

    }

    else {

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

        Write-Host "User is neither forwaring or unanswered enabled"

        $mdUserCallingSettings = @"
        $userNode
"@

    }

    # user is immediate forwarding enabled
    elseif ($userCallingSettings.ForwardingType -eq "Immediate") {

        Write-Host "user is immediate forwarding enabled."

        switch ($userCallingSettings.ForwardingTargetType) {
            MyDelegates {

                $mdUserCallingSettings = @"

                $userNode --> userForwarding$UserId(Immediate Forwarding)

                subgraph subgraphSettings$UserId[ ]
                userForwarding$UserId --> subgraphDelegates$UserId

"@

                $mdSubgraphDelegates = @"

                subgraph subgraphDelegates$UserId[Delegates of $($teamsUser.DisplayName)]
                direction LR
                ringType$UserId[(Simultaneous Ring)]

"@

                $delegateCounter = 1

                foreach ($delegate in $userCallingSettings.Delegates) {

                    $delegateUserObject = (Get-CsOnlineUser -Identity $delegate.Id)

                    $delegateRing = "                ringType$UserId -.-> delegate$($delegateUserObject.Identity)$delegateCounter($($delegateUserObject.DisplayName))`n"

                    $mdSubgraphDelegates += $delegateRing

                    $delegateCounter ++
                }

                $mdUserCallingSettings += $mdSubgraphDelegates

                switch ($userCallingSettings.UnansweredTargetType) {
                    Voicemail {
                        $mdUnansweredTarget = "--> userVoicemail$UserId(Voicemail<br> $($teamsUser.DisplayName))"
                        $subgraphUnansweredSettings = $null
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

                        $callGroupMemberCounter = 1

                        foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {

                            $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)

                            if ($ringOrder -eq "Serial") {

                                $linkNumber = " |$callGroupMemberCounter|"

                            }

                            else {

                                $linkNumber = $null

                            }

                            $callGroupRing = "                       callGroupRingType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"

                            $subgraphUnansweredSettings += $callGroupRing

                            $callGroupMemberCounter ++
                        }

                        $subgraphUnansweredSettings += "`n                end"

                        $mdUnansweredTarget = "--> subgraphCallGroups$UserId"

    
                    }
                    SingleTarget {

                        if ($userCallingSettings.UnansweredTarget -match "sip:" -or $userCallingSettings.UnansweredTarget -notmatch "\+") {

                            $userForwardingTarget = (Get-CsOnlineUser -Identity $userCallingSettings.UnansweredTarget).DisplayName
                            $forwardingTargetType = "Internal User"
        
                            if ($null -eq $userForwardingTarget) {
        
                                $userForwardingTarget = "External Tenant"
                                $forwardingTargetType = "Federated User"
        
                            }
        
                        }
        
                        else {
        
                            $userForwardingTarget = $userCallingSettings.UnansweredTarget
                            $forwardingTargetType = "External PSTN"
        
                        }

                        $mdUnansweredTarget = "--> userUnansweredTarget$UserId($forwardingTargetType<br> $userForwardingTarget)"
                        $subgraphUnansweredSettings = $null       

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

                $mdUserCallingSettings += $mdUserCallingSettingsAddition

            }
            Voicemail {

                $mdUserCallingSettings = @"

                $userNode --> userForwarding$UserId(Immediate Forwarding)

                subgraph subgraphSettings$UserId[ ]
                userForwarding$UserId --> voicemail$UserId(Voicemail<br> $($teamsUser.DisplayName))
                end

"@


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

                $mdSubgraphcallGroups = @"

                subgraph subgraphCallGroups$UserId[Call Group of $($teamsUser.DisplayName)]
                direction LR
                ringType$UserId[($ringOrder Ring)]

"@

                $callGroupMemberCounter = 1

                foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {

                    $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)

                    if ($ringOrder -eq "Serial") {

                        $linkNumber = " |$callGroupMemberCounter|"

                    }

                    else {

                        $linkNumber = $null

                    }

                    $callGroupRing = "                ringType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"

                    $mdSubgraphcallGroups += $callGroupRing

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
                    $forwardingTargetType = "Internal User"

                    if ($null -eq $userForwardingTarget) {

                        $userForwardingTarget = "External Tenant"
                        $forwardingTargetType = "Federated User"

                    }

                }

                else {

                    $userForwardingTarget = $userCallingSettings.ForwardingTarget
                    $forwardingTargetType = "External PSTN"

                }


                $mdUserCallingSettings = @"

                $userNode --> userForwarding$UserId(Immediate Forwarding)

                subgraph subgraphSettings$UserId[ ]
                userForwarding$UserId --> userForwardingTarget$UserId($forwardingTargetType<br> $userForwardingTarget)
                end

"@

            }
            Default {}
        }

    }

    # user is either forwarding or unansered enabled
    else {

        # user is forwarding and unanswered enabled
        if ($userCallingSettings.IsForwardingEnabled -and $userCallingSettings.IsUnansweredEnabled) {

            Write-Host "user is forwaring and unanswered enabled"

            switch ($userCallingSettings.UnansweredTargetType) {
                MyDelegates {

                    $ringOrder = "Simultaneous"

                    $subgraphUnansweredSettings = @"

                    subgraph subgraphdelegates$UserId[Delegates of $($teamsUser.DisplayName)]
                    direction LR
                    delegateRingType$UserId[($ringOrder Ring)]
    
"@

                    $delegateCounter = 1

                    foreach ($delegate in $userCallingSettings.Delegates) {

                        $delegateUserObject = (Get-CsOnlineUser -Identity $delegate.Id)

                        if ($ringOrder -eq "Serial") {

                            $linkNumber = " |$delegateCounter|"

                        }

                        else {

                            $linkNumber = $null

                        }

                        $delegateRing = "                       delegateRingType$UserId -.->$linkNumber delegateMember$($delegateUserObject.Identity)$delegateCounter($($delegateUserObject.DisplayName))`n"

                        $subgraphUnansweredSettings += $delegateRing

                        $delegateCounter ++
                    }

                    $subgraphUnansweredSettings += "`n                end"

                    $mdUnansweredTarget = "--> subgraphdelegates$UserId"


                }

                Voicemail {
                    $mdUnansweredTarget = "--> userVoicemail$UserId(Voicemail<br> $($teamsUser.DisplayName))"
                    $subgraphUnansweredSettings = $null
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
    
                        $callGroupMemberCounter = 1
    
                        foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {
    
                            $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)
    
                            if ($ringOrder -eq "Serial") {
    
                                $linkNumber = " |$callGroupMemberCounter|"
    
                            }
    
                            else {
    
                                $linkNumber = $null
    
                            }
    
                            $callGroupRing = "                       callGroupRingType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"
    
                            $subgraphUnansweredSettings += $callGroupRing
    
                            $callGroupCounter ++
                        }
    
                        $subgraphUnansweredSettings += "`n                end"
    
                        $mdUnansweredTarget = "--> subgraphCallGroups$UserId"
    
    
                    }
    

                }
                SingleTarget {

                    if ($userCallingSettings.UnansweredTarget -match "sip:" -or $userCallingSettings.UnansweredTarget -notmatch "\+") {

                        $userForwardingTarget = (Get-CsOnlineUser -Identity $userCallingSettings.UnansweredTarget).DisplayName
                        $forwardingTargetType = "Internal User"
    
                        if ($null -eq $userForwardingTarget) {
    
                            $userForwardingTarget = "External Tenant"
                            $forwardingTargetType = "Federated User"
    
                        }
    
                    }
    
                    else {
    
                        $userForwardingTarget = $userCallingSettings.UnansweredTarget
                        $forwardingTargetType = "External PSTN"
    
                    }

                    $mdUnansweredTarget = "--> userUnansweredTarget$UserId($forwardingTargetType<br> $userForwardingTarget)"
                    $subgraphUnansweredSettings = $null       

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
    
                    $mdSubgraphDelegates = @"
    
                    subgraph subgraphDelegates$UserId[Delegates of $($teamsUser.DisplayName)]
                    direction LR
                    ringType$UserId[(Simultaneous Ring)]
    
"@
    
                    $delegateCounter = 1
    
                    foreach ($delegate in $userCallingSettings.Delegates) {
    
                        $delegateUserObject = (Get-CsOnlineUser -Identity $delegate.Id)
    
                        $delegateRing = "                ringType$UserId -.-> delegate$($delegateUserObject.Identity)$delegateCounter($($delegateUserObject.DisplayName))`n"
    
                        $mdSubgraphDelegates += $delegateRing
    
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
    
                    $mdUserCallingSettings += $mdUserCallingSettingsAddition
    
                }
                Voicemail {
    
                    $mdUserCallingSettings = @"
    
                    $userNode --> userForwarding$UserId(Immediate Forwarding)
    
                    subgraph subgraphSettings$UserId[ ]
                    userForwarding$UserId --> voicemail$UserId(Voicemail<br> $($teamsUser.DisplayName))
                    end
    
"@
    
    
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
    
                    $mdSubgraphcallGroups = @"
    
                    subgraph subgraphCallGroups$UserId[Call Group of $($teamsUser.DisplayName)]
                    direction LR
                    ringType$UserId[($ringOrder Ring)]
    
"@
    
                    $callGroupMemberCounter = 1
    
                    foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {
    
                        $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)
    
                        if ($ringOrder -eq "Serial") {
    
                            $linkNumber = " |$callGroupMemberCounter|"
    
                        }
    
                        else {
    
                            $linkNumber = $null
    
                        }
    
                        $callGroupRing = "                ringType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"
    
                        $mdSubgraphcallGroups += $callGroupRing
    
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
    
                    $mdUserCallingSettings += $mdUserCallingSettingsAddition
    
    
    
    
                }
                SingleTarget {
    
                    if ($userCallingSettings.ForwardingTarget -match "sip:" -or $userCallingSettings.ForwardingTarget -notmatch "\+") {
    
                        $userForwardingTarget = (Get-CsOnlineUser -Identity $userCallingSettings.ForwardingTarget).DisplayName
                        $forwardingTargetType = "Internal User"
    
                        if ($null -eq $userForwardingTarget) {
    
                            $userForwardingTarget = "External Tenant"
                            $forwardingTargetType = "Federated User"
    
                        }
    
                    }
    
                    else {
    
                        $userForwardingTarget = $userCallingSettings.ForwardingTarget
                        $forwardingTargetType = "External PSTN"
    
                    }
    
    
                    $mdUserCallingSettings = @"
    
                    $userNode --> userForwarding$UserId(Also Ring)
                    userForwarding$UserId -.-> userParallelRing$userId(Teams Clients<br> $($teamsUser.DisplayName)) ---> userForwardingResult$UserId
                    userForwarding$UserId -.-> userForwardingTarget$UserId

    
                    subgraph subgraphSettings$UserId[ ]
                    userForwardingTarget$UserId($forwardingTargetType<br> $userForwardingTarget)                    
    
"@

                    $mdUserCallingSettingsAddition = @"
                        
                    userForwardingTarget$UserId --> userForwardingResult$UserId{Call Answered?}    
                    $subgraphUnansweredSettings
                    userForwardingResult$UserId --> |No| userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)]
                    end
                    userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)] $mdUnansweredTarget
                    userForwardingResult$UserId --> |Yes| userForwardingConnected$UserId((Call Connected))    

"@

                    $mdUserCallingSettings += $mdUserCallingSettingsAddition

    
                }

            }

        }

        # user is forwarding enabled but not unanswered enabled
        elseif ($userCallingSettings.IsForwardingEnabled -and !$userCallingSettings.IsUnansweredEnabled) {

            Write-Host "user is forwarding enabled but not unanswered enabled"

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
    
                    $mdSubgraphcallGroups = @"
    
                    subgraph subgraphCallGroups$UserId[Call Group of $($teamsUser.DisplayName)]
                    direction LR
                    ringType$UserId[($ringOrder Ring)]
    
"@
    
                    $callGroupMemberCounter = 1
    
                    foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {
    
                        $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)
    
                        if ($ringOrder -eq "Serial") {
    
                            $linkNumber = " |$callGroupMemberCounter|"
    
                        }
    
                        else {
    
                            $linkNumber = $null
    
                        }
    
                        $callGroupRing = "                ringType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"
    
                        $mdSubgraphcallGroups += $callGroupRing
    
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
                        $forwardingTargetType = "Internal User"
    
                        if ($null -eq $userForwardingTarget) {
    
                            $userForwardingTarget = "External Tenant"
                            $forwardingTargetType = "Federated User"
    
                        }
    
                    }
    
                    else {
    
                        $userForwardingTarget = $userCallingSettings.ForwardingTarget
                        $forwardingTargetType = "External PSTN"
    
                    }
    
    
                    $mdUserCallingSettings = @"
    
                    $userNode --> userForwarding$UserId(Also Ring)
                    userForwarding$UserId -.-> userParallelRing$userId(Teams Clients<br> $($teamsUser.DisplayName))
                    userForwarding$UserId -.-> userForwardingTarget$UserId

    
                    subgraph subgraphSettings$UserId[ ]
                    userForwardingTarget$UserId($forwardingTargetType<br> $userForwardingTarget)
                    end
    
"@
    
                }

            }

        }

        # user is unanswered enabled but not forwarding enabled
        elseif ($userCallingSettings.IsUnansweredEnabled -and !$userCallingSettings.IsForwardingEnabled) {

            Write-Host "user is unanswered enabled but not forwarding enabled"

            switch ($userCallingSettings.UnansweredTargetType) {
                MyDelegates {

                    $ringOrder = "Simultaneous"

                    $subgraphUnansweredSettings = @"

                    subgraph subgraphdelegates$UserId[Delegates of $($teamsUser.DisplayName)]
                    direction LR
                    delegateRingType$UserId[($ringOrder Ring)]
    
"@

                    $delegateCounter = 1

                    foreach ($delegate in $userCallingSettings.Delegates) {

                        $delegateUserObject = (Get-CsOnlineUser -Identity $delegate.Id)

                        if ($ringOrder -eq "Serial") {

                            $linkNumber = " |$delegateCounter|"

                        }

                        else {

                            $linkNumber = $null

                        }

                        $delegateRing = "                       delegateRingType$UserId -.->$linkNumber delegateMember$($delegateUserObject.Identity)$delegateCounter($($delegateUserObject.DisplayName))`n"

                        $subgraphUnansweredSettings += $delegateRing

                        $delegateCounter ++
                    }

                    $subgraphUnansweredSettings += "`n                end"

                    $mdUnansweredTarget = "--> subgraphdelegates$UserId"


                }

                Voicemail {
                    $mdUnansweredTarget = "--> userVoicemail$UserId(Voicemail<br> $($teamsUser.DisplayName))"
                    $subgraphUnansweredSettings = $null
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

                    $callGroupMemberCounter = 1

                    foreach ($callGroupMember in $userCallingSettings.CallGroupTargets) {

                        $callGroupUserObject = (Get-CsOnlineUser -Identity $callGroupMember)

                        if ($ringOrder -eq "Serial") {

                            $linkNumber = " |$callGroupMemberCounter|"

                        }

                        else {

                            $linkNumber = $null

                        }

                        $callGroupRing = "                       callGroupRingType$UserId -.->$linkNumber callGroupMember$($callGroupUserObject.Identity)$callGroupMemberCounter($($callGroupUserObject.DisplayName))`n"

                        $subgraphUnansweredSettings += $callGroupRing

                        $callGroupCounter ++
                    }

                    $subgraphUnansweredSettings += "`n                end"

                    $mdUnansweredTarget = "--> subgraphCallGroups$UserId"

                }
                SingleTarget {

                    if ($userCallingSettings.UnansweredTarget -match "sip:" -or $userCallingSettings.UnansweredTarget -notmatch "\+") {

                        $userForwardingTarget = (Get-CsOnlineUser -Identity $userCallingSettings.UnansweredTarget).DisplayName
                        $forwardingTargetType = "Internal User"
    
                        if ($null -eq $userForwardingTarget) {
    
                            $userForwardingTarget = "External Tenant"
                            $forwardingTargetType = "Federated User"
    
                        }
    
                    }
    
                    else {
    
                        $userForwardingTarget = $userCallingSettings.UnansweredTarget
                        $forwardingTargetType = "External PSTN"
    
                    }

                    $mdUnansweredTarget = "--> userUnansweredTarget$UserId($forwardingTargetType<br> $userForwardingTarget)"
                    $subgraphUnansweredSettings = $null       

                }
                Default {}
            }

            $mdUserCallingSettings = @"
    
            $userNode --> userParallelRing$userId(Teams Clients<br> $($teamsUser.DisplayName))
            
            subgraph subgraphSettings$UserId[ ]
            userParallelRing$userId --> userForwardingResult$UserId{Call Answered?}
"@

            $mdUserCallingSettingsAddition = @"

            userForwardingResult$UserId --> |No| userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)]
            $subgraphUnansweredSettings
            end
            userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)] $mdUnansweredTarget
            userForwardingResult$UserId --> |Yes| userForwardingConnected$UserId((Call Connected))

"@

            $mdUserCallingSettings += $mdUserCallingSettingsAddition

            
        }

    }

    $mdFlowChart += $mdUserCallingSettings

}

. Get-TeamsUserCallFlow -UserId "fa19b242-8bae-419d-a4eb-12796577c81f"

$mdFlowChart | Set-Clipboard