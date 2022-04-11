function Get-TeamsUserCallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$UserId
    )

    $teamsUser = Get-CsOnlineUser -Identity $UserId

    $userCallingSettings = Get-CsUserCallingSettings -Identity $UserId

    $userCallingSettings

    [int]$userUnansweredTimeoutMinutes = ($userCallingSettings.UnansweredDelay).Split(":")[1]
    [int]$userUnansweredTimeoutSeconds = ($userCallingSettings.UnansweredDelay).Split(":")[-1]

    if ($userUnansweredTimeoutMinutes -eq 1) {

        $userUnansweredTimeout = "60 Seconds"

    }

    else {

        $userUnansweredTimeout = "$userUnansweredTimeoutSeconds Seconds"

    }


    # user is neither forwarding or unanswered enabled
    if (!$userCallingSettings.IsForwardingEnabled -and !$userCallingSettings.IsUnansweredEnabled) {

        Write-Host "User is neither forwaring or unanswered enabled"

        $mdUserCallingSettings = $null

    }

    # user is immediate forwarding enabled
    elseif ($userCallingSettings.ForwardingType -eq "Immediate") {

        Write-Host "user is immediate forwarding enabled."

        switch ($userCallingSettings.ForwardingTargetType) {
            MyDelegates {

                $mdUserCallingSettings = @"

                $UserId --> userForwarding$UserId(Immediate Forwarding <br> Delegates)

                subgraph subgraphSettings$UserId[User Forwarding Settings]
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

                $mdUserCallingSettingsAddition = @"

                end
                userForwardingResult$UserId --> |No| userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)]
                subgraphDelegates$UserId --> userForwardingResult$UserId{Call Connected?}
                end
                userForwardingTimeout$UserId[(Timeout: $userUnansweredTimeout)] --> voicemail$UserId((Voicemail <br> $($teamsUser.DisplayName)))
                userForwardingResult$UserId --> |Yes| userForwardingConnected$UserId((Call Connected))

"@

                $mdUserCallingSettings += $mdUserCallingSettingsAddition

            }
            Voicemail {



            }
            Group {



            }
            SingleTarget {



            }
            Default {}
        }

    }

    # user is either forwarding or unansered enabled
    else {

        # user is forwarding and unanswered enabled
        if ($userCallingSettings.IsForwardingEnabled -and $userCallingSettings.IsUnansweredEnabled) {

            Write-Host "user is forwaring and unanswered enabled"

        }

        # user is forwarding enabled but not unanswered enabled
        elseif ($userCallingSettings.IsForwardingEnabled -and !$userCallingSettings.IsUnansweredEnabled) {

            Write-Host "user is forwarding enabled but not unanswered enabled"

        }

        # user is unanswered enabled but not forwarding enabled
        elseif ($userCallingSettings.IsUnansweredEnabled -and !$userCallingSettings.IsForwardingEnabled) {

            Write-Host "user is unanswered enabled but not forwarding enabled"
            
        }

    }

}

. Get-TeamsUserCallFlow -UserId "fa19b242-8bae-419d-a4eb-12796577c81f"

$mdUserCallingSettings