function Get-AutoAttendantHolidayCallFlow {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceAppId
    )

    $aaHolidayCallFlowId = $holidayCallFlow.Id

    $languageId = $aa.LanguageId

    # Get the current auto attendants holiday call flow and holiday call flow action
    $holidayCallFlow = $aa.CallFlows | Where-Object {$_.Id -eq $HolidayCallHandling.CallFlowId}
    $holidayCallFlowAction = $aa.Schedules | Where-Object {$_.Id -eq $HolidayCallHandling.ScheduleId}

    # Get the current auto attentans holiday call flow greeting
    if (!$holidayCallFlow.Greetings.ActiveType){
        $holidayCallFlowGreeting = "Greeting <br> None"
    }

    else {

        $holidayCallFlowGreeting = "Greeting <br> $($holidayCallFlow.Greetings.ActiveType)"

        if ($($holidayCallFlow.Greetings.ActiveType) -eq "TextToSpeech" -and $ShowTTSGreetingText) {

            $audioFileName = $null

            $holidayTTSGreetingValueExport = $holidayCallFlow.Greetings.TextToSpeechPrompt
            $holidayTTSGreetingExport = Optimize-DisplayName -String $holidayCallFlow.Greetings.TextToSpeechPrompt

            if ($ExportTTSGreetings) {

                $holidayTTSGreetingValueExport | Out-File "$FilePath\$($aaHolidayCallFlowId)_holidayCallFlowGreeting.txt"

                $ttsGreetings += ("click holidayCallFlowGreeting$($aaHolidayCallFlowId) " + '"' + "$FilePath\$($aaHolidayCallFlowId)_holidayCallFlowGreeting.txt" + '"')

            }

            if ($holidayTTSGreetingExport.Length -gt $truncateGreetings) {

                $holidayTTSGreetingExport = $holidayTTSGreetingExport.Remove($holidayTTSGreetingExport.Length - ($holidayTTSGreetingExport.Length -$truncateGreetings)).TrimEnd() + "..."
            
            }

            $holidayCallFlowGreeting += " <br> ''$holidayTTSGreetingExport''"
        
        }

        elseif ($($holidayCallFlow.Greetings.ActiveType) -eq "AudioFile" -and $ShowAudioFileName) {

            $holidayTTSGreetingExport = $null

            # Audio File Greeting Name
            $audioFileName = Optimize-DisplayName -String ($holidayCallFlow.Greetings.AudioFilePrompt.FileName)

            if ($ExportAudioFiles) {

                $content = Export-CsOnlineAudioFile -Identity $holidayCallFlow.Greetings.AudioFilePrompt.Id -ApplicationId OrgAutoAttendant
                [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)

                $audioFileNames += ("click holidayCallFlowGreeting$($aaHolidayCallFlowId) " + '"' + "$FilePath\$audioFileName" + '"')

            }


            if ($audioFileName.Length -gt $truncateGreetings) {

                $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"

            }

            $holidayCallFlowGreeting += " <br> $audioFileName"


        }

    }

    # Check if holiday callflow action is disconnect call
    if ($holidayCallFlowAction -eq "DisconnectCall") {

        $mdAutoAttendantHolidayCallFlow = "holidayCallFlowGreeting$($aaHolidayCallFlowId)>$holidayCallFlowGreeting] --> holidayCallFlow$($aaHolidayCallFlowId)(($holidayCallFlowAction))`n"

        $allMermaidNodes += @("holidayCallFlowGreeting$($aaHolidayCallFlowId)","holidayCallFlow$($aaHolidayCallFlowId)")

    }

    # Check if the holiday callflow action is transfer call to target
    else {

        $holidayCallFlowMenuOptions = $holidayCallFlow.Menu.MenuOptions

        if ($holidayCallFlowMenuOptions.Count -lt 2 -and !$holidayCallFlow.Menu.Prompts.ActiveType) {

            $mdHolidayCallFlowGreeting = "holidayCallFlowGreeting$($aaHolidayCallFlowId)>$holidayCallFlowGreeting] --> "

            $allMermaidNodes += @("holidayCallFlowGreeting$($aaHolidayCallFlowId)")

            $holidayCallFlowMenuOptionsKeyPress = $null

            $mdAutoAttendantHolidayCallFlowMenuOptions = $null

        }

        # Auto Attendant has multiple options / voice menu
        else {

            $holidayCallFlowMenuOptionsGreeting = "IVR Greeting <br> $($holidayCallFlow.Menu.Prompts.ActiveType)"

            if ($($holidayCallFlow.Menu.Prompts.ActiveType) -eq "TextToSpeech" -and $ShowTTSGreetingText) {

                $audioFileName = $null
    
                $holidayCallFlowMenuOptionsTTSGreetingValueExport = $holidayCallFlow.Menu.Prompts.TextToSpeechPrompt
                $holidayCallFlowMenuOptionsTTSGreetingValue = Optimize-DisplayName -String $holidayCallFlow.Menu.Prompts.TextToSpeechPrompt

                if ($ExportTTSGreetings) {

                    $holidayCallFlowMenuOptionsTTSGreetingValueExport | Out-File "$FilePath\$($aaHolidayCallFlowId)_holidayCallFlowMenuOptionsGreeting.txt"
    
                    $ttsGreetings += ("click holidayCallFlowMenuOptionsGreeting$($aaHolidayCallFlowId) " + '"' + "$FilePath\$($aaHolidayCallFlowId)_holidayCallFlowMenuOptionsGreeting.txt" + '"')
    
                }    
    
                if ($holidayCallFlowMenuOptionsTTSGreetingValue.Length -gt $truncateGreetings) {
    
                    $holidayCallFlowMenuOptionsTTSGreetingValue = $holidayCallFlowMenuOptionsTTSGreetingValue.Remove($holidayCallFlowMenuOptionsTTSGreetingValue.Length - ($holidayCallFlowMenuOptionsTTSGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                
                }
    
                $holidayCallFlowMenuOptionsGreeting += " <br> ''$holidayCallFlowMenuOptionsTTSGreetingValue''"
            
            }
    
            elseif ($($holidayCallFlow.Menu.Prompts.ActiveType) -eq "AudioFile" -and $ShowAudioFileName) {
    
                $holidayCallFlowMenuOptionsTTSGreetingValue = $null
    
                # Audio File Greeting Name
                $audioFileName = Optimize-DisplayName -String ($holidayCallFlow.Menu.Prompts.AudioFilePrompt.FileName)

                if ($ExportAudioFiles) {

                    $content = Export-CsOnlineAudioFile -Identity $holidayCallFlow.Menu.Prompts.AudioFilePrompt.Id -ApplicationId OrgAutoAttendant
                    [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
    
                    $audioFileNames += ("click holidayCallFlowMenuOptionsGreeting$($aaHolidayCallFlowId) " + '"' + "$FilePath\$audioFileName" + '"')
    
                }

    
                if ($audioFileName.Length -gt $truncateGreetings) {
    
                    $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                    $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
    
                }
    
                $holidayCallFlowMenuOptionsGreeting += " <br> $audioFileName"
    
    
            }

            if ($holidayCallFlow.ForceListenMenuEnabled -eq $true) {

                $holidayCallFlowForceListen = "<br> Force Listen: True"

            }

            else {

                $holidayCallFlowForceListen = "<br> Force Listen: False"

            }

            if ($aaIsVoiceResponseEnabled) {

                $holidayCallFlowVoiceResponse = " or <br> Voice Response <br> Language: $($aa.LanguageId)<br>Voice Style: $($aa.VoiceId)$holidayCallFlowForceListen"


            }

            else {

                $holidayCallFlowVoiceResponse = "$holidayCallFlowForceListen"

            }

            $holidayCallFlowMenuOptionsKeyPress = @"

holidayCallFlowMenuOptions$($aaHolidayCallFlowId){Key Press$holidayCallFlowVoiceResponse}
"@

            $mdHolidayCallFlowGreeting =@"
holidayCallFlowGreeting$($aaHolidayCallFlowId)>$holidayCallFlowGreeting] --> holidayCallFlowMenuOptionsGreeting$($aaHolidayCallFlowId)>$holidayCallFlowMenuOptionsGreeting] --> $holidayCallFlowMenuOptionsKeyPress

"@

            $mdAutoAttendantHolidayCallFlowMenuOptions =@"

"@

            $allMermaidNodes += @("holidayCallFlowMenuOptions$($aaHolidayCallFlowId)","holidayCallFlowGreeting$($aaHolidayCallFlowId)","holidayCallFlowMenuOptionsGreeting$($aaHolidayCallFlowId)")

        }

        $holidayCallFlowSharedVoicemailCounter = 1
        $holidayCallFlowUserOptionCounter = 1
        $holidayCallFlowPSTNOptionCounter = 1

        foreach ($MenuOption in $holidayCallFlowMenuOptions) {

            if ($holidayCallFlowMenuOptions.Count -lt 2 -and !$holidayCallFlow.Menu.Prompts.ActiveType) {

                $mdDtmfLink = $null
                $DtmfKey = $null
                $voiceResponse = $null

            }

            else {

                if ($aaIsVoiceResponseEnabled) {

                    if ($MenuOption.VoiceResponses) {

                        $voiceResponse = "/ Voice Response: ''$($MenuOption.VoiceResponses)''"

                    }

                    else {

                        $voiceResponse = "/ No Voice Response Configured"

                    }

                }

                else {

                    $voiceResponse = $null

                }

                [String]$DtmfKey = ($MenuOption.DtmfResponse)

                $DtmfKey = $DtmfKey.Replace("Tone","")

                $mdDtmfLink = "$holidayCallFlowMenuOptionsKeyPress --> |$DtmfKey $voiceResponse|"

            }

            # Get transfer target type
            $holidayCallFlowTargetType = $MenuOption.CallTarget.Type
            $holidayCallFlowAction = $MenuOption.Action

            if ($holidayCallFlowAction -eq "TransferCallToOperator") {

                if ($aaIsVoiceResponseEnabled) {

                    $holidayCallFlowOperatorVoiceResponse = "/ Voice Response: ''Operator''"

                    $mdDtmfLink = $mdDtmfLink.Replace($voiceResponse,$holidayCallFlowOperatorVoiceResponse)

                }

                else {

                    $holidayCallFlowOperatorVoiceResponse = $null

                }

                if ($ShowTTSGreetingText) {

                    $holidayCallFlowTransferGreetingOperatorValue = (. Get-IvrTransferMessage)[2]
                    $holidayCallFlowTransferGreetingOperatorValueExport = (. Get-IvrTransferMessage)[3]

                    if ($ExportTTSGreetings) {

                        $holidayCallFlowTransferGreetingOperatorValueExport | Out-File "$FilePath\$($aaHolidayCallFlowId)_autoAttendantholidayCallFlowTransferGreetingOperator.txt"
                        $ttsGreetings += ("click holidayCallFlowTransferGreetingOperator$($aaHolidayCallFlowId) " + '"' + "$FilePath\$($aaHolidayCallFlowId)_autoAttendantholidayCallFlowTransferGreetingOperator.txt" + '"')

                    }

                    if ($holidayCallFlowTransferGreetingOperatorValue.Length -gt $truncateGreetings) {

                        $holidayCallFlowTransferGreetingOperatorValue = $holidayCallFlowTransferGreetingOperatorValue.Remove($holidayCallFlowTransferGreetingOperatorValue.Length - ($holidayCallFlowTransferGreetingOperatorValue.Length -$truncateGreetings)).TrimEnd() + "..."
            
                    }

                    $holidayCallFlowTransferGreetingOperatorValue = "holidayCallFlowTransferGreetingOperator$($aaHolidayCallFlowId)>Greeting<br>Transfer Message<br>''$holidayCallFlowTransferGreetingOperatorValue''] -->"
    
                }

                else {

                    $holidayCallFlowTransferGreetingOperatorValueExport = $null
                    $holidayCallFlowTransferGreetingOperatorValue = "holidayCallFlowTransferGreetingOperator$($aaHolidayCallFlowId)>Greeting<br>Transfer Message] -->"

                }

                $allMermaidNodes += "holidayCallFlowTransferGreetingOperator$($aaHolidayCallFlowId)"

                $mdAutoAttendantholidayCallFlow = "$mdDtmfLink $holidayCallFlowTransferGreetingOperatorValue holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey($holidayCallFlowAction) --> $($OperatorIdentity)($OperatorTypeFriendly <br> $OperatorName)`n"

                $allMermaidNodes += @("holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey","$($OperatorIdentity)")

                $holidayCallFlowVoicemailSystemGreeting = $null

                if ($nestedVoiceApps -notcontains $OperatorIdentity -and $AddOperatorToNestedVoiceApps -eq $true) {

                    $nestedVoiceApps += $OperatorIdentity

                }

                if ($FindUserLinks -eq $true) {
         
                    . New-VoiceAppUserLinkProperties -userLinkUserId $OperatorIdentity -userLinkUserName $OperatorName -userLinkVoiceAppType "Auto Attendant" -userLinkVoiceAppActionType "TransferCallToOperator (DTMF Option: $DtmfKey)" -userLinkVoiceAppName $aa.Name -userLinkVoiceAppId $aa.Identity
                
                }

            }

            elseif ($holidayCallFlowAction -eq "Announcement") {
                
                $voiceMenuOptionAnnouncementType = $MenuOption.Prompt.ActiveType

                $holidayCallFlowMenuOptionsAnnouncement = "$voiceMenuOptionAnnouncementType"

                if ($voiceMenuOptionAnnouncementType -eq "TextToSpeech" -and $ShowTTSGreetingText) {
    
                    $audioFileName = $null
        
                    $holidayCallFlowMenuOptionsTTSAnnouncementValueExport = $MenuOption.Prompt.TextToSpeechPrompt
                    $holidayCallFlowMenuOptionsTTSAnnouncementValue = Optimize-DisplayName -String $MenuOption.Prompt.TextToSpeechPrompt

                    if ($ExportTTSGreetings) {

                        $holidayCallFlowMenuOptionsTTSAnnouncementValueExport | Out-File "$FilePath\$($aaHolidayCallFlowId)_holidayCallFlowMenuOptionsAnnouncement$DtmfKey.txt"
        
                        $ttsGreetings += ("click holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey " + '"' + "$FilePath\$($aaHolidayCallFlowId)_holidayCallFlowMenuOptionsAnnouncement$DtmfKey.txt" + '"')
        
                    }    
    
                    if ($holidayCallFlowMenuOptionsTTSAnnouncementValue.Length -gt $truncateGreetings) {
        
                        $holidayCallFlowMenuOptionsTTSAnnouncementValue = $holidayCallFlowMenuOptionsTTSAnnouncementValue.Remove($holidayCallFlowMenuOptionsTTSAnnouncementValue.Length - ($holidayCallFlowMenuOptionsTTSAnnouncementValue.Length -$truncateGreetings)).TrimEnd() + "..."
                    
                    }
        
                    $holidayCallFlowMenuOptionsAnnouncement += " <br> ''$holidayCallFlowMenuOptionsTTSAnnouncementValue''"
                
                }
        
                elseif ($voiceMenuOptionAnnouncementType -eq "AudioFile" -and $ShowAudioFileName) {
        
                    $holidayCallFlowMenuOptionsTTSAnnouncementValue = $null
        
                    # Audio File Announcement Name
                    $audioFileName = Optimize-DisplayName -String ($MenuOption.Prompt.AudioFilePrompt.FileName)

                    if ($ExportAudioFiles) {

                        $content = Export-CsOnlineAudioFile -Identity $MenuOption.Prompt.AudioFilePrompt.Id -ApplicationId OrgAutoAttendant
                        [System.IO.File]::WriteAllBytes("$FilePath\$audioFileName", $content)
        
                        $audioFileNames += ("click holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey " + '"' + "$FilePath\$audioFileName" + '"')
        
                    }
    
        
                    if ($audioFileName.Length -gt $truncateGreetings) {
        
                        $audioFileNameExtension = ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[0] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[1] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[2] + ($audioFileName[($audioFileName.Length -4)..$audioFileName.Length])[3]
                        $audioFileName = $audioFileName.Remove($audioFileName.Length -($audioFileName.Length - $truncateGreetings)) + "... $audioFileNameExtension"
        
                    }
        
                    $holidayCallFlowMenuOptionsAnnouncement += " <br> $audioFileName"
        
        
                }

                $mdAutoAttendantholidayCallFlow = "$mdDtmfLink holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey>$holidayCallFlowAction <br> $holidayCallFlowMenuOptionsAnnouncement] ---> holidayCallFlowMenuOptionsGreeting$($aaHolidayCallFlowId)`n"

                $allMermaidNodes += @("holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey","holidayCallFlowMenuOptionsGreeting$($aaHolidayCallFlowId)")

                $holidayCallFlowVoicemailSystemGreeting = $null

            }

            else {

                # Switch through transfer target type and set variables accordingly
                switch ($holidayCallFlowTargetType) {
                    User { 
                        $holidayCallFlowTargetTypeFriendly = "User"
                        $holidayCallFlowTargetUser = (Get-MgUser -UserId $($MenuOption.CallTarget.Id))
                        $holidayCallFlowTargetName = Optimize-DisplayName -String $holidayCallFlowTargetUser.DisplayName
                        $holidayCallFlowTargetIdentity = $holidayCallFlowTargetUser.Id

                        if ($FindUserLinks -eq $true) {

                            if ($DtmfKey) {

                                $DtmfOption = " (DTMF Option: $DtmfKey)"

                            }

                            else {

                                $DtmfOption =$null

                            }
         
                            . New-VoiceAppUserLinkProperties -userLinkUserId $($MenuOption.CallTarget.Id) -userLinkUserName $holidayCallFlowTargetUser.DisplayName -userLinkVoiceAppType "Auto Attendant" -userLinkVoiceAppActionType "HolidayCallFlowTransferCallToTarget$DtmfOption" -userLinkVoiceAppName $aa.Name -userLinkVoiceAppId $aa.Identity
                        
                        }                

                        if ($nestedVoiceApps -notcontains $holidayCallFlowTargetUser.Id) {

                            $nestedVoiceApps += $holidayCallFlowTargetUser.Id

                        }

                        $holidayCallFlowVoicemailSystemGreeting = $null

                        if ($ShowTTSGreetingText) {

                            $holidayCallFlowTransferGreetingValue = (. Get-IvrTransferMessage)[0]
                            $holidayCallFlowTransferGreetingValueExport = (. Get-IvrTransferMessage)[1]
        
                            if ($ExportTTSGreetings) {
        
                                $holidayCallFlowTransferGreetingValueExport | Out-File "$FilePath\$($aaHolidayCallFlowId)_autoAttendantholidayCallFlowTransferUserGreeting.txt"
                                $ttsGreetings += ("click holidayCallFlowTransferUserGreeting$($aaHolidayCallFlowId)-$holidayCallFlowUserOptionCounter " + '"' + "$FilePath\$($aaHolidayCallFlowId)_autoAttendantholidayCallFlowTransferUserGreeting.txt" + '"')
        
                            }
        
                            if ($holidayCallFlowTransferGreetingValue.Length -gt $truncateGreetings) {
        
                                $holidayCallFlowTransferGreetingValue = $holidayCallFlowTransferGreetingValue.Remove($holidayCallFlowTransferGreetingValue.Length - ($holidayCallFlowTransferGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                    
                            }
        
                            $holidayCallFlowTransferGreetingValue = "holidayCallFlowTransferUserGreeting$($aaHolidayCallFlowId)-$holidayCallFlowUserOptionCounter>Greeting<br>Transfer Message<br>''$holidayCallFlowTransferGreetingValue''] -->"
            
                        }
        
                        else {
        
                            $holidayCallFlowTransferGreetingValueExport = $null
                            $holidayCallFlowTransferGreetingValue = "holidayCallFlowTransferUserGreeting$($aaHolidayCallFlowId)-$holidayCallFlowUserOptionCounter>Greeting<br>Transfer Message] -->"
        
                        }
        
                        $allMermaidNodes += "holidayCallFlowTransferUserGreeting$($aaHolidayCallFlowId)-$holidayCallFlowUserOptionCounter"

                        $holidayCallFlowUserOptionCounter ++
        

                    }
                    ExternalPstn { 
                        $holidayCallFlowTargetTypeFriendly = "External Number"
                        $holidayCallFlowTargetName = ($MenuOption.CallTarget.Id).Replace("tel:","")
                        $holidayCallFlowTargetIdentity = $holidayCallFlowTargetName

                        if ($ObfuscatePhoneNumbers -eq $true) {

                            $holidayCallFlowTargetName = $holidayCallFlowTargetName.Remove(($holidayCallFlowTargetName.Length -4)) + "****"
        
                        }        

                        $holidayCallFlowVoicemailSystemGreeting = $null

                        if ($ShowTTSGreetingText) {

                            $holidayCallFlowTransferGreetingValue = (. Get-IvrTransferMessage)[0]
                            $holidayCallFlowTransferGreetingValueExport = (. Get-IvrTransferMessage)[1]
        
                            if ($ExportTTSGreetings) {
        
                                $holidayCallFlowTransferGreetingValueExport | Out-File "$FilePath\$($aaHolidayCallFlowId)_autoAttendantholidayCallFlowTransferPSTNGreeting.txt"
                                $ttsGreetings += ("click holidayCallFlowTransferPSTNGreeting$($aaHolidayCallFlowId)-$holidayCallFlowPSTNOptionCounter " + '"' + "$FilePath\$($aaHolidayCallFlowId)_autoAttendantholidayCallFlowTransferPSTNGreeting.txt" + '"')
        
                            }
        
                            if ($holidayCallFlowTransferGreetingValue.Length -gt $truncateGreetings) {
        
                                $holidayCallFlowTransferGreetingValue = $holidayCallFlowTransferGreetingValue.Remove($holidayCallFlowTransferGreetingValue.Length - ($holidayCallFlowTransferGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                    
                            }
        
                            $holidayCallFlowTransferGreetingValue = "holidayCallFlowTransferPSTNGreeting$($aaHolidayCallFlowId)-$holidayCallFlowPSTNOptionCounter>Greeting<br>Transfer Message<br>''$holidayCallFlowTransferGreetingValue''] -->"
            
                        }
        
                        else {
        
                            $holidayCallFlowTransferGreetingValueExport = $null
                            $holidayCallFlowTransferGreetingValue = "holidayCallFlowTransferPSTNGreeting$($aaHolidayCallFlowId)-$holidayCallFlowPSTNOptionCounter>Greeting<br>Transfer Message] -->"
        
                        }
        
                        $allMermaidNodes += "holidayCallFlowTransferPSTNGreeting$($aaHolidayCallFlowId)-$holidayCallFlowPSTNOptionCounter"

                        $holidayCallFlowPSTNOptionCounter ++

                    }
                    ApplicationEndpoint {

                        # Check if application endpoint is auto attendant or call queue

                        $matchingApplicationInstanceCheckAa = $allResourceAccounts | Where-Object {$_.ObjectId -eq $MenuOption.CallTarget.Id -and $_.ApplicationId -eq $applicationIdAa}

                        if ($matchingApplicationInstanceCheckAa) {

                            $MatchingAaHolidayCallFlowAa = $allAutoAttendants | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id}

                            $holidayCallFlowTargetTypeFriendly = "[Auto Attendant"
                            $holidayCallFlowTargetName = (Optimize-DisplayName -String $($MatchingAaHolidayCallFlowAa.Name)) + "]"

                        }

                        else {

                            $MatchingCqAaHolidayCallFlow = $allCallQueues | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id}

                            $holidayCallFlowTargetTypeFriendly = "[Call Queue"
                            $holidayCallFlowTargetName = (Optimize-DisplayName -String $($MatchingCqAaHolidayCallFlow.Name)) + "]"

                        }

                        $holidayCallFlowVoicemailSystemGreeting = $null

                    }
                    SharedVoicemail {

                        $holidayCallFlowTargetTypeFriendly = "Shared Voicemail"
                        $holidayCallFlowTargetGroup = (Get-MgGroup -GroupId $MenuOption.CallTarget.Id)
                        $holidayCallFlowTargetName = Optimize-DisplayName -String $holidayCallFlowTargetGroup.DisplayName
                        $holidayCallFlowTargetIdentity = $holidayCallFlowTargetGroup.Id

                        if ($MenuOption.CallTarget.EnableSharedVoicemailSystemPromptSuppression -eq $false) {
                            
                            $holidayCallFlowVoicemailSystemGreeting = "holidayCallFlowSystemGreeting$($aaHolidayCallFlowId)-$holidayCallFlowSharedVoicemailCounter>Greeting <br> MS System Message] -->"

                            $holidayCallFlowVoicemailSystemGreetingValue = (. Get-MsSystemMessage)[-1]
                            $holidayCallFlowVoicemailSystemGreetingValueExport = (. Get-MsSystemMessage)[0]

                            if ($ShowTTSGreetingText) {
            
                                if ($ExportTTSGreetings) {
            
                                    $holidayCallFlowVoicemailSystemGreetingValueExport | Out-File "$FilePath\$($aaHolidayCallFlowId)_autoAttendantholidayCallFlowMsSystemMessage.txt"
                    
                                    $ttsGreetings += ("click holidayCallFlowSystemGreeting$($aaHolidayCallFlowId)-$holidayCallFlowSharedVoicemailCounter " + '"' + "$FilePath\$($aaHolidayCallFlowId)_autoAttendantholidayCallFlowMsSystemMessage.txt" + '"')
                    
                                }    
                
            
                                if ($holidayCallFlowVoicemailSystemGreetingValue.Length -gt $truncateGreetings) {
            
                                    $holidayCallFlowVoicemailSystemGreetingValue = $holidayCallFlowVoicemailSystemGreetingValue.Remove($holidayCallFlowVoicemailSystemGreetingValue.Length - ($holidayCallFlowVoicemailSystemGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
            
                                }
            
                                $holidayCallFlowVoicemailSystemGreeting = $holidayCallFlowVoicemailSystemGreeting.Replace("]"," <br> ''$holidayCallFlowVoicemailSystemGreetingValue'']")
            
                            }
            
                            $holidayCallFlowTransferGreetingValue = $null

                        }

                        else {
                            
                            $holidayCallFlowVoicemailSystemGreeting = $null

                        }

                        if ($ShowSharedVoicemailGroupMembers -eq $true) {

                            . Get-SharedVoicemailGroupMembers -SharedVoicemailGroupId $MenuOption.CallTarget.Id

                            $holidayCallFlowTargetName = "$holidayCallFlowTargetName$mdSharedVoicemailGroupMembers"

                        }


                        if ($ShowTTSGreetingText) {

                            $holidayCallFlowTransferGreetingValue = (. Get-IvrTransferMessage)[0]
                            $holidayCallFlowTransferGreetingValueExport = (. Get-IvrTransferMessage)[1]
        
                            if ($ExportTTSGreetings) {
        
                                $holidayCallFlowTransferGreetingValueExport | Out-File "$FilePath\$($aaHolidayCallFlowId)_autoAttendantholidayCallFlowTransferSharedVoicemailGreeting.txt"
                                $ttsGreetings += ("click holidayCallFlowTransferSharedVoicemailGreeting$($aaHolidayCallFlowId)-$holidayCallFlowSharedVoicemailCounter " + '"' + "$FilePath\$($aaHolidayCallFlowId)_autoAttendantholidayCallFlowTransferSharedVoicemailGreeting.txt" + '"')
        
                            }
        
                            if ($holidayCallFlowTransferGreetingValue.Length -gt $truncateGreetings) {
        
                                $holidayCallFlowTransferGreetingValue = $holidayCallFlowTransferGreetingValue.Remove($holidayCallFlowTransferGreetingValue.Length - ($holidayCallFlowTransferGreetingValue.Length -$truncateGreetings)).TrimEnd() + "..."
                    
                            }
        
                            $holidayCallFlowTransferGreetingValue = "holidayCallFlowTransferSharedVoicemailGreeting$($aaHolidayCallFlowId)-$holidayCallFlowSharedVoicemailCounter>Greeting<br>Transfer Message<br>''$holidayCallFlowTransferGreetingValue''] -->"
            
                        }
        
                        else {
        
                            $holidayCallFlowTransferGreetingValueExport = $null
                            $holidayCallFlowTransferGreetingValue = "holidayCallFlowTransferSharedVoicemailGreeting$($aaHolidayCallFlowId)-$holidayCallFlowSharedVoicemailCounter>Greeting<br>Transfer Message] -->"
        
                        }
        
                        $allMermaidNodes += "holidayCallFlowTransferSharedVoicemailGreeting$($aaHolidayCallFlowId)-$holidayCallFlowSharedVoicemailCounter"

                        $allMermaidNodes += "holidayCallFlowSystemGreeting$($aaHolidayCallFlowId)-$holidayCallFlowSharedVoicemailCounter"

                        $holidayCallFlowSharedVoicemailCounter ++

                    }
                }

                # Check if transfer target type is call queue
                if ($holidayCallFlowTargetTypeFriendly -eq "[Call Queue") {

                    $MatchingCQIdentity = ($allCallQueues | Where-Object {$_.ApplicationInstances -contains $MenuOption.CallTarget.Id}).Identity

                    $mdAutoAttendantHolidayCallFlow = "$mdDtmfLink holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey($holidayCallFlowAction) --> $($MatchingCQIdentity)($holidayCallFlowTargetTypeFriendly <br> $holidayCallFlowTargetName)`n"
                    
                    if ($nestedVoiceApps -notcontains $MatchingCQIdentity) {

                        $nestedVoiceApps += $MatchingCQIdentity

                    }

                    $allMermaidNodes += @("holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey","$($MatchingCQIdentity)")

                
                } # End if transfer target type is call queue

                elseif ($holidayCallFlowTargetTypeFriendly -eq "[Auto Attendant") {

                    $mdAutoAttendantHolidayCallFlow = "$mdDtmfLink holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey($holidayCallFlowAction) --> $($MatchingAaHolidayCallFlowAa.Identity)($holidayCallFlowTargetTypeFriendly <br> $holidayCallFlowTargetName)`n"

                    if ($nestedVoiceApps -notcontains $MatchingAaHolidayCallFlowAa.Identity) {

                        $nestedVoiceApps += $MatchingAaHolidayCallFlowAa.Identity

                    }

                    $allMermaidNodes += @("holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey","$($MatchingAaHolidayCallFlowAa.Identity)")

                }

                # Check if holiday callflow action target is trasnfer call to target but something other than call queue
                else {

                    $mdAutoAttendantHolidayCallFlow = $mdDtmfLink + "$holidayCallFlowTransferGreetingValue holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey($holidayCallFlowAction) --> $holidayCallFlowVoicemailSystemGreeting $($holidayCallFlowTargetIdentity)($holidayCallFlowTargetTypeFriendly <br> $holidayCallFlowTargetName)`n"

                    $allMermaidNodes += @("holidayCallFlow$($aaHolidayCallFlowId)$DtmfKey","$($holidayCallFlowTargetIdentity)")

                }

            }

            $mdAutoAttendantHolidayCallFlowMenuOptions += $mdAutoAttendantHolidayCallFlow

        }

        # Greeting can only exist once. Add greeting before call flow and set mdAutoAttendantHolidayCallFlow to the new variable.

            $mdHolidayCallFlowGreeting += $mdAutoAttendantHolidayCallFlowMenuOptions
            $mdAutoAttendantHolidayCallFlow = $mdHolidayCallFlowGreeting
    
    }

    # Remove Greeting node, if none is configured
    if ($holidayCallFlowGreeting -eq "Greeting <br> None") {

        $mdAutoAttendantHolidayCallFlow = ($mdAutoAttendantHolidayCallFlow.Replace("holidayCallFlowGreeting$($aaHolidayCallFlowId)>$holidayCallFlowGreeting] --> ","")).TrimStart()

    }

    $mdAutoAttendantHolidayCallFlow = "elementAAHolidayIvr$($aaObjectId)-$($HolidayCounter){IVR<br>$holidayCallHandlingName} --> " + $mdAutoAttendantHolidayCallFlow
    
}