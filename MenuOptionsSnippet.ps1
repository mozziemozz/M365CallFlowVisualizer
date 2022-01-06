$defaultCallFlowMenuOptions = $aa.DefaultCallFlow.Menu.MenuOptions

# Get the current auto attentans default call flow greeting
if (!$defaultCallFlow.Greetings.ActiveType.Value){
    $defaultCallFlowGreeting = "Greeting <br> None"
}

else {
    $defaultCallFlowGreeting = "Greeting <br> $($defaultCallFlow.Greetings.ActiveType.Value)"
}

if ($defaultCallFlowMenuOptions.Count -gt 1) {

    $DefaultCallFlowId = $aa.DefaultCallFlow.Id
    
    foreach ($MenuOption in $defaultCallFlowMenuOptions) {

        $defaultCallFlowMenuOptionAction = $MenuOption.Action.Value
        $defaultCallFlowMenuOptionDTMFResponse = $MenuOption.DtmfResponse.Value
        $defaultCallFlowMenuOptionVoiceResponse = $MenuOption.VoiceResponses
        $defaultCallFlowMenuOptionCallTargetIdentity = $MenuOption.CallTarget.Id
        $defaultCallFlowMenuOptionCallTargetIdentity = $MenuOption.CallTarget.Type.Value
        $defaultCallFlowMenuOptionGreeting = $MenuOption.Prompt.ActiveType.Value


    }

}