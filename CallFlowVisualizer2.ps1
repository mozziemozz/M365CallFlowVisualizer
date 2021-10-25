#Requires -Modules MsOnline, MicrosoftTeams

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)][ValidateSet("Markdown","Mermaid")][String]$docType = "Markdown",
    [Parameter(Mandatory=$false)][Int32]$ShowNestedDepth = 1,
    [Parameter(Mandatory=$false)][Switch]$SubSequentRun,
    [Parameter(Mandatory=$false)][string]$PhoneNumber

)

# From: https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/clearing-all-user-variables
function Get-UserVariable ($Name = '*') {
# these variables may exist in certain environments (like ISE, or after use of foreach)
$special = 'ps','psise','psunsupportedconsoleapplications', 'foreach', 'profile'

$ps = [PowerShell]::Create()
$null = $ps.AddScript('$null=$host;Get-Variable') 
$reserved = $ps.Invoke() | Select-Object -ExpandProperty Name
$ps.Runspace.Close()
$ps.Dispose()
Get-Variable -Scope Global | 
    Where-Object Name -like $Name |
    Where-Object { $reserved -notcontains $_.Name } |
    Where-Object { $special -notcontains $_.Name } |
    Where-Object Name 
}

if ($SubSequentRun) {
    Get-UserVariable | Remove-Variable
}

function Set-Mermaid {
    param (
        [Parameter(Mandatory=$true)][String]$docType
        )

    if ($docType -eq "Markdown") {
        $mdStart =@"
``````mermaid
flowchart TB
"@

        $mdEnd =@"

``````
"@

        $fileExtension = ".md"
    }

    else {
        $mdStart =@"
flowchart TB
"@

        $mdEnd =@"

"@

        $fileExtension = ".mmd"
    }

    $mermaidCode = @()

    $mermaidCode += $mdStart
    $mermaidCode += $mdIncomingCall
    $mermaidCode += $mdVoiceApp
    $mermaidCode += $mdNodeAdditionalNumbers
    $mermaidCode += $mdEnd
    
}

function Get-VoiceApp {
    param (
        [Parameter(Mandatory=$false)][String]$VoiceApp
        )

        if ($PhoneNumber) {
            $resourceAccount = Get-CsOnlineApplicationInstance | Where-Object {$_.PhoneNumber -match $PhoneNumber}
        }

        else {
            # Get resource account (it was a design choice to select a resource account instead of a voice app, people tend to know the phone number and want to know what happens when a particular number is called.)
            $resourceAccount = Get-CsOnlineApplicationInstance | Where-Object {$_.PhoneNumber -notlike ""} | Select-Object DisplayName, PhoneNumber, ObjectId, ApplicationId | Out-GridView -PassThru -Title "Choose an auto attendant or a call queue from the list:"

        }

        switch ($resourceAccount.ApplicationId) {
            # Application Id for auto attendants
            "ce933385-9390-45d1-9512-c8d228074e07" {
                $voiceAppType = "Auto Attendant"
                $voiceApp = Get-CsAutoAttendant | Where-Object {$_.ApplicationInstances -contains $resourceAccount.ObjectId}
            }
            # Application Id for call queues
            "11cd3e2e-fccb-42ad-ad00-878b93575e07" {
                $voiceAppType = "Call Queue"
                $voiceApp = Get-CsCallQueue | Where-Object {$_.ApplicationInstances -contains $resourceAccount.ObjectId}
            }
        }

        # Create ps object to store properties from voice app and resource account
        $voiceAppProperties = New-Object -TypeName psobject
        $voiceAppProperties | Add-Member -MemberType NoteProperty -Name "PhoneNumber" -Value $($resourceAccount.PhoneNumber).Replace("tel:","")
        $voiceAppProperties | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $VoiceApp.Name

        $resourceAccountCounter = 1

        $mdIncomingCall = "start$($resourceAccountCounter)((Incoming Call at <br> $($voiceAppProperties.PhoneNumber))) --> "
        $mdVoiceApp = "voiceApp([$($voiceAppType) <br> $($voiceAppProperties.DisplayName)])"

        $mdNodeAdditionalNumbers = @()

        foreach ($ApplicationInstance in ($VoiceApp.ApplicationInstances | Where-Object {$_ -notcontains $resourceAccount.ObjectId})) {

            $resourceAccountCounter ++

            $additionalResourceAccount = ((Get-CsOnlineApplicationInstance -Identity $ApplicationInstance).PhoneNumber) -replace ("tel:","")

            $mdNodeAdditionalNumber = "start$($resourceAccountCounter)((Incoming Call at <br> $additionalResourceAccount)) -.-> voiceApp"

            $mdNodeAdditionalNumbers += $mdNodeAdditionalNumber

        }

        

}

. Get-VoiceApp
. Set-Mermaid -docType $docType



