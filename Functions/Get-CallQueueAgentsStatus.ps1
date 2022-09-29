[CmdletBinding(DefaultParametersetName="None")]
param(
    [Parameter(Mandatory=$false)][ValidateSet("WorkingDir","CustomDir",$null)][String]$Export
)

function Get-CallQueueAgentsStatus {
    param (
        [Parameter(Mandatory=$false)][String]$CallQueueId,
        [Parameter(Mandatory=$false)][ValidateSet("WorkingDir","CustomDir",$null)][String]$Export
    )

    # Import Function to Connect to Teams and Graph
    . .\Functions\Connect-M365CFV.ps1

    . Connect-M365CFV

    if (!$CallQueueId) {

        $CallQueueId = (Get-CsCallQueue -WarningAction SilentlyContinue | Select-Object Name, Identity | Out-GridView -Title "Choose a Call Queue from the list..." -PassThru).Identity

    }

    $callQueue = Get-CsCallQueue -WarningAction SilentlyContinue -Identity $CallQueueId

    # Check if call queue useses users, group or teams channel as distribution list
    if (!$callQueue.DistributionLists) {

        $CqAgentListType = "Direct"

    }

    else {

        if (!$callQueue.ChannelId) {

            $CqAgentListType = "Groups"

            $allGroups = @()

            foreach ($DistributionList in $callQueue.DistributionLists.Guid) {

                $groupDetails = New-Object -TypeName psobject

                $DistributionListGroup = Get-MgGroup -GroupId $DistributionList

                $groupDetails | Add-Member -MemberType NoteProperty -Name "GroupId" -Value $DistributionListGroup.Id
                $groupDetails | Add-Member -MemberType NoteProperty -Name "GroupName" -Value $DistributionListGroup.DisplayName
                $groupDetails | Add-Member -MemberType NoteProperty -Name "GroupMembers" -Value (Get-MgGroupMember -GroupId $DistributionList).Id

                $allGroups += $groupDetails

                $DistributionLists += $DistributionList

            }

        }

        else {

            $TeamName = (Get-Team -GroupId $callQueue.DistributionLists.Guid).DisplayName
            $ChannelName = (Get-TeamChannel -GroupId $callQueue.DistributionLists.Guid | Where-Object {$_.Id -eq $callQueue.ChannelId}).DisplayName

            $CqAgentListType = "Channel"

        }

    }


    $callQueueAgents = @()

    foreach ($agent in $callQueue.Agents) {

        $agentProperties = New-Object -TypeName psobject

        $agentTeamsUser = Get-CsOnlineUser -Identity $agent.ObjectId

        $agentProperties | Add-Member -MemberType NoteProperty -Name "User Principal Name" -Value $agentTeamsUser.UserPrincipalName
        $agentProperties | Add-Member -MemberType NoteProperty -Name "Display Name" -Value $agentTeamsUser.DisplayName

        if ($agentTeamsUser.LineUri) {
            $agentProperties | Add-Member -MemberType NoteProperty -Name "Phone Number" -Value $agentTeamsUser.LineUri.Replace("tel:","")
        }
        else {
            $agentProperties | Add-Member -MemberType NoteProperty -Name "Phone Number" -Value "No Number Assigned"
        }
        $agentProperties | Add-Member -MemberType NoteProperty -Name "Opt In Status" -Value $agent.OptIn
        $agentProperties | Add-Member -MemberType NoteProperty -Name "Queue Name" -Value $callQueue.Name
        
        switch ($cqAgentListType) {
            Channel {

                $agentProperties | Add-Member -MemberType NoteProperty -Name "Assignment" -Value "Team: $TeamName; Channel: $ChannelName"

            }
            Groups {

                $assignments = ($allGroups | Where-Object {$_.GroupMembers -contains $agent.ObjectId}).GroupName

                if ($assignments) {

                    $groupMember = ""

                    $assignmentCount = $assignments.Count -1

                    foreach ($assignment in $assignments) {

                        if ($assignments.IndexOf($assignment) -eq $assignmentCount) {
                            
                            $groupMember += "$assignment"

                        }

                        else {

                            $groupMember += "$assignment;"

                        }
                    }

                    $agentProperties | Add-Member -MemberType NoteProperty -Name "Assignment" -Value "Groups: $groupMember"

                }

                else {

                    $agentProperties | Add-Member -MemberType NoteProperty -Name "Assignment" -Value "Direct"

                }

            }
            Direct {

                $agentProperties | Add-Member -MemberType NoteProperty -Name "Assignment" -Value "Direct"

            }
            Default {}
        }

        $callQueueAgents += $agentProperties

    }

    if ($Export -eq "WorkingDir") {

        $callQueueAgents | Export-Csv -Path ".\$($callQueue.Name)_Agents.csv" -Delimiter ";" -NoTypeInformation

    }

    if ($Export -eq "CustomDir") {
        
        $defaultPath = Get-Location
        Add-Type -AssemblyName System.Windows.Forms
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            selectedPath = $defaultPath
        }
        [void]$folderBrowser.ShowDialog()
        $exportPath = $folderBrowser.SelectedPath

        $callQueueAgents | Export-Csv -Path "$exportPath\$($callQueue.Name)_Agents.csv" -Delimiter ";" -NoTypeInformation

    }

    return $callQueueAgents | Out-GridView -Title "$($callQueue.Name) - Agents Status"
    
}

. Get-CallQueueAgentsStatus -Export $Export