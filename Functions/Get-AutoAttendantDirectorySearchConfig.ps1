<#
    .SYNOPSIS
    Reads the directory search configuration of an auto attendant.
    
    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Changelog:          .\Changelog.md

#>

function Get-AutoAttendantDirectorySearchConfig {
    param (
        [Parameter(Mandatory=$true)][String]$CalLFlowType
    )

    switch ($CalLFlowType) {
        defaultCallFlow {

            $aaCurrentCallFlow = $defaultCallFlow
        
        }
        afterHoursCallFlow {

            $aaCurrentCallFlow = $afterHoursCallFlow
        
        }
        holidayCallFlow {

            $aaCurrentCallFlow = $holidayCallFlow

        }
        Default {}
    }
    
    if ($aaCurrentCallFlow.Menu.DirectorySearchMethod -ne "None") {

        # No inclusions and no exclusions
        if (!$aa.DirectoryLookupScope.InclusionScope -and !$aa.DirectoryLookupScope.ExclusionScope) {

            $directoryLookupScope = "<br>Included Users: All Online Users<br>Excluded Groups: None"

        }

        # Inclusions but no exclusions
        elseif ($aa.DirectoryLookupScope.InclusionScope -and !$aa.DirectoryLookupScope.ExclusionScope) {

            $directoryLookupScope = "<br>Included Users: Specific Groups:"

            $includedGroups = $aa.DirectoryLookupScope.InclusionScope.GroupScope.GroupIds

            foreach ($includedGroup in $includedGroups) {

                $groupDisplayName = Optimize-DisplayName -String ((Get-MgGroup -GroupId $includedGroup).DisplayName)

                $directoryLookupScope = $directoryLookupScope + "<br>- $groupDisplayName"

            }

            $directoryLookupScope = $directoryLookupScope + "<br>Excluded Groups: None"

        }

        # Inclusions and exclusions
        elseif ($aa.DirectoryLookupScope.InclusionScope -and $aa.DirectoryLookupScope.ExclusionScope) {
           
            $directoryLookupScope = "<br>Included Users: Specific Groups:"

            $includedGroups = $aa.DirectoryLookupScope.InclusionScope.GroupScope.GroupIds

            foreach ($includedGroup in $includedGroups) {

                $groupDisplayName = Optimize-DisplayName -String ((Get-MgGroup -GroupId $includedGroup).DisplayName)

                $directoryLookupScope = $directoryLookupScope + "<br>- $groupDisplayName"

            }

            $directoryLookupScope = $directoryLookupScope + "<br>Excluded Groups: Specific Groups:"

            $excludedGroups = $aa.DirectoryLookupScope.ExclusionScope.GroupScope.GroupIds

            foreach ($excludedGroup in $excludedGroups) {

                $groupDisplayName = Optimize-DisplayName -String ((Get-MgGroup -GroupId $excludedGroup).DisplayName)

                $directoryLookupScope = $directoryLookupScope + "<br>- $groupDisplayName"

            }

        }

        # Only exclusions
        else {

            $directoryLookupScope = "<br>Included Users: All Online Users<br>Excluded Groups: Specific Groups:"

            $excludedGroups = $aa.DirectoryLookupScope.ExclusionScope.GroupScope.GroupIds

            foreach ($excludedGroup in $excludedGroups) {

                $groupDisplayName = Optimize-DisplayName -String ((Get-MgGroup -GroupId $excludedGroup).DisplayName)

                $directoryLookupScope = $directoryLookupScope + "<br>- $groupDisplayName"

            }

        }

        switch ($aaCurrentCallFlow.Menu.DirectorySearchMethod) {
            ByName {
                $currentCallFlowDirectorySearchType = "- By Name / Keypad Entry"
            }
            ByExtension {
                $currentCallFlowDirectorySearchType = "- By Extension / AAD Phone Ext."
            }
            Default {}
        }


        if ($aaIsVoiceResponseEnabled) {

            $currentCallFlowDirectorySearchType = $currentCallFlowDirectorySearchType + "<br>- Speech / Voice Input"

        }

        else {

            $currentCallFlowDirectorySearchType = $currentCallFlowDirectorySearchType

        }

        switch ($CalLFlowType) {
            defaultCallFlow {

                $mermaidCode += "defaultCallFlowMenuOptions$($aaDefaultCallFlowAaObjectId) -.-> defaultCallFlowDirectorySearch$($aaDefaultCallFlowAaObjectId)[(Directory Search Methods:<br>$currentCallFlowDirectorySearchType $directoryLookupScope)]"

                $allMermaidNodes += "defaultCallFlowDirectorySearch$($aaDefaultCallFlowAaObjectId)"
            
            }
            afterHoursCallFlow {

                $mermaidCode += "afterHoursCallFlowMenuOptions$($aaAfterHoursCallFlowAaObjectId) -.-> afterHoursCallFlowDirectorySearch$($aaAfterHoursCallFlowAaObjectId)[(Directory Search Methods:<br>$currentCallFlowDirectorySearchType $directoryLookupScope)]"

                $allMermaidNodes += "afterHoursCallFlowDirectorySearch$($aaAfterHoursCallFlowAaObjectId)"
            
            }
            holidayCallFlow {

                $mermaidCode += "holidayCallFlowMenuOptions$($aaHolidayCallFlowId) -.-> holidayCallFlowDirectorySearch$($aaHolidayCallFlowId)[(Directory Search Methods:<br>$currentCallFlowDirectorySearchType $directoryLookupScope)]"

                $allMermaidNodes += "holidayCallFlowDirectorySearch$($aaHolidayCallFlowId)"

            }
            Default {}
        }

    }

}