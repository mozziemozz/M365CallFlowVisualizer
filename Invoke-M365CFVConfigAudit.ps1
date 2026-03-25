<#
    .SYNOPSIS
    Audits Microsoft 365 Phone System configuration for common misconfigurations and best practice violations.

    .DESCRIPTION
    Author:             Patrik Lleshaj (https://github.com/realgarit)
    Version:            1.0.0
    Changelog:          .\Changelog.md

    Scans all Auto Attendants, Call Queues, and Resource Accounts in the connected tenant
    and checks for issues like missing agents, dead-end routes, orphaned resource accounts,
    missing phone numbers, disabled users in queues, and more.

    Outputs a health report to the console and optionally saves it as CSV and Markdown.

    .PARAMETER CustomFilePath
    Directory where the report files will be saved.

    .PARAMETER ExportCsv
    Export findings as a CSV file.

    .PARAMETER ExportMarkdown
    Export findings as a Markdown report.

    .PARAMETER SaveSnapshot
    Save a JSON snapshot of the current configuration for later diffing with Compare-M365CFVConfigSnapshot.ps1.

    .EXAMPLE
    .\Invoke-M365CFVConfigAudit.ps1

    .EXAMPLE
    .\Invoke-M365CFVConfigAudit.ps1 -ExportMarkdown -SaveSnapshot -CustomFilePath ".\Output\audit"
#>

#Requires -Modules @{ ModuleName = "MicrosoftTeams"; ModuleVersion = "5.0.0" }, "Microsoft.Graph.Users", "Microsoft.Graph.Groups"

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)][String]$CustomFilePath = ".\Output\audit\$(Get-Date -Format 'yyyy-MM-dd')",
    [Parameter(Mandatory = $false)][Switch]$ExportCsv,
    [Parameter(Mandatory = $false)][Switch]$ExportMarkdown,
    [Parameter(Mandatory = $false)][Switch]$SaveSnapshot
)

# Load shared functions
. .\Functions\Connect-M365CFV.ps1
. .\Functions\Get-AllVoiceAppsAndResourceAccounts.ps1
. .\Functions\Format-PhoneNumber.ps1

# Connect
. Connect-M365CFV

# Fetch all data
$CacheResults = $false
. Get-AllVoiceAppsAndResourceAccounts

# Build lookup tables
$allAutoAttendantIds = $allAutoAttendants.Identity
$allCallQueueIds = $allCallQueues.Identity
$resourceAccountLookup = @{}
foreach ($ra in $allResourceAccounts) {
    $resourceAccountLookup[$ra.ObjectId] = $ra
}

# All voice app application instance IDs (for orphan detection)
$allAssociatedInstanceIds = @()
$allAssociatedInstanceIds += $allAutoAttendants.ApplicationInstances
$allAssociatedInstanceIds += $allCallQueues.ApplicationInstances
$allAssociatedInstanceIds = $allAssociatedInstanceIds | Sort-Object -Unique

# Resource accounts with phone numbers
$resourceAccountsWithPhone = $allResourceAccounts | Where-Object { $_.PhoneNumber }

# Findings collection
$findings = [System.Collections.ArrayList]::new()

function Add-Finding {
    param(
        [string]$Severity, # Critical, Warning, Info
        [string]$Category,
        [string]$ResourceType,
        [string]$ResourceName,
        [string]$ResourceId,
        [string]$Description
    )

    [void]$findings.Add([PSCustomObject]@{
        Severity     = $Severity
        Category     = $Category
        ResourceType = $ResourceType
        ResourceName = $ResourceName
        ResourceId   = $ResourceId
        Description  = $Description
    })
}

Write-Host "`nRunning configuration audit..." -ForegroundColor Cyan

# ========================================
# AUTO ATTENDANT CHECKS
# ========================================

foreach ($aa in $allAutoAttendants) {

    $aaName = $aa.Name
    $aaId = $aa.Identity

    # Check: AA without any resource account
    if (-not $aa.ApplicationInstances -or $aa.ApplicationInstances.Count -eq 0) {

        # Check if it's nested somewhere
        $isNested = $false
        foreach ($otherAa in $allAutoAttendants) {
            if ($otherAa.Identity -eq $aaId) { continue }
            $callFlowJson = $otherAa | ConvertTo-Json -Depth 10
            if ($callFlowJson -match $aaId) {
                $isNested = $true
                break
            }
        }

        if (-not $isNested) {
            Add-Finding -Severity "Warning" -Category "Orphaned" -ResourceType "Auto Attendant" `
                -ResourceName $aaName -ResourceId $aaId `
                -Description "No resource account assigned and not nested in any other voice app. This AA is unreachable."
        }
    }

    # Check: AA has resource account but none have phone numbers (top-level check)
    if ($aa.ApplicationInstances -and $aa.ApplicationInstances.Count -gt 0) {
        $hasPhone = $false
        foreach ($instanceId in $aa.ApplicationInstances) {
            $ra = $resourceAccountLookup[$instanceId]
            if ($ra -and $ra.PhoneNumber) {
                $hasPhone = $true
                break
            }
        }

        # Only flag if it doesn't appear to be nested
        if (-not $hasPhone) {
            $isNested = $false
            foreach ($otherAa in $allAutoAttendants) {
                if ($otherAa.Identity -eq $aaId) { continue }
                $callFlowJson = $otherAa | ConvertTo-Json -Depth 10
                if ($callFlowJson -match $aaId) {
                    $isNested = $true
                    break
                }
            }
            foreach ($cq in $allCallQueues) {
                $callFlowJson = $cq | ConvertTo-Json -Depth 10
                if ($callFlowJson -match $aaId) {
                    $isNested = $true
                    break
                }
            }

            if (-not $isNested) {
                Add-Finding -Severity "Critical" -Category "No Phone Number" -ResourceType "Auto Attendant" `
                    -ResourceName $aaName -ResourceId $aaId `
                    -Description "Top-level AA has resource account(s) but none have a phone number assigned. Callers can't reach this AA externally."
            }
        }
    }

    # Check: Default call flow disconnect without greeting
    $defaultMenu = $aa.DefaultCallFlow.Menu
    if ($defaultMenu) {
        $allDisconnect = $true
        foreach ($option in $defaultMenu.MenuOptions) {
            if ($option.Action -ne "DisconnectCall") {
                $allDisconnect = $false
                break
            }
        }

        if ($allDisconnect -and (-not $aa.DefaultCallFlow.Greetings -or $aa.DefaultCallFlow.Greetings.Count -eq 0)) {
            Add-Finding -Severity "Warning" -Category "Dead End" -ResourceType "Auto Attendant" `
                -ResourceName $aaName -ResourceId $aaId `
                -Description "Default call flow disconnects callers without any greeting. Callers hear nothing before being disconnected."
        }
    }

    # Check: No after-hours handling configured
    $hasAfterHours = $false
    if ($aa.CallHandlingAssociations) {
        foreach ($assoc in $aa.CallHandlingAssociations) {
            if ($assoc.Type -eq "AfterHours") {
                $hasAfterHours = $true
                break
            }
        }
    }

    if (-not $hasAfterHours) {
        Add-Finding -Severity "Info" -Category "No After-Hours" -ResourceType "Auto Attendant" `
            -ResourceName $aaName -ResourceId $aaId `
            -Description "No after-hours call handling configured. This AA uses the same call flow 24/7."
    }

    # Check: No holiday handling configured
    $hasHolidays = $false
    if ($aa.CallHandlingAssociations) {
        foreach ($assoc in $aa.CallHandlingAssociations) {
            if ($assoc.Type -eq "Holiday") {
                $hasHolidays = $true
                break
            }
        }
    }

    if (-not $hasHolidays) {
        Add-Finding -Severity "Info" -Category "No Holidays" -ResourceType "Auto Attendant" `
            -ResourceName $aaName -ResourceId $aaId `
            -Description "No holiday call handling configured."
    }

    # Check: Menu options that reference non-existent voice apps
    if ($defaultMenu -and $defaultMenu.MenuOptions) {
        foreach ($option in $defaultMenu.MenuOptions) {
            if ($option.CallTarget -and $option.CallTarget.Id) {
                $targetId = $option.CallTarget.Id

                if ($option.CallTarget.Type -eq "ConfigurationEndpoint" -or $option.CallTarget.Type -eq "ApplicationEndpoint") {
                    if ($allAutoAttendantIds -notcontains $targetId -and $allCallQueueIds -notcontains $targetId) {
                        # Could be a resource account ID - check
                        $matchedRa = $resourceAccountLookup[$targetId]
                        if (-not $matchedRa) {
                            Add-Finding -Severity "Critical" -Category "Broken Reference" -ResourceType "Auto Attendant" `
                                -ResourceName $aaName -ResourceId $aaId `
                                -Description "Menu option references target '$targetId' which doesn't exist as a voice app or resource account."
                        }
                    }
                }
            }
        }
    }
}

# ========================================
# CALL QUEUE CHECKS
# ========================================

foreach ($cq in $allCallQueues) {

    $cqName = $cq.Name
    $cqId = $cq.Identity

    # Check: CQ without any resource account
    if (-not $cq.ApplicationInstances -or $cq.ApplicationInstances.Count -eq 0) {
        $isNested = $false
        foreach ($aa in $allAutoAttendants) {
            $callFlowJson = $aa | ConvertTo-Json -Depth 10
            if ($callFlowJson -match $cqId) {
                $isNested = $true
                break
            }
        }

        if (-not $isNested) {
            Add-Finding -Severity "Warning" -Category "Orphaned" -ResourceType "Call Queue" `
                -ResourceName $cqName -ResourceId $cqId `
                -Description "No resource account assigned and not nested in any other voice app. This CQ is unreachable."
        }
    }

    # Check: CQ with no agents at all
    $agentCount = 0
    if ($cq.Agents) { $agentCount += $cq.Agents.Count }
    if ($cq.DistributionLists) { $agentCount += $cq.DistributionLists.Count }
    if ($cq.ChannelId) { $agentCount += 1 }

    if ($agentCount -eq 0) {
        Add-Finding -Severity "Critical" -Category "No Agents" -ResourceType "Call Queue" `
            -ResourceName $cqName -ResourceId $cqId `
            -Description "Call queue has no agents, distribution lists, or channels assigned. Calls will immediately hit the timeout/overflow action."
    }

    # Check: All individual agents opted out
    if ($cq.Agents -and $cq.Agents.Count -gt 0 -and $cq.AllowOptOut -eq $true) {
        $optedInCount = ($cq.Agents | Where-Object { $_.OptIn -eq $true }).Count
        if ($optedInCount -eq 0) {
            Add-Finding -Severity "Critical" -Category "All Agents Opted Out" -ResourceType "Call Queue" `
                -ResourceName $cqName -ResourceId $cqId `
                -Description "All $($cq.Agents.Count) agents have opted out. No one is receiving calls."
        }
        elseif ($optedInCount -eq 1) {
            Add-Finding -Severity "Warning" -Category "Low Agent Count" -ResourceType "Call Queue" `
                -ResourceName $cqName -ResourceId $cqId `
                -Description "Only 1 out of $($cq.Agents.Count) agents is opted in."
        }
    }

    # Check: Timeout action is Disconnect
    if ($cq.TimeoutAction -eq "Disconnect") {
        Add-Finding -Severity "Warning" -Category "Disconnect on Timeout" -ResourceType "Call Queue" `
            -ResourceName $cqName -ResourceId $cqId `
            -Description "Timeout action is set to Disconnect. Callers who wait longer than $($cq.TimeoutThreshold)s will be dropped without transfer or voicemail."
    }

    # Check: Overflow action is Disconnect
    if ($cq.OverflowAction -eq "Disconnect") {
        Add-Finding -Severity "Warning" -Category "Disconnect on Overflow" -ResourceType "Call Queue" `
            -ResourceName $cqName -ResourceId $cqId `
            -Description "Overflow action is set to Disconnect. Once $($cq.OverflowThreshold) callers are waiting, new callers are dropped."
    }

    # Check: Very short timeout
    if ($cq.TimeoutThreshold -and $cq.TimeoutThreshold -le 30 -and $cq.TimeoutThreshold -gt 0) {
        Add-Finding -Severity "Warning" -Category "Short Timeout" -ResourceType "Call Queue" `
            -ResourceName $cqName -ResourceId $cqId `
            -Description "Timeout threshold is only $($cq.TimeoutThreshold) seconds. Callers may be routed away before an agent can pick up."
    }

    # Check: Presence-based routing off with Longest Idle
    if ($cq.RoutingMethod -eq "LongestIdle" -and $cq.PresenceBasedRouting -eq $false) {
        Add-Finding -Severity "Info" -Category "Routing Config" -ResourceType "Call Queue" `
            -ResourceName $cqName -ResourceId $cqId `
            -Description "Longest Idle routing normally enables presence-based routing automatically. Current state shows PresenceBasedRouting=false, which is unusual."
    }

    # Check: Conference mode disabled
    if ($cq.ConferenceMode -eq $false) {
        Add-Finding -Severity "Info" -Category "Performance" -ResourceType "Call Queue" `
            -ResourceName $cqName -ResourceId $cqId `
            -Description "Conference mode is disabled. Enabling it reduces the time callers wait to connect to an agent after the agent accepts."
    }
}

# ========================================
# RESOURCE ACCOUNT CHECKS
# ========================================

foreach ($ra in $allResourceAccounts) {

    $raName = $ra.DisplayName
    $raId = $ra.ObjectId

    # Check: Resource account not associated with any voice app
    if ($allAssociatedInstanceIds -notcontains $raId) {
        Add-Finding -Severity "Warning" -Category "Unassigned" -ResourceType "Resource Account" `
            -ResourceName "$raName ($($ra.UserPrincipalName))" -ResourceId $raId `
            -Description "Resource account is not assigned to any Auto Attendant or Call Queue."
    }

    # Check: Resource account has phone number but is not associated
    if ($ra.PhoneNumber -and $allAssociatedInstanceIds -notcontains $raId) {
        $phoneNumber = Format-PhoneNumber -PhoneNumberId $ra.PhoneNumber
        Add-Finding -Severity "Warning" -Category "Wasted Phone Number" -ResourceType "Resource Account" `
            -ResourceName "$raName ($($ra.UserPrincipalName))" -ResourceId $raId `
            -Description "Has phone number $phoneNumber assigned but is not linked to any voice app. This number is consuming a license but doing nothing."
    }
}

# ========================================
# CROSS-REFERENCE CHECKS
# ========================================

# Check: Multiple AAs/CQs sharing the same resource account
$instanceToVoiceApp = @{}
foreach ($aa in $allAutoAttendants) {
    foreach ($instanceId in $aa.ApplicationInstances) {
        if (-not $instanceToVoiceApp.ContainsKey($instanceId)) {
            $instanceToVoiceApp[$instanceId] = @()
        }
        $instanceToVoiceApp[$instanceId] += "AA: $($aa.Name)"
    }
}
foreach ($cq in $allCallQueues) {
    foreach ($instanceId in $cq.ApplicationInstances) {
        if (-not $instanceToVoiceApp.ContainsKey($instanceId)) {
            $instanceToVoiceApp[$instanceId] = @()
        }
        $instanceToVoiceApp[$instanceId] += "CQ: $($cq.Name)"
    }
}

foreach ($instanceId in $instanceToVoiceApp.Keys) {
    $apps = $instanceToVoiceApp[$instanceId]
    if ($apps.Count -gt 1) {
        $ra = $resourceAccountLookup[$instanceId]
        $raDisplayName = if ($ra) { $ra.DisplayName } else { $instanceId }
        Add-Finding -Severity "Warning" -Category "Shared Resource Account" -ResourceType "Resource Account" `
            -ResourceName $raDisplayName -ResourceId $instanceId `
            -Description "Shared across multiple voice apps: $($apps -join ', '). This can cause unexpected routing."
    }
}

# ========================================
# SUMMARY
# ========================================

$criticalCount = ($findings | Where-Object { $_.Severity -eq "Critical" }).Count
$warningCount = ($findings | Where-Object { $_.Severity -eq "Warning" }).Count
$infoCount = ($findings | Where-Object { $_.Severity -eq "Info" }).Count

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  CONFIGURATION AUDIT RESULTS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Tenant:           $((Get-CsTenant).DisplayName)"
Write-Host "  Auto Attendants:  $($allAutoAttendants.Count)"
Write-Host "  Call Queues:      $($allCallQueues.Count)"
Write-Host "  Resource Accounts: $($allResourceAccounts.Count)"
Write-Host ""

if ($criticalCount -gt 0) {
    Write-Host "  CRITICAL: $criticalCount" -ForegroundColor Red
}
else {
    Write-Host "  CRITICAL: 0" -ForegroundColor Green
}

if ($warningCount -gt 0) {
    Write-Host "  WARNING:  $warningCount" -ForegroundColor Yellow
}
else {
    Write-Host "  WARNING:  0" -ForegroundColor Green
}

Write-Host "  INFO:     $infoCount" -ForegroundColor Gray
Write-Host ""

# Health score: 100 minus penalties
$score = 100 - ($criticalCount * 15) - ($warningCount * 5) - ($infoCount * 1)
if ($score -lt 0) { $score = 0 }

$scoreColor = if ($score -ge 80) { "Green" } elseif ($score -ge 50) { "Yellow" } else { "Red" }
Write-Host "  Health Score: $score / 100" -ForegroundColor $scoreColor
Write-Host "========================================`n" -ForegroundColor Cyan

# Print detailed findings
if ($findings.Count -gt 0) {

    Write-Host "DETAILED FINDINGS:" -ForegroundColor Cyan
    Write-Host ""

    foreach ($severity in @("Critical", "Warning", "Info")) {

        $sevFindings = $findings | Where-Object { $_.Severity -eq $severity }

        if ($sevFindings.Count -eq 0) { continue }

        $sevColor = switch ($severity) {
            "Critical" { "Red" }
            "Warning" { "Yellow" }
            "Info" { "Gray" }
        }

        Write-Host "--- $($severity.ToUpper()) ---" -ForegroundColor $sevColor

        foreach ($f in $sevFindings) {
            Write-Host "  [$($f.ResourceType)] $($f.ResourceName)" -ForegroundColor $sevColor
            Write-Host "    Category: $($f.Category)"
            Write-Host "    $($f.Description)"
            Write-Host ""
        }
    }
}
else {
    Write-Host "No issues found. Your configuration looks clean." -ForegroundColor Green
}

# ========================================
# EXPORTS
# ========================================

if ($ExportCsv -or $ExportMarkdown -or $SaveSnapshot) {

    if (!(Test-Path -Path $CustomFilePath)) {
        New-Item -Path $CustomFilePath -ItemType Directory | Out-Null
    }
}

if ($ExportCsv) {
    $csvPath = "$CustomFilePath\ConfigAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $findings | Export-Csv -Path $csvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force
    Write-Host "CSV report saved to: $csvPath" -ForegroundColor Green
}

if ($ExportMarkdown) {
    $mdPath = "$CustomFilePath\ConfigAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').md"

    $mdContent = @"
# M365 Voice Configuration Audit Report

**Tenant:** $((Get-CsTenant).DisplayName)
**Date:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
**Health Score:** $score / 100

## Summary

| Metric | Count |
|--------|-------|
| Auto Attendants | $($allAutoAttendants.Count) |
| Call Queues | $($allCallQueues.Count) |
| Resource Accounts | $($allResourceAccounts.Count) |
| Critical Issues | $criticalCount |
| Warnings | $warningCount |
| Info | $infoCount |

## Findings

"@

    foreach ($severity in @("Critical", "Warning", "Info")) {
        $sevFindings = $findings | Where-Object { $_.Severity -eq $severity }
        if ($sevFindings.Count -eq 0) { continue }

        $emoji = switch ($severity) {
            "Critical" { "!!" }
            "Warning" { "!" }
            "Info" { "i" }
        }

        $mdContent += "`n### $severity ($($sevFindings.Count))`n`n"

        foreach ($f in $sevFindings) {
            $mdContent += "- **[$emoji] [$($f.ResourceType)] $($f.ResourceName)**`n"
            $mdContent += "  - Category: $($f.Category)`n"
            $mdContent += "  - $($f.Description)`n`n"
        }
    }

    $mdContent += "`n---`n*Generated by M365 Call Flow Visualizer Config Audit*`n"

    Set-Content -Path $mdPath -Value $mdContent -Encoding UTF8 -Force
    Write-Host "Markdown report saved to: $mdPath" -ForegroundColor Green
}

# ========================================
# SNAPSHOT EXPORT (for diff comparison)
# ========================================

if ($SaveSnapshot) {

    $snapshot = [PSCustomObject]@{
        Timestamp        = (Get-Date -Format 'o')
        TenantName       = (Get-CsTenant).DisplayName
        TenantId         = (Get-CsTenant).TenantId
        AutoAttendants   = $allAutoAttendants | ForEach-Object {
            [PSCustomObject]@{
                Identity              = $_.Identity
                Name                  = $_.Name
                LanguageId            = $_.LanguageId
                TimeZoneId            = $_.TimeZoneId
                VoiceResponseEnabled  = $_.VoiceResponseEnabled
                VoiceId               = $_.VoiceId
                Operator              = $_.Operator
                DefaultCallFlow       = $_.DefaultCallFlow
                CallFlows             = $_.CallFlows
                Schedules             = $_.Schedules
                CallHandlingAssociations = $_.CallHandlingAssociations
                DialByNameResourceId  = $_.DialByNameResourceId
                DirectoryLookupScope  = $_.DirectoryLookupScope
                ApplicationInstances  = $_.ApplicationInstances
            }
        }
        CallQueues       = $allCallQueues | ForEach-Object {
            [PSCustomObject]@{
                Identity                     = $_.Identity
                Name                         = $_.Name
                RoutingMethod                = $_.RoutingMethod
                PresenceBasedRouting         = $_.PresenceBasedRouting
                ConferenceMode               = $_.ConferenceMode
                AllowOptOut                  = $_.AllowOptOut
                AgentAlertTime               = $_.AgentAlertTime
                LanguageId                   = $_.LanguageId
                OverflowThreshold            = $_.OverflowThreshold
                OverflowAction               = $_.OverflowAction
                OverflowActionTarget         = $_.OverflowActionTarget
                TimeoutThreshold             = $_.TimeoutThreshold
                TimeoutAction                = $_.TimeoutAction
                TimeoutActionTarget          = $_.TimeoutActionTarget
                Agents                       = $_.Agents | ForEach-Object {
                    [PSCustomObject]@{
                        ObjectId = $_.ObjectId
                        OptIn    = $_.OptIn
                    }
                }
                DistributionLists            = $_.DistributionLists
                ChannelId                    = $_.ChannelId
                ApplicationInstances         = $_.ApplicationInstances
                IsCallbackEnabled            = $_.IsCallbackEnabled
                UseDefaultMusicOnHold        = $_.UseDefaultMusicOnHold
                MusicOnHoldAudioFileId       = $_.MusicOnHoldAudioFileId
            }
        }
        ResourceAccounts = $allResourceAccounts | ForEach-Object {
            [PSCustomObject]@{
                ObjectId            = $_.ObjectId
                DisplayName         = $_.DisplayName
                UserPrincipalName   = $_.UserPrincipalName
                ApplicationId       = $_.ApplicationId
                PhoneNumber         = $_.PhoneNumber
            }
        }
    }

    $snapshotPath = "$CustomFilePath\ConfigSnapshot_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $snapshot | ConvertTo-Json -Depth 20 | Set-Content -Path $snapshotPath -Encoding UTF8 -Force
    Write-Host "Configuration snapshot saved to: $snapshotPath" -ForegroundColor Green
    Write-Host "Use Compare-M365CFVConfigSnapshot.ps1 to diff two snapshots." -ForegroundColor Cyan
}
