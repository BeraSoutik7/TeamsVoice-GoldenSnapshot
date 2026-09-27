<#
================================================================================
SCRIPT NAME : Export-TeamsVoiceGoldenSnapshot.ps1
DESCRIPTION : Exports an immutable golden configuration snapshot of Microsoft
              Teams Auto Attendants, linked Call Queues, agent rosters,
              schedules, and presence-based routing rules to a structured CSV.
AUTHOR      : Soutik Bera
REPOSITORY  : TeamsVoice-GoldenSnapshot
MODULES REQ : MicrosoftTeams, Microsoft.Graph.Users, Microsoft.Graph.Groups
================================================================================
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$OutputDir = ".\snapshots",

    [Parameter(Mandatory = $false)]
    [string]$FileName
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

if (-not $FileName) {
    $FileName = "TeamsVoice_GoldenSnapshot_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
}

$OutputPath = Join-Path $OutputDir $FileName

# ------------------------------------------------------------------------------
# STEP 0: SESSION AUTHENTICATION
# ------------------------------------------------------------------------------
Write-Host "[*] Validating session connections..." -ForegroundColor Cyan

try {
    $null = Get-CsTenant -ErrorAction SilentlyContinue
} catch {
    Connect-MicrosoftTeams
}

$graph = Get-MgContext -ErrorAction SilentlyContinue
if (-not $graph) {
    Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -NoWelcome
}

# In-memory caches to eliminate redundant Graph API queries (avoids 429 throttling)
$UserCache  = @{}
$GroupCache = @{}

function Get-CachedAgentInfo {
    param ([string]$UserId)
    if (-not $UserCache.ContainsKey($UserId)) {
        $u = Get-MgUser -UserId $UserId -Property UserPrincipalName, DisplayName, AccountEnabled -ErrorAction SilentlyContinue
        if ($u) {
            $status = if ($u.AccountEnabled) { "Active" } else { "Disabled" }
            $UserCache[$UserId] = "$($u.DisplayName) ($($u.UserPrincipalName)) [$status]"
        } else {
            $UserCache[$UserId] = $UserId
        }
    }
    return $UserCache[$UserId]
}

function Get-CachedGroupMembers {
    param ([string]$GroupId)
    if (-not $GroupCache.ContainsKey($GroupId)) {
        $members = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction SilentlyContinue
        $roster = foreach ($m in $members) { Get-CachedAgentInfo -UserId $m.Id }
        $GroupCache[$GroupId] = ($roster -join " ; ")
    }
    return $GroupCache[$GroupId]
}

# ------------------------------------------------------------------------------
# STEP 1: LOAD VOICE TOPOLOGY CACHE
# ------------------------------------------------------------------------------
Write-Host "[*] Fetching Auto Attendants, Call Queues, and Resource Accounts..." -ForegroundColor Cyan

$AAs          = Get-CsAutoAttendant
$CQs          = Get-CsCallQueue
$AppInstances = Get-CsOnlineApplicationInstance

# Build fast hash map lookups
$CQMap = @{}
foreach ($q in $CQs) { $CQMap[$q.Identity.ToString()] = $q }

$RAMap = @{}
foreach ($ra in $AppInstances) { $RAMap[$ra.ObjectId.ToString()] = $ra }

$SnapshotData = [System.Collections.Generic.List[PSCustomObject]]::new()

# ------------------------------------------------------------------------------
# STEP 2: COMPILE GOLDEN SNAPSHOT PER ROUTE
# ------------------------------------------------------------------------------
Write-Host "[*] Processing call trees and building golden records..." -ForegroundColor Cyan

foreach ($aa in $AAs) {
    # 2.1 Bound Resource Accounts & Phone Numbers
    $raList   = [System.Collections.Generic.List[string]]::new()
    $raDIDs   = [System.Collections.Generic.List[string]]::new()

    if ($aa.ApplicationInstances) {
        foreach ($raId in $aa.ApplicationInstances) {
            $idStr = $raId.ToString()
            if ($RAMap.ContainsKey($idStr)) {
                $raObj = $RAMap[$idStr]
                $raList.Add($raObj.UserPrincipalName)
                if ($raObj.PhoneNumber) { $raDIDs.Add($raObj.PhoneNumber) }
            } else {
                $raList.Add($idStr)
            }
        }
    }

    $raAccountsJoined = if ($raList.Count -gt 0) { $raList -join " | " } else { "None" }
    $raDIDsJoined     = if ($raDIDs.Count -gt 0) { $raDIDs -join " | " } else { "None" }

    # 2.2 Operator Target
    $operator = if ($aa.Operator) { "$($aa.Operator.Type): $($aa.Operator.Target)" } else { "None" }

    # 2.3 Authorized Users (Users allowed to edit voice greetings via Teams client)
    $authUsers = [System.Collections.Generic.List[string]]::new()
    if ($aa.AuthorizedUsers) {
        foreach ($authId in $aa.AuthorizedUsers) {
            $authUsers.Add((Get-CachedAgentInfo -UserId $authId.ToString()))
        }
    }
    $authUsersJoined = if ($authUsers.Count -gt 0) { $authUsers -join " | " } else { "None" }

    # 2.4 After-Hours & Holiday Associations
    $schedules = [System.Collections.Generic.List[string]]::new()
    if ($aa.CallHandlingAssociations) {
        foreach ($assoc in $aa.CallHandlingAssociations) {
            $schedules.Add("$($assoc.Type): $($assoc.ScheduleId)")
        }
    }
    $schedulesJoined = if ($schedules.Count -gt 0) { $schedules -join " | " } else { "DefaultOnly" }

    # Helper function to write standardized CSV record
    function Add-GoldenRecord {
        param (
            [string]$Key,
            [string]$Action,
            [string]$TargetName,
            [string]$TargetId,
            $TargetQueue
        )

        $cqMethod    = "N/A"
        $cqPresence  = "N/A"
        $cqConfMode  = "N/A"
        $cqRoster    = "N/A"
        $cqOverflow  = "N/A"
        $cqTimeout   = "N/A"

        if ($TargetQueue) {
            $cqMethod   = $TargetQueue.DistributionMethod
            $cqPresence = $TargetQueue.ConferenceMode
            $cqConfMode = $TargetQueue.ConferenceMode

            $agents = [System.Collections.Generic.List[string]]::new()

            # Direct Users
            if ($TargetQueue.Users) {
                foreach ($u in $TargetQueue.Users) {
                    $agents.Add((Get-CachedAgentInfo -UserId $u.ToString()))
                }
            }

            # Security Groups / Distribution Lists
            if ($TargetQueue.DistributionLists) {
                foreach ($gid in $TargetQueue.DistributionLists) {
                    $agents.Add("[Group: $gid] " + (Get-CachedGroupMembers -GroupId $gid.ToString()))
                }
            }

            $cqRoster   = if ($agents.Count -gt 0) { $agents -join " | " } else { "EMPTY_ROSTER" }
            $cqOverflow = "MaxCalls: $($TargetQueue.OverflowThreshold) -> $($TargetQueue.OverflowAction) ($($TargetQueue.OverflowActionTarget))"
            $cqTimeout  = "Timeout: $($TargetQueue.TimeoutThreshold)s -> $($TargetQueue.TimeoutAction) ($($TargetQueue.TimeoutActionTarget))"
        }

        $SnapshotData.Add([PSCustomObject]@{
            SnapshotDate            = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            AutoAttendantName       = $aa.Name
            AutoAttendantId         = $aa.Identity
            TimeZone                = $aa.TimeZoneId
            Language                = $aa.LanguageId
            ResourceAccounts        = $raAccountsJoined
            AssignedDIDs            = $raDIDsJoined
            AuthorizedUsers         = $authUsersJoined
            OperatorTarget          = $operator
            ScheduleAssociations   = $schedulesJoined
            Route_Key               = $Key
            Route_Action            = $Action
            Target_EntityName       = $TargetName
            Target_EntityId         = $TargetId
            CQ_DistributionMethod   = $cqMethod
            CQ_PresenceBasedRouting = $cqPresence
            CQ_AgentRoster          = $cqRoster
            CQ_OverflowRule         = $cqOverflow
            CQ_TimeoutRule          = $cqTimeout
        })
    }

    # 2.5 Extract Menu Options or Direct Forwarding
    $menuOptions = @()
    if ($aa.DefaultCallFlow -and $aa.DefaultCallFlow.Menu -and $aa.DefaultCallFlow.Menu.MenuOptions) {
        $menuOptions = $aa.DefaultCallFlow.Menu.MenuOptions
    }

    if ($menuOptions.Count -gt 0) {
        foreach ($opt in $menuOptions) {
            $tQueue = $null
            $tName  = "External/DirectTarget"
            $tId    = "None"

            if ($opt.CallTarget -and $opt.CallTarget.Identity) {
                $tId = [string]$opt.CallTarget.Identity
            }

            if ($opt.CallTarget -and $opt.CallTarget.Type -eq "Queue" -and $tId -ne "None" -and $CQMap.ContainsKey($tId)) {
                $tQueue = $CQMap[$tId]
                $tName  = $tQueue.Name
            }

            Add-GoldenRecord -Key $opt.Key -Action $opt.Action -TargetName $tName -TargetId $tId -TargetQueue $tQueue
        }
    } else {
        $directId = "None"
        if ($aa.DefaultCallFlow -and $aa.DefaultCallFlow.TransferTarget -and $aa.DefaultCallFlow.TransferTarget.Identity) {
            $directId = [string]$aa.DefaultCallFlow.TransferTarget.Identity
        }

        $tQueue = if ($directId -ne "None" -and $CQMap.ContainsKey($directId)) { $CQMap[$directId] } else { $null }
        $tName  = if ($tQueue) { $tQueue.Name } else { "DirectCallFlow" }

        Add-GoldenRecord -Key "DirectRoute" -Action "TransferCallToTarget" -TargetName $tName -TargetId $directId -TargetQueue $tQueue
    }
}

# ------------------------------------------------------------------------------
# STEP 3: EXPORT
# ------------------------------------------------------------------------------
$SnapshotData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
Write-Host "`n[SUCCESS] Golden Snapshot exported successfully to: $OutputPath" -ForegroundColor Green