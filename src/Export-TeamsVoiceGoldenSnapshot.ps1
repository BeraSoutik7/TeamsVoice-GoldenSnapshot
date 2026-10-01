<#
================================================================================
SCRIPT NAME : Export-TeamsVoiceGoldenSnapshot.ps1
DESCRIPTION : Exports an immutable golden configuration snapshot of Microsoft
              Teams Auto Attendants, linked Call Queues, agent rosters,
              schedules, and presence-based routing rules to a structured CSV.
AUTHOR      : Soutik Bera
REPOSITORY  : TeamsVoice-GoldenSnapshot
MODULES REQ : MicrosoftTeams (v4.0.0+), Microsoft.Graph.Authentication, 
              Microsoft.Graph.Users, Microsoft.Graph.Groups
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
# SECTION 0: SESSION AUTHENTICATION
# ------------------------------------------------------------------------------
Write-Host "[*] Checking session connections..." -ForegroundColor Cyan

try {
    $null = Get-CsTenant -ErrorAction SilentlyContinue
} catch {
    Connect-MicrosoftTeams
}

$graph = Get-MgContext -ErrorAction SilentlyContinue
if (-not $graph) {
    Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -NoWelcome
}

# In-memory caches to eliminate redundant Graph API queries (prevents HTTP 429 throttling)
$UserCache  = @{}
$GroupCache = @{}

function Get-CachedAgentInfo {
    param ([string]$UserId)
    if ([string]::IsNullOrWhiteSpace($UserId)) { return "None" }

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
    if ([string]::IsNullOrWhiteSpace($GroupId)) { return "None" }

    if (-not $GroupCache.ContainsKey($GroupId)) {
        $members = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction SilentlyContinue
        $roster = foreach ($m in $members) { Get-CachedAgentInfo -UserId $m.Id }
        $GroupCache[$GroupId] = ($roster -join " ; ")
    }
    return $GroupCache[$GroupId]
}

# ------------------------------------------------------------------------------
# SECTION 1: LOAD TOPOLOGY MAPS & INITIALIZE CACHE
# ------------------------------------------------------------------------------
Write-Host "[*] Querying Auto Attendants, Call Queues, and Resource Accounts..." -ForegroundColor Cyan

$AAs          = Get-CsAutoAttendant
$CQs          = Get-CsCallQueue
$AppInstances = Get-CsOnlineApplicationInstance

# Case-insensitive queue lookup dictionary using trimmed GUID strings
$CQMap = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($q in $CQs) {
    if ($q.Identity) {
        $CQMap[$q.Identity.ToString().Trim()] = $q
    }
}

# Case-insensitive resource account lookup dictionary
$RAMap = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($ra in $AppInstances) {
    if ($ra.ObjectId) {
        $RAMap[$ra.ObjectId.ToString().Trim()] = $ra
    }
}

$SnapshotData = [System.Collections.Generic.List[PSCustomObject]]::new()

# ------------------------------------------------------------------------------
# SECTION 2: COMPILE GOLDEN SNAPSHOT PER ROUTE
# ------------------------------------------------------------------------------
Write-Host "[*] Compiling golden configuration snapshot across all call flows..." -ForegroundColor Cyan

foreach ($aa in $AAs) {
    # 2.1 Associated Resource Accounts & DIDs
    $raList = [System.Collections.Generic.List[string]]::new()
    $raDIDs = [System.Collections.Generic.List[string]]::new()

    if ($aa.ApplicationInstances) {
        foreach ($raId in $aa.ApplicationInstances) {
            $idStr = $raId.ToString().Trim()
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

    # 2.2 Operator Details
    $operator = "None"
    if ($aa.Operator) {
        $opType   = $aa.Operator.Type
        $opTarget = [string]$aa.Operator.Target
        if ($opType -eq "User") {
            $operator = "User: " + (Get-CachedAgentInfo -UserId $opTarget)
        } elseif ($opType -eq "Queue" -and $CQMap.ContainsKey($opTarget)) {
            $operator = "Queue: $($CQMap[$opTarget].Name)"
        } else {
            $operator = "$opType : $opTarget"
        }
    }

    # 2.3 Authorized Users
    $authUsers = [System.Collections.Generic.List[string]]::new()
    if ($aa.AuthorizedUsers) {
        foreach ($authId in $aa.AuthorizedUsers) {
            $authUsers.Add((Get-CachedAgentInfo -UserId $authId.ToString().Trim()))
        }
    }
    $authUsersJoined = if ($authUsers.Count -gt 0) { $authUsers -join " | " } else { "None" }

    # 2.4 Schedule Associations (After-Hours & Holiday)
    $schedules = [System.Collections.Generic.List[string]]::new()
    if ($aa.CallHandlingAssociations) {
        foreach ($assoc in $aa.CallHandlingAssociations) {
            $schedules.Add("$($assoc.Type): $($assoc.ScheduleId)")
        }
    }
    $schedulesJoined = if ($schedules.Count -gt 0) { $schedules -join " | " } else { "DefaultOnly" }

    # Helper function to generate standardized CSV records
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
                    $agents.Add((Get-CachedAgentInfo -UserId $u.ToString().Trim()))
                }
            }

            # Security Groups / Distribution Lists
            if ($TargetQueue.DistributionLists) {
                foreach ($gid in $TargetQueue.DistributionLists) {
                    $agents.Add("[Group: $gid] " + (Get-CachedGroupMembers -GroupId $gid.ToString().Trim()))
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

    # 2.5 Extract All Menu Options Across Default & Handling Associations
    $allMenusToProcess = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Inspect Default Call Flow
    if ($aa.DefaultCallFlow -and $aa.DefaultCallFlow.Menu -and $aa.DefaultCallFlow.Menu.MenuOptions) {
        foreach ($opt in $aa.DefaultCallFlow.Menu.MenuOptions) {
            $allMenusToProcess.Add([PSCustomObject]@{
                FlowType = "DefaultFlow"
                Option   = $opt
            })
        }
    }

    # Inspect Handling Associations (AfterHours / Holiday)
    if ($aa.CallHandlingAssociations -and $aa.CallFlows) {
        foreach ($assoc in $aa.CallHandlingAssociations) {
            $assocFlow = $aa.CallFlows | Where-Object { $_.Identity -eq $assoc.CallFlowId }
            if ($assocFlow -and $assocFlow.Menu -and $assocFlow.Menu.MenuOptions) {
                foreach ($opt in $assocFlow.Menu.MenuOptions) {
                    $allMenusToProcess.Add([PSCustomObject]@{
                        FlowType = $assoc.Type
                        Option   = $opt
                    })
                }
            }
        }
    }

    # Process Menu Routes
    if ($allMenusToProcess.Count -gt 0) {
        foreach ($entry in $allMenusToProcess) {
            $opt        = $entry.Option
            $tQueue     = $null
            $tName      = "Unknown"
            $tId        = "None"

            # 1. Resolve Keypad Key
            $dtmfKey = "KeyUnknown"
            if ($opt.Key) { $dtmfKey = [string]$opt.Key }
            elseif ($opt.Option) { $dtmfKey = [string]$opt.Option }
            elseif ($opt.DtmfResponse) { $dtmfKey = [string]$opt.DtmfResponse }

            # 2. Resolve Target ID across SDK property variants
            if ($opt.CallTarget) {
                if ($opt.CallTarget.Id) { $tId = [string]$opt.CallTarget.Id.ToString().Trim() }
                elseif ($opt.CallTarget.Target) { $tId = [string]$opt.CallTarget.Target.ToString().Trim() }
                elseif ($opt.CallTarget.Identity) { $tId = [string]$opt.CallTarget.Identity.ToString().Trim() }
                elseif ($opt.CallTarget.PhoneNumber) { $tId = [string]$opt.CallTarget.PhoneNumber.ToString().Trim() }
            }

            $targetType = if ($opt.CallTarget -and $opt.CallTarget.Type) { [string]$opt.CallTarget.Type } else { "Unknown" }

            # 3. Match against Queues or resolve target type
            if ($targetType -eq "Queue" -or $CQMap.ContainsKey($tId)) {
                if ($CQMap.ContainsKey($tId)) {
                    $tQueue = $CQMap[$tId]
                    $tName  = $tQueue.Name
                } else {
                    $tName = "CallQueue (ID: $tId)"
                }
            } elseif ($targetType -like "*External*" -or $targetType -like "*Pstn*") {
                $tName = "External PSTN: $tId"
            } elseif ($targetType -eq "User") {
                $tName = "User: " + (Get-CachedAgentInfo -UserId $tId)
            } elseif ($targetType -like "*Voicemail*") {
                $tName = "Shared Voicemail: $tId"
            } else {
                if ($tId -ne "None") {
                    $tName = "$targetType ($tId)"
                } else {
                    $tName = "Announcement/Disconnect"
                }
            }

            Add-GoldenRecord `
                -Key "$($entry.FlowType):$dtmfKey" `
                -Action $opt.Action `
                -TargetName $tName `
                -TargetId $tId `
                -TargetQueue $tQueue
        }
    } else {
        # Process Direct Unconditional Call Forwarding
        $directId = "None"
        if ($aa.DefaultCallFlow -and$aa.DefaultCallFlow.TransferTarget) {
            $tt =$aa.DefaultCallFlow.TransferTarget
            if ($tt.Id) { $directId = [string]$tt.Id.ToString().Trim() }
            elseif ($tt.Target) { $directId = [string]$tt.Target.ToString().Trim() }
            elseif ($tt.Identity) { $directId = [string]$tt.Identity.ToString().Trim() }
            elseif ($tt.PhoneNumber) { $directId = [string]$tt.PhoneNumber.ToString().Trim() }
        }

        $tQueue = if ($directId -ne "None" -and $CQMap.ContainsKey($directId)) {$CQMap[$directId] } else {$null }
        $tName  = if ($tQueue) { $tQueue.Name } else { "DirectRoute ($directId)" }

        Add-GoldenRecord `
            -Key "DirectRoute" `
            -Action "TransferCallToTarget" `
            -TargetName $tName `
            -TargetId $directId `
            -TargetQueue $tQueue
    }
}

# ------------------------------------------------------------------------------
# SECTION 3: EXPORT GOLDEN SNAPSHOT
# ------------------------------------------------------------------------------
$SnapshotData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
Write-Host "`n[SUCCESS] Golden Snapshot exported successfully to: $OutputPath" -ForegroundColor Green