<#
.SYNOPSIS
Generates an HTML documentation report for all Microsoft Teams Call Queues.

.DESCRIPTION
This read-only script enumerates all Teams Call Queues (CQs) and produces a structured HTML report.
It resolves friendly names for targets and agents, and includes bracketed audit details such as
raw IDs and resource account attachments.

The report includes:
 - Routing Method
 - Agent Alert Time
 - Agent Opt Out Allowed
 - Agents (Users and Distribution Lists), resolved to friendly names
 - Exceptions: Overflow (threshold, action, target, greeting) — each on its own line
 - Exceptions: Timeout (threshold, action, target, greeting) — each on its own line
 - Exception: No Agents (apply-to, action, target, greeting) — each on its own line
 - Callback Eligible (Enabled + detailed thresholds)
 - Callback Message (TTS: <text> | AudioFile: <name> | None)
 - Callback Key (DTMF: 0–9, *, #)
 - Callback Fail Notification (friendly target)
 - Authorised Users (friendly DisplayName)

TARGET RESOLUTION (best effort)
 - PSTN: strips 'tel:' and displays E.164
 - Call Queue: CQ name by Id
 - Auto Attendant: AA name by Id
 - Resource Account (ApplicationInstance): RA display name and, if mapped, the attached CQ/AA name
 - Group (Shared Voicemail / DL): Team/M365 Group display name where possible (Teams/EXO/Graph)
 - User: DisplayName (falls back to UPN/Id)
 - Otherwise: shows the raw Id as “Unknown”

REQUIREMENTS
 - Teams/Skype Online PowerShell with access to:
     * Get-CsCallQueue
     * Get-CsAutoAttendant (for target resolution)
     * Get-CsOnlineApplicationInstance (resource accounts)
     * Get-CsOnlineUser (agents/users)
 - Optional (for richer group names):
     * MicrosoftTeams: Get-Team
     * ExchangeOnlineManagement: Get-UnifiedGroup
     * Microsoft.Graph: Get-MgGroup

.PARAMETER OutputPath
Full or relative path to the HTML report file.
Default: .\Document-TeamsCallQueues-<timestamp>.html

.PARAMETER Open
If specified, opens the generated HTML report when complete.

.EXAMPLE
PS> .\Document-TeamsCallQueues.ps1
Generates ".\Document-TeamsCallQueues-YYYYMMDD-HHmmss.html".

.EXAMPLE
PS> .\Document-TeamsCallQueues.ps1 -OutputPath "C:\Temp\CQ-Report.html" -Open
Writes the report to C:\Temp\CQ-Report.html and opens it in the default browser.

.NOTES
 - Parser-safety: Hashtables use ';' between key/value pairs. Complex values are built in variables first.
 - Null-safety: All .ContainsKey / indexing operations guard against null/empty keys.
 - This script is read-only; it does not modify any tenant configuration.
#>

[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path (Get-Location) ("Document-TeamsCallQueues-{0}.html" -f (Get-Date -Format "yyyyMMdd-HHmmss"))),
    [switch]$Open
)

$ErrorActionPreference = "Stop"

# ==============================
# Utilities (HTML)
# ==============================
function HtmlEncode {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    [System.Net.WebUtility]::HtmlEncode($Text)
}

# ==============================
# Caches for lookups
# ==============================
$script:Cache = @{
    Users   = @{}  # userId/upn -> label
    Groups  = @{}  # groupId     -> display
    CQs     = @{}  # cqId        -> name
    AAs     = @{}  # aaId        -> name
    RAs     = @{}  # raId        -> display
    CQByRA  = @{}  # raId        -> cqName
    AAByRA  = @{}  # raId        -> aaName
}

# ==============================
# Index building (CQ, AA, RA maps)
# ==============================
function Build-Index {
    <#
      .SYNOPSIS
      Pre-loads CQ/AA names and RA→CQ/AA mappings for fast target resolution.
    #>

    # Call Queues
    $allCqs = @()
    try { $allCqs = @(Get-CsCallQueue) } catch {}
    foreach ($cq in $allCqs) {
        if ($null -eq $cq) { continue }
        $cqId = [string]$cq.Identity
        if (-not [string]::IsNullOrWhiteSpace($cqId)) {
            $script:Cache.CQs[$cqId] = $cq.Name
        }
        foreach ($ra in @($cq.ResourceAccounts)) {
            $raStr = [string]$ra
            if ([string]::IsNullOrWhiteSpace($raStr)) { continue }
            $script:Cache.CQByRA[$raStr] = $cq.Name
        }
    }

    # Auto Attendants
    $allAas = @()
    try { $allAas = @(Get-CsAutoAttendant) } catch {}
    foreach ($aa in $allAas) {
        if ($null -eq $aa) { continue }
        $aaId = [string]$aa.Identity
        if (-not [string]::IsNullOrWhiteSpace($aaId)) {
            $script:Cache.AAs[$aaId] = $aa.Name
        }
        foreach ($ra in @($aa.ApplicationInstances)) {
            $raStr = [string]$ra
            if ([string]::IsNullOrWhiteSpace($raStr)) { continue }
            $script:Cache.AAByRA[$raStr] = $aa.Name
        }
    }
}

# ==============================
# Identity / label resolvers
# ==============================
function Resolve-User {
    <#
      .SYNOPSIS
      Returns a friendly label for a user id/UPN; caches results (DisplayName-first).
    #>
    param($Id)
    $idStr = [string]$Id
    if ([string]::IsNullOrWhiteSpace($idStr)) { return "" }
    if (-not $script:Cache.Users.ContainsKey($idStr)) {
        try {
            $u = Get-CsOnlineUser -Identity $idStr -ErrorAction Stop
            $label = $null
            if ($u) {
                if     ($u.DisplayName)       { $label = $u.DisplayName }
                elseif ($u.UserPrincipalName) { $label = $u.UserPrincipalName }
                elseif ($u.SipAddress)        { $label = $u.SipAddress }
            }
            if (-not $label) { $label = $idStr }
            $script:Cache.Users[$idStr] = $label
        } catch {
            $script:Cache.Users[$idStr] = $idStr
        }
    }
    $script:Cache.Users[$idStr]
}

function Resolve-Group {
    <#
      .SYNOPSIS
      Returns a friendly label for a Group/Team/DL id; caches results.
      .NOTES
      Attempts Get-Team, then Get-UnifiedGroup, then Get-MgGroup; if all fail, returns the raw id.
    #>
    param($Id)
    $idStr = [string]$Id
    if ([string]::IsNullOrWhiteSpace($idStr)) { return "" }
    if ($script:Cache.Groups.ContainsKey($idStr)) { return $script:Cache.Groups[$idStr] }

    $name = $null

    # Try Microsoft Teams (Team) first
    try {
        $t = Get-Team -GroupId $idStr -ErrorAction Stop
        if ($t) { $name = $t.DisplayName }
    } catch {}

    # Then EXO Unified Group
    if (-not $name) {
        try {
            $g = Get-UnifiedGroup -Identity $idStr -ErrorAction Stop
            if ($g) { $name = $g.DisplayName }
        } catch {}
    }

    # Then Microsoft Graph group (if available)
    if (-not $name) {
        try {
            $gg = Get-MgGroup -GroupId $idStr -Property DisplayName -ErrorAction Stop
            if ($gg) { $name = $gg.DisplayName }
        } catch {}
    }

    if (-not $name) { $name = $idStr }

    $script:Cache.Groups[$idStr] = $name
    $name
}

function Get-RawIdFromValue {
    <#
      .SYNOPSIS
      Extracts an ID string from strings, GUIDs, or typed target objects.
    #>
    param($Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
        return $Value
    }

    if ($Value -is [guid]) { return $Value.Guid }

    if ($Value.PSObject) {
        foreach ($prop in 'Id','Identity','ObjectId','Target','TargetId') {
            if ($Value.PSObject.Properties[$prop]) {
                $v = [string]$Value.$prop
                if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
            }
        }
    }

    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $s
}

# Robust target resolver for CQ exception targets (GUID/UPN/tel/typed)
function Resolve-Target {
    <#
      .SYNOPSIS
      Best-effort target resolver for CQ exception targets.
      .RETURNS
      @{ Friendly = "<label>"; Bracket = "(human-type; raw-id [; attached to ...])" }
    #>
    param($Value)

    $rawId = Get-RawIdFromValue $Value
    if ([string]::IsNullOrWhiteSpace($rawId)) {
        return @{ Friendly = ""; Bracket = "" }
    }

    # PSTN
    if ($rawId -match '^tel:\+?\d' -or $rawId -match '^\+?\d[\d\s\-()]{6,}$') {
        $num = ($rawId -replace '^tel:')
        return @{ Friendly = $num; Bracket = "(phone number; $rawId)" }
    }

    # Call Queue
    if ($script:Cache.CQs.ContainsKey($rawId)) {
        return @{ Friendly = $script:Cache.CQs[$rawId]; Bracket = "(call queue; $rawId)" }
    }
    try {
        $cq = Get-CsCallQueue -Identity $rawId -ErrorAction Stop
        if ($cq) {
            $script:Cache.CQs[$rawId] = $cq.Name
            return @{ Friendly = $cq.Name; Bracket = "(call queue; $rawId)" }
        }
    } catch {}

    # Auto Attendant
    if ($script:Cache.AAs.ContainsKey($rawId)) {
        return @{ Friendly = $script:Cache.AAs[$rawId]; Bracket = "(auto attendant; $rawId)" }
    }
    try {
        $aa = Get-CsAutoAttendant -Identity $rawId -ErrorAction Stop
        if ($aa) {
            $script:Cache.AAs[$rawId] = $aa.Name
            return @{ Friendly = $aa.Name; Bracket = "(auto attendant; $rawId)" }
        }
    } catch {}

    # Resource Account (RA) and RA→CQ/AA attachments
    try {
        if (-not $script:Cache.RAs.ContainsKey($rawId)) {
            $ra = Get-CsOnlineApplicationInstance -Identity $rawId -ErrorAction Stop
            if ($ra) { $script:Cache.RAs[$rawId] = $ra.DisplayName }
        }
    } catch {}
    if ($script:Cache.RAs.ContainsKey($rawId)) {
        if ($script:Cache.CQByRA.ContainsKey($rawId)) {
            $name = $script:Cache.CQByRA[$rawId]
            return @{ Friendly = $name; Bracket = "(resource account; $rawId; attached to call queue)" }
        }
        if ($script:Cache.AAByRA.ContainsKey($rawId)) {
            $name = $script:Cache.AAByRA[$rawId]
            return @{ Friendly = $name; Bracket = "(resource account; $rawId; attached to auto attendant)" }
        }
        return @{ Friendly = $script:Cache.RAs[$rawId]; Bracket = "(resource account; $rawId)" }
    }

    # Group (Shared Voicemail / DL) — try Team/UnifiedGroup/Graph without gating
    $grpName = Resolve-Group $rawId
    if ($grpName -and $grpName -ne $rawId) {
        return @{ Friendly = $grpName; Bracket = "(group; $rawId)" }
    }

    # User (UPN/ObjectId) fallback
    $userLabel = Resolve-User $rawId
    if ($userLabel -and $userLabel -ne $rawId) {
        return @{ Friendly = $userLabel; Bracket = "(user; $rawId)" }
    }

    # Unknown
    return @{ Friendly = $rawId; Bracket = "(unknown; $rawId)" }
}

function Map-ActionLabel {
    <#
      .SYNOPSIS
      Normalizes CQ action enums to human labels; removes accidental duplicates.
    #>
    param([string]$Action)
    if ([string]::IsNullOrWhiteSpace($Action)) { return "Unknown" }

    $label = switch -Regex ($Action) {
        'SharedVoicemail'   { "Shared Voicemail"; break }
        'Voicemail'         { "Voicemail"; break }
        'Disconnect|Busy'   { "Disconnect"; break }
        'Queue'             { "Queue"; break }
        'Transfer|Redirect' { "Forward/Transfer"; break }
        default             { $Action }
    }

    # De-duplicate if concatenated (e.g., "Shared Voicemail Voicemail")
    $label = ($label -replace '\bShared Voicemail\s+Voicemail\b','Shared Voicemail')
    $label = ($label -replace '\bVoicemail\s+Voicemail\b','Voicemail')
    $label
}

function Get-DtmfLabel {
    <#
      .SYNOPSIS
      Maps DTMF enums (Tone1, Star, Pound) to keys (1, *, #).
    #>
    param([string]$Dtmf)
    switch ($Dtmf) {
        "Tone0" { "0" }
        "Tone1" { "1" }
        "Tone2" { "2" }
        "Tone3" { "3" }
        "Tone4" { "4" }
        "Tone5" { "5" }
        "Tone6" { "6" }
        "Tone7" { "7" }
        "Tone8" { "8" }
        "Tone9" { "9" }
        "Star"  { "*" }
        "Pound" { "#" }
        default { $Dtmf }
    }
}

# ==============================
# Exception greeting helper
# ==============================
function Get-ExceptionGreeting {
    <#
      .SYNOPSIS
      Returns "TTS: <text>" | "AudioFile: <name>" | "None" for the given exception prefix.
      .PARAMETER Prefix
      One of: Overflow | Timeout | NoAgent
    #>
    param(
        [Parameter(Mandatory)]$CQ,
        [Parameter(Mandatory)][ValidateSet('Overflow','Timeout','NoAgent')] [string]$Prefix
    )

    # Prefer TTS first
    $ttsCandidates = @(
        "${Prefix}SharedVoicemailTextToSpeechPrompt",
        "${Prefix}DisconnectTextToSpeechPrompt",
        "${Prefix}RedirectPersonTextToSpeechPrompt",
        "${Prefix}RedirectVoiceAppTextToSpeechPrompt",
        "${Prefix}RedirectPhoneNumberTextToSpeechPrompt",
        "${Prefix}RedirectVoicemailTextToSpeechPrompt"
    )
    foreach ($f in $ttsCandidates) {
        if ($CQ.PSObject.Properties[$f]) {
            $val = [string]$CQ.$f
            if (-not [string]::IsNullOrWhiteSpace($val)) { return "TTS: $(HtmlEncode $val)" }
        }
    }

    # Then audio file names
    $fileCandidates = @(
        "${Prefix}SharedVoicemailAudioFilePromptFileName",
        "${Prefix}DisconnectAudioFilePromptFileName",
        "${Prefix}RedirectPersonAudioFilePromptFileName",
        "${Prefix}RedirectVoiceAppAudioFilePromptFileName",
        "${Prefix}RedirectPhoneNumberAudioFilePromptFileName",
        "${Prefix}RedirectVoicemailAudioFilePromptFileName"
    )
    foreach ($f in $fileCandidates) {
        if ($CQ.PSObject.Properties[$f]) {
            $val = [string]$CQ.$f
            if (-not [string]::IsNullOrWhiteSpace($val)) { return "AudioFile: $(HtmlEncode $val)" }
        }
    }

    "None"
}

# ==============================
# HTML table helper
# ==============================
function New-Table {
    <#
      .SYNOPSIS
      Builds a "Parameter / Value" HTML table for a section.
    #>
    param($Title, $Rows)
    $html = "<h2>$(HtmlEncode $Title)</h2><table><thead><tr><th>Parameter</th><th>Value</th></tr></thead><tbody>"
    foreach ($r in $Rows) {
        $paramCell = HtmlEncode $r.Parameter
        $valCell   = [string]$r.Value
        $html += "<tr><td>$paramCell</td><td>$valCell</td></tr>"
    }
    $html += "</tbody></table>"
    $html
}

# ==============================
# MAIN
# ==============================
Build-Index

$cqs = @(Get-CsCallQueue)
if ($cqs.Count -eq 0) { throw "No call queues found." }

$sections = @()

foreach ($cq in $cqs) {

    $rows = @()

    # --- Core routing lines ---
    $rows += @{ Parameter = "Routing Method";        Value = HtmlEncode ([string]$cq.RoutingMethod) }
    $rows += @{ Parameter = "Agent Alert Time";       Value = HtmlEncode ([string]$cq.AgentAlertTime) }
    $rows += @{ Parameter = "Agent Opt Out Allowed";  Value = ($(if ($cq.AllowOptOut) { "Yes" } else { "No" })) }

    # --- Agents (Users + DLs) ---
    $agentBlocks = @()

    if ($cq.Users) {
        $userLines = $cq.Users | ForEach-Object { HtmlEncode (Resolve-User $_) }
        if ($userLines -and $userLines.Count -gt 0) {
            $agentBlocks += "<b>Users:</b><br/>" + ($userLines -join "<br/>")
        }
    }

    if ($cq.DistributionLists) {
        $dlLines = $cq.DistributionLists | ForEach-Object { HtmlEncode (Resolve-Group $_) }
        if ($dlLines -and $dlLines.Count -gt 0) {
            $agentBlocks += "<b>Distribution Lists:</b><br/>" + ($dlLines -join "<br/>")
        }
    }

    if (-not $agentBlocks -or $agentBlocks.Count -eq 0) { $agentBlocks = @("None") }
    $rows += @{ Parameter = "Agents"; Value = ($agentBlocks -join "<br/><br/>") }

    # --- Exceptions: Overflow ---
    $ovParts = New-Object System.Collections.Generic.List[string]
    $ovTh   = HtmlEncode ([string]$cq.OverflowThreshold)
    $ovAct  = Map-ActionLabel -Action ([string]$cq.OverflowAction)
    $ovTgt  = Resolve-Target $cq.OverflowActionTarget
    $ovTgtText = if ($ovTgt.Friendly) { (HtmlEncode $ovTgt.Friendly) + ($(if ($ovTgt.Bracket) { " <span style='color:#605E5C'>$($ovTgt.Bracket)</span>" } else { "" })) } else { "None" }
    $ovGreet = Get-ExceptionGreeting -CQ $cq -Prefix Overflow
    [void]$ovParts.Add("<b>Threshold:</b> $ovTh")
    [void]$ovParts.Add("<b>Action:</b> $(HtmlEncode $ovAct)")
    [void]$ovParts.Add("<b>Target:</b> $ovTgtText")
    [void]$ovParts.Add("<b>Greeting:</b> $ovGreet")
    $rows += @{ Parameter = "Exceptions: Overflow"; Value = ($ovParts -join "<br/>") }

    # --- Exceptions: Timeout ---
    $toParts = New-Object System.Collections.Generic.List[string]
    $toTh   = HtmlEncode ([string]$cq.TimeoutThreshold)
    $toAct  = Map-ActionLabel -Action ([string]$cq.TimeoutAction)
    $toTgt  = Resolve-Target $cq.TimeoutActionTarget
    $toTgtText = if ($toTgt.Friendly) { (HtmlEncode $toTgt.Friendly) + ($(if ($toTgt.Bracket) { " <span style='color:#605E5C'>$($toTgt.Bracket)</span>" } else { "" })) } else { "None" }
    $toGreet = Get-ExceptionGreeting -CQ $cq -Prefix Timeout
    [void]$toParts.Add("<b>Threshold (sec):</b> $toTh")
    [void]$toParts.Add("<b>Action:</b> $(HtmlEncode $toAct)")
    [void]$toParts.Add("<b>Target:</b> $toTgtText")
    [void]$toParts.Add("<b>Greeting:</b> $toGreet")
    $rows += @{ Parameter = "Exceptions: Timeout"; Value = ($toParts -join "<br/>") }

    # --- Exception: No Agents ---
    $naParts = New-Object System.Collections.Generic.List[string]
    [void]$naParts.Add("<b>Apply to:</b> $(HtmlEncode ([string]$cq.NoAgentApplyTo))")
    $naAct = Map-ActionLabel -Action ([string]$cq.NoAgentAction)
    $naTgt = Resolve-Target $cq.NoAgentActionTarget
    $naTgtText = if ($naTgt.Friendly) { (HtmlEncode $naTgt.Friendly) + ($(if ($naTgt.Bracket) { " <span style='color:#605E5C'>$($naTgt.Bracket)</span>" } else { "" })) } else { "None" }
    $naGreet = Get-ExceptionGreeting -CQ $cq -Prefix NoAgent
    [void]$naParts.Add("<b>Action:</b> $(HtmlEncode $naAct)")
    [void]$naParts.Add("<b>Target:</b> $naTgtText")
    [void]$naParts.Add("<b>Greeting:</b> $naGreet")
    $rows += @{ Parameter = "Exception: No Agents"; Value = ($naParts -join "<br/>") }

    # --- Callback Eligible (expanded) ---
    $cbLines = @()
    $cbLines += "Enabled: " + ($(if ($cq.IsCallbackEnabled) { "Yes" } else { "No" }))
    if ($cq.IsCallbackEnabled) {
        if ($cq.WaitTimeBeforeOfferingCallbackInSecond) {
            $cbLines += "After $($cq.WaitTimeBeforeOfferingCallbackInSecond) seconds"
        }
        if ($cq.NumberOfCallsInQueueBeforeOfferingCallback) {
            $cbLines += "More than $($cq.NumberOfCallsInQueueBeforeOfferingCallback) calls in queue"
        }
        if ($cq.CallToAgentRatioThresholdBeforeOfferingCallback) {
            $cbLines += "Calls to agent ratio greater than $($cq.CallToAgentRatioThresholdBeforeOfferingCallback)"
        }
    }
    $rows += @{ Parameter = "Callback Eligible"; Value = ($cbLines -join "<br/>") }

    # --- Callback Message (TTS / AudioFile / None) ---
    $cbMsg = "None"
    if ($cq.CallbackOfferTextToSpeechPrompt) {
        $cbMsg = "TTS: $(HtmlEncode $cq.CallbackOfferTextToSpeechPrompt)"
    } elseif ($cq.CallbackOfferAudioFilePromptFileName) {
        $cbMsg = "AudioFile: $(HtmlEncode $cq.CallbackOfferAudioFilePromptFileName)"
    }
    $rows += @{ Parameter = "Callback Message"; Value = $cbMsg }

    # --- Callback Key (DTMF) ---
    $cbKey = "Not configured"
    if ($cq.CallbackRequestDtmf) {
        $cbKey = HtmlEncode (Get-DtmfLabel ([string]$cq.CallbackRequestDtmf))
    }
    $rows += @{ Parameter = "Callback Key"; Value = $cbKey }

    # --- Callback Fail Notification (friendly) ---
    if ($cq.CallbackEmailNotificationTarget) {
        $cbRes = Resolve-Target $cq.CallbackEmailNotificationTarget
        $cbVal = if ($cbRes.Friendly) {
            $tmp = HtmlEncode $cbRes.Friendly
            if ($cbRes.Bracket) { $tmp += " <span style='color:#605E5C'>$($cbRes.Bracket)</span>" }
            $tmp
        } else { "None" }
        $rows += @{ Parameter = "Callback Fail Notification"; Value = $cbVal }
    } else {
        $rows += @{ Parameter = "Callback Fail Notification"; Value = "None" }
    }

    # --- Authorised Users (DisplayName-first) ---
    if ($cq.AuthorizedUsers) {
        $authLines = $cq.AuthorizedUsers | ForEach-Object { HtmlEncode (Resolve-User $_) }
        $rows += @{ Parameter = "Authorised Users"; Value = ($authLines -join "<br/>") }
    } else {
        $rows += @{ Parameter = "Authorised Users"; Value = "None" }
    }

    $sections += (New-Table $cq.Name $rows)
}

# ==============================
# HTML output
# ==============================
$style = @"
<style>
body { font-family: "Segoe UI Variable","Segoe UI",Arial,sans-serif; font-size: 13px; color: #323130; margin: 20px; }
h1 { margin-bottom: 10px; }
h2 { margin-top: 28px; }

table { border-collapse: collapse; width: 100%; font-size: 12px; table-layout: auto; }
th, td { border: 1px solid #ddd; padding: 8px; vertical-align: top; word-wrap: break-word; }
th { background: #f3f3f3; text-align: left; font-weight: 600; }
td:first-child { width: 260px; white-space: nowrap; }
tr:nth-child(even) { background: #fafafa; }
</style>
"@

@"
<html>
<head>
<meta charset="utf-8"/>
<title>Teams Call Queue Documentation</title>
$style
</head>
<body>
<h1>Teams Call Queue Documentation</h1>
<p>Generated: $(Get-Date)</p>
$($sections -join "`n")
</body>
</html>
"@ | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host "Report generated: $OutputPath" -ForegroundColor Green
if ($Open) { Invoke-Item $OutputPath }