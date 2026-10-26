<#
.SYNOPSIS
    Finds applications that hold EWS permissions but no mailbox-level Microsoft Graph
    permissions (Mail, Calendars, Contacts, Mailbox import/export), are still
    authenticating to Exchange Online, and are NOT in EwsAllowedAppIDs.

.DESCRIPTION
    1. Reads the tenant's EwsAllowedAppIDs list (Exchange Online PowerShell).
    2. Finds every service principal granted EWS access to Office 365 Exchange Online:
         - full_access_as_app        (application permission / app role assignment)
         - EWS.AccessAsUser.All      (delegated permission / OAuth2 grant)
    3. Drops allow-listed apps.
    4. Checks granted Microsoft Graph permissions (app roles + delegated scopes).
       Only mailbox-level Graph permissions count as evidence the app has moved off EWS:
         Mail.*  Calendars.*  Contacts.*  MailboxItem.*  MailboxFolder.*
       (MailboxItem.* / MailboxFolder.* are the Graph mailbox import/export permissions.)
       An app with any of these is skipped. Anything else (User.Read, openid, Directory.*,
       Files.*, MailboxSettings.* ...) does NOT exclude the app; it is shown in the
       OtherGraph column for context. Override the match with -MailboxGraphPattern.
    5. Pre-filters on the service principal sign-in activity report (one cheap lookup per
       app) so apps that haven't signed in to anything in the window skip the slow
       sign-in log queries.
    6. Looks up the most recent sign-in to the Exchange Online resource
       (resourceId 00000002-0000-0ff1-ce00-000000000000) in the Entra sign-in logs.
    7. Reports apps that are EWS-only, not allow-listed, and (by default) still signing in.

    PERFORMANCE
    Per-app lookups are sent as Graph JSON batches (20 requests per HTTP call) instead of
    one call each, and the expensive steps run only on what survives the cheap ones:
    allow list -> Graph mailbox permissions -> sign-in activity report -> sign-in logs.
    Throttled requests (429/503/504) are retried using Retry-After with backoff.

.REQUIREMENTS
    Microsoft.Graph.Authentication and ExchangeOnlineManagement modules.
    Graph scopes: Application.Read.All, Directory.Read.All, AuditLog.Read.All
    Entra ID P1/P2 for sign-in log / sign-in activity API access (30-day retention).

.EXAMPLE
    .\Find-EwsOnlyApps.ps1
.EXAMPLE
    .\Find-EwsOnlyApps.ps1 -Days 7 -IncludeNoSignIns -CsvPath .\EwsOnlyApps.csv
.EXAMPLE
    # Cleanup report: EWS-only apps (allow-listed or not) with no Exchange sign-ins in 30 days
    .\Find-EwsOnlyApps.ps1 -UnusedOnly -CsvPath .\UnusedEwsApps.csv
.EXAMPLE
    # Target a specific tenant (closes any session left open to another tenant)
    .\Find-EwsOnlyApps.ps1 -TenantId contoso.onmicrosoft.com -UserPrincipalName admin@contoso.onmicrosoft.com
.EXAMPLE
    # Avoid the Windows sign-in broker if Connect-ExchangeOnline throws the WAM/RuntimeBroker error
    .\Find-EwsOnlyApps.ps1 -DisableWAM
.EXAMPLE
    # Fastest: skip the sign-in log queries, use only the activity report (not EXO-specific)
    .\Find-EwsOnlyApps.ps1 -ActivityReportOnly
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 30)]
    [int]    $Days = 30,
    # Graph permissions that indicate the app has a Graph mailbox path (and so is skipped)
    [string] $MailboxGraphPattern = '^(Mail|Calendars|Contacts|MailboxItem|MailboxFolder)\.',
    [switch] $IncludeNoSignIns,     # also report matching apps with no Exchange sign-ins in the window
    [switch] $UnusedOnly,           # report ONLY EWS-only apps with no Exchange sign-ins (cleanup candidates)
    [switch] $NoActivityPrefilter,  # don't use the sign-in activity report to skip idle apps
    [switch] $ActivityReportOnly,   # skip sign-in log queries entirely (fast, but not resource-specific)
    [string] $CsvPath,
    [string] $TenantId,             # tenant GUID or domain; forces both connections to this tenant
    [string] $UserPrincipalName,    # account to sign in to Exchange Online with (defaults to the Graph account)
    [switch] $DisableWAM            # skip the Windows sign-in broker (WAM) up front
)

$ErrorActionPreference = 'Stop'

$ExoAppId      = '00000002-0000-0ff1-ce00-000000000000'   # Office 365 Exchange Online
$GraphAppId    = '00000003-0000-0000-c000-000000000000'   # Microsoft Graph
$EwsAppRole    = 'full_access_as_app'
$EwsScope      = 'EWS.AccessAsUser.All'
$MicrosoftOrgs = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47'
$EventTypes    = 'servicePrincipal', 'nonInteractiveUser', 'interactiveUser'
$BatchSize     = 20     # Graph JSON batch limit
$MaxRetries    = 6

$sw = [System.Diagnostics.Stopwatch]::StartNew()
function Write-Step([string]$Text) { Write-Host ("[{0:mm\:ss}] {1}" -f $sw.Elapsed, $Text) -ForegroundColor Cyan }

# ---------------------------------------------------------------- helpers

function Invoke-Graph([string]$Uri) {
    Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
}

function Get-GraphPaged([string]$Uri) {
    do {
        $r = Invoke-Graph $Uri
        if ($r.value) { $r.value }
        $Uri = $r.'@odata.nextLink'
    } while ($Uri)
}

function Format-Filter([string]$Path, [string]$Filter, [string]$Extra = '') {
    "$Path`?`$filter=" + [uri]::EscapeDataString($Filter) + $Extra
}

function Add-ToMap([hashtable]$Map, [string]$Key, [string]$Value) {
    if (-not $Map.ContainsKey($Key)) { $Map[$Key] = [System.Collections.Generic.List[string]]::new() }
    if (-not $Map[$Key].Contains($Value)) { $Map[$Key].Add($Value) }
}

# Graph timestamps arrive as ISO strings or DateTime depending on deserialisation; normalise to UTC
function Get-UtcDate($Value) {
    if ($null -eq $Value -or $Value -eq '') { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    [datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
}

<#
  Sends GET requests as Graph JSON batches.
  $Requests: list of @{ Id = '<unique key>'; Url = '/relative/url?...' }  (no version prefix)
  Returns hashtable Id -> response object (status, body). Retries 429/503/504.
#>
function Invoke-GraphBatch {
    param(
        [System.Collections.IList] $Requests,
        [ValidateSet('v1.0', 'beta')] [string] $Version = 'v1.0',
        [string] $Activity = 'Graph batch'
    )
    $results = @{}
    if (-not $Requests -or $Requests.Count -eq 0) { return $results }

    $total   = $Requests.Count
    $queue   = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $Requests) { $queue.Add($r) }
    $attempt = 0

    while ($queue.Count -gt 0) {
        $retry = [System.Collections.Generic.List[object]]::new()
        $wait  = 0

        for ($i = 0; $i -lt $queue.Count; $i += $BatchSize) {
            $chunk = $queue.GetRange($i, [Math]::Min($BatchSize, $queue.Count - $i))
            $byId  = @{}
            $reqs  = foreach ($c in $chunk) {
                $byId[$c.Id] = $c
                @{ id = $c.Id; method = 'GET'; url = $c.Url }
            }
            $body = @{ requests = @($reqs) } | ConvertTo-Json -Depth 5 -Compress
            $resp = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/$Version/`$batch" `
                        -Body $body -ContentType 'application/json' -OutputType PSObject

            foreach ($r in $resp.responses) {
                if ($r.status -in 429, 503, 504) {
                    $retry.Add($byId[$r.id])
                    $ra = $r.headers.'Retry-After'
                    if ($ra) { $wait = [Math]::Max($wait, [int]$ra) }
                }
                else { $results[$r.id] = $r }
            }
            Write-Progress -Activity $Activity -Status "$($results.Count) / $total" `
                -PercentComplete ([Math]::Min(100, 100 * $results.Count / $total))
        }

        $queue = $retry
        if ($queue.Count -gt 0) {
            $attempt++
            if ($attempt -gt $MaxRetries) { throw "$Activity : $($queue.Count) request(s) still throttled after $MaxRetries retries" }
            $sleep = [Math]::Max($wait, [Math]::Pow(2, $attempt))
            Write-Verbose "$Activity : $($queue.Count) throttled, retrying in $sleep s"
            Start-Sleep -Seconds $sleep
        }
    }
    Write-Progress -Activity $Activity -Completed
    $results
}

# Values from a batch response, following @odata.nextLink if the first page wasn't everything
function Get-BatchValues($Response) {
    if (-not $Response -or $Response.status -ne 200) { return }
    if ($Response.body.value) { $Response.body.value }
    $next = $Response.body.'@odata.nextLink'
    if ($next) { Get-GraphPaged $next }
}

# ---------------------------------------------------------------- connect

$GraphScopes = 'Application.Read.All', 'Directory.Read.All', 'AuditLog.Read.All'

# Accept a tenant GUID or any verified domain (e.g. contoso.onmicrosoft.com)
function Resolve-TenantId([string]$Value) {
    if (-not $Value) { return $null }
    if ($Value -match '^[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}$') { return $Value }
    $cfg = Invoke-RestMethod "https://login.microsoftonline.com/$Value/v2.0/.well-known/openid-configuration"
    ($cfg.issuer -split '/')[3]
}
$wantTenant = Resolve-TenantId $TenantId

# --- Microsoft Graph: reuse an existing session only if it's the right tenant and has the scopes
$ctx = Get-MgContext
$needGraph = (-not $ctx) -or
             ($wantTenant -and $ctx.TenantId -ne $wantTenant) -or
             ($ctx.AuthType -eq 'Delegated' -and @($GraphScopes | Where-Object { $_ -notin $ctx.Scopes }).Count -gt 0)
if ($ctx -and $needGraph) {
    Write-Host "  Disconnecting existing Graph session ($($ctx.Account), tenant $($ctx.TenantId))"
    Disconnect-MgGraph | Out-Null
}
if ($needGraph) {
    $gp = @{ Scopes = $GraphScopes; NoWelcome = $true }
    if ($wantTenant) { $gp.TenantId = $wantTenant }
    try { Connect-MgGraph @gp }
    catch {
        $opt = Get-Command Set-MgGraphOption -ErrorAction SilentlyContinue
        if ($DisableWAM -or -not ($opt -and $opt.Parameters.ContainsKey('DisableLoginByWAM'))) { throw }
        Write-Warning "Graph sign-in failed ($($_.Exception.Message)); retrying without the Windows broker (WAM)."
        Set-MgGraphOption -DisableLoginByWAM $true
        Connect-MgGraph @gp
    }
}
$ctx      = Get-MgContext
$tenantId = $ctx.TenantId

# --- Exchange Online: must be the same tenant as Graph
$exo = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1
if ($exo -and $exo.TenantID -and $exo.TenantID -ne $tenantId) {
    Write-Host "  Disconnecting Exchange Online session for tenant $($exo.TenantID) (Graph is $tenantId)"
    Disconnect-ExchangeOnline -Confirm:$false
    $exo = $null
}
if (-not $exo) {
    $xp = @{ ShowBanner = $false }
    # Sign in to EXO as the same account as Graph so both land in the same tenant
    if ($UserPrincipalName)                         { $xp.UserPrincipalName = $UserPrincipalName }
    elseif ($ctx.AuthType -eq 'Delegated' -and $ctx.Account) { $xp.UserPrincipalName = $ctx.Account }

    $hasWamSwitch = (Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')
    if ($DisableWAM -and $hasWamSwitch) { $xp.DisableWAM = $true }

    try { Connect-ExchangeOnline @xp }
    catch {
        if ($hasWamSwitch -and -not $xp.DisableWAM) {
            # Known ExchangeOnlineManagement issue: the WAM broker throws NullReferenceException
            # (RuntimeBroker..ctor) when it can't get a window handle from the console host.
            Write-Warning "Exchange Online sign-in failed ($($_.Exception.Message)); retrying with -DisableWAM."
            $xp.DisableWAM = $true
            Connect-ExchangeOnline @xp
        }
        elseif (-not $hasWamSwitch) {
            throw "Connect-ExchangeOnline failed: $($_.Exception.Message). If this is the WAM/RuntimeBroker error, " +
                  "update ExchangeOnlineManagement to 3.7.2 or later (adds -DisableWAM): Update-Module ExchangeOnlineManagement"
        }
        else { throw }
    }
    $exo = Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1
}
if ($exo.TenantID -and $exo.TenantID -ne $tenantId) {
    throw "Tenant mismatch: Graph is connected to $tenantId but Exchange Online to $($exo.TenantID). " +
          "Re-run with -TenantId and -UserPrincipalName for the tenant you want."
}
Write-Host "  Tenant $tenantId  (Graph: $($ctx.Account); Exchange: $($exo.UserPrincipalName))"
$sinceDt  = (Get-Date).ToUniversalTime().AddDays(-$Days)
$since    = $sinceDt.ToString('yyyy-MM-ddTHH:mm:ssZ')

# ---------------------------------------------------------------- 1. EWS allow list

Write-Step 'Reading EwsAllowedAppIDs...'
$orgCfg   = Get-OrganizationConfig -RetrieveEwsOperationAccessPolicy
$allowSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($entry in @($orgCfg.EwsAllowedAppIDs)) {
    foreach ($id in ($entry -split '[,;\s]+')) { if ($id) { [void]$allowSet.Add($id.Trim()) } }
}
Write-Host ("  EwsEnabled = {0}; {1} app ID(s) in allow list" -f $orgCfg.EwsEnabled, $allowSet.Count)

# ---------------------------------------------------------------- 2. Resource service principals

$spSelect = '&$select=id,appId,displayName,appRoles,oauth2PermissionScopes'
$exoSp   = (Invoke-Graph (Format-Filter 'https://graph.microsoft.com/v1.0/servicePrincipals' "appId eq '$ExoAppId'" $spSelect)).value |
           Select-Object -First 1
$graphSp = (Invoke-Graph (Format-Filter 'https://graph.microsoft.com/v1.0/servicePrincipals' "appId eq '$GraphAppId'" $spSelect)).value |
           Select-Object -First 1

$ewsRoleIds  = @($exoSp.appRoles | Where-Object value -eq $EwsAppRole | ForEach-Object id)
$graphRoles  = @{}; foreach ($r in $graphSp.appRoles)               { $graphRoles[$r.id]  = $r.value }
$graphScopes = @{}; foreach ($s in $graphSp.oauth2PermissionScopes) { $graphScopes[$s.id] = $s.value }

# ---------------------------------------------------------------- 3. Who holds EWS permissions

Write-Step 'Finding service principals with EWS permissions...'
$ewsHolders = @{}

$uri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($exoSp.id)/appRoleAssignedTo?`$top=999&`$select=principalId,principalType,appRoleId"
foreach ($a in Get-GraphPaged $uri) {
    if ($a.principalType -eq 'ServicePrincipal' -and $a.appRoleId -in $ewsRoleIds) {
        Add-ToMap $ewsHolders $a.principalId "$EwsAppRole (Application)"
    }
}

# In large tenants this can be many per-user (Principal) grants, so keep the payload small
$uri = Format-Filter 'https://graph.microsoft.com/v1.0/oauth2PermissionGrants' "resourceId eq '$($exoSp.id)'" '&$select=clientId,scope,consentType'
foreach ($g in Get-GraphPaged $uri) {
    if (($g.scope -split '\s+') -contains $EwsScope) {
        Add-ToMap $ewsHolders $g.clientId "$EwsScope (Delegated, $($g.consentType))"
    }
}
Write-Host "  $($ewsHolders.Count) service principal(s) hold EWS permissions"

# ---------------------------------------------------------------- 4. SP details (batched) + allow list

Write-Step 'Resolving service principals...'
$reqs = foreach ($spId in $ewsHolders.Keys) {
    @{ Id = $spId; Url = "/servicePrincipals/$($spId)?`$select=id,appId,displayName,appOwnerOrganizationId,accountEnabled" }
}
$spResp = Invoke-GraphBatch -Requests @($reqs) -Activity 'Resolving service principals'

$candidates = @{}   # spId -> sp
foreach ($spId in $ewsHolders.Keys) {
    $r = $spResp[$spId]
    if (-not $r -or $r.status -ne 200) { Write-Warning "Could not read service principal $spId (status $($r.status))"; continue }
    # Unused report keeps allow-listed apps: an allow-listed app nobody uses is a cleanup candidate too
    if ($UnusedOnly -or -not $allowSet.Contains($r.body.appId)) { $candidates[$spId] = $r.body }
}
if ($UnusedOnly) { Write-Host "  $($candidates.Count) to check (allow-listed apps included for the unused report)" }
else             { Write-Host "  $($candidates.Count) not in the allow list" }

# ---------------------------------------------------------------- 5. Graph permissions (batched)

Write-Step 'Checking Graph permissions...'
$reqs = foreach ($spId in $candidates.Keys) {
    @{ Id = "ara|$spId"; Url = "/servicePrincipals/$spId/appRoleAssignments?`$select=resourceId,appRoleId" }
    @{ Id = "grt|$spId"; Url = (Format-Filter '/oauth2PermissionGrants' "clientId eq '$spId'" '&$select=resourceId,scope') }
}
$permResp = Invoke-GraphBatch -Requests @($reqs) -Activity 'Checking Graph permissions'

$graphOtherBySp = @{}
foreach ($spId in @($candidates.Keys)) {
    $mailbox = [System.Collections.Generic.List[string]]::new()
    $other   = [System.Collections.Generic.List[string]]::new()

    foreach ($a in Get-BatchValues $permResp["ara|$spId"]) {
        if ($a.resourceId -ne $graphSp.id) { continue }
        $name = $graphRoles[$a.appRoleId]
        if ($name -match $MailboxGraphPattern) { $mailbox.Add("$name (Application)") } else { $other.Add("$name (Application)") }
    }
    foreach ($g in Get-BatchValues $permResp["grt|$spId"]) {
        if ($g.resourceId -ne $graphSp.id) { continue }
        foreach ($scope in ($g.scope -split '\s+' | Where-Object { $_ })) {
            if ($scope -match $MailboxGraphPattern) { $mailbox.Add("$scope (Delegated)") } else { $other.Add("$scope (Delegated)") }
        }
    }

    if ($mailbox.Count -gt 0) { $candidates.Remove($spId); continue }   # has a Graph mailbox path -> migrated
    $graphOtherBySp[$spId] = ($other | Select-Object -Unique) -join '; '
}
Write-Host "  $($candidates.Count) have no Graph mailbox permissions"

# Everything that could appear in the report. $candidates is narrowed below to the apps that
# still need sign-in log queries; the unused report needs the ones dropped along the way.
$reportSet = $candidates.Clone()

# ---------------------------------------------------------------- 6. Activity pre-filter (batched, beta)

$activity    = @{}   # appId -> last sign-in (any resource) per the activity report
$activityObj = @{}   # appId -> full activity record (per-flow timestamps, used to prune step 7)
if (-not $NoActivityPrefilter -and $candidates.Count -gt 0) {
    Write-Step 'Checking service principal sign-in activity...'
    $reqs = foreach ($sp in $candidates.Values) {
        @{ Id = $sp.appId; Url = (Format-Filter '/reports/servicePrincipalSignInActivities' "appId eq '$($sp.appId)'") }
    }
    $actResp = Invoke-GraphBatch -Requests @($reqs) -Version beta -Activity 'Sign-in activity report'

    foreach ($spId in @($candidates.Keys)) {
        $appId = $candidates[$spId].appId
        $r     = $actResp[$appId]
        if (-not $r -or $r.status -ne 200) { continue }               # report unavailable -> keep, check logs
        $rec   = $r.body.value | Select-Object -First 1
        $last  = $rec.lastSignInActivity.lastSignInDateTime
        $activity[$appId]    = $last
        $activityObj[$appId] = $rec
        # Idle for the whole window: no sign-in log query needed. In the normal report it's
        # dropped; in the unused report it's already settled as unused (kept in $reportSet).
        if (($UnusedOnly -or -not $IncludeNoSignIns) -and (-not $last -or (Get-UtcDate $last) -lt $sinceDt)) {
            $candidates.Remove($spId)
        }
    }
    Write-Host "  $($candidates.Count) signed in to something in the last $Days day(s)"
    if ($UnusedOnly) { Write-Host "  $($reportSet.Count - $candidates.Count) had no sign-ins at all (unused without querying logs)" }
}

# ---------------------------------------------------------------- 7. Exchange sign-ins (batched, beta)

<#
  Optimisations in this step:
   a. Only query the sign-in flows the app's EWS permission can actually produce:
        full_access_as_app    -> servicePrincipal
        EWS.AccessAsUser.All  -> nonInteractiveUser + interactiveUser
   b. Drop flows the activity report says haven't been used in the window
      (applicationAuthenticationClientSignInActivity / delegatedClientSignInActivity).
   c. $top=1 and $select of only the fields reported, so each response is one small row.
      Support for $select and the resourceId filter is probed ONCE up front rather than
      failing per request.
   d. Search newest-first in time slices (last 1 day, then 7, then the rest). Most active
      apps are found in the first, cheapest slice and are not queried again; only
      unresolved app/flow pairs move on to the older, more expensive slices.
#>
$lastExo = @{}   # appId -> sign-in object
if (-not $ActivityReportOnly -and $candidates.Count -gt 0) {
    Write-Step 'Querying Exchange Online sign-ins...'
    $signInPath = 'https://graph.microsoft.com/beta/auditLogs/signIns'
    $selFields  = 'createdDateTime,userPrincipalName,ipAddress,status,resourceId'

    # --- (a) + (b): which app/flow pairs are worth querying
    $pending = [System.Collections.Generic.List[object]]::new()
    foreach ($kv in $candidates.GetEnumerator()) {
        $appId = $kv.Value.appId
        $perms = $ewsHolders[$kv.Key] -join ' '
        $rec   = $activityObj[$appId]

        # Flows this app's EWS permissions can produce
        $permFlows = [ordered]@{}
        if ($perms -match [regex]::Escape($EwsAppRole)) {
            $permFlows['app'] = @{ Types = @('servicePrincipal')
                                   Ts    = $rec.applicationAuthenticationClientSignInActivity.lastSignInDateTime }
        }
        if ($perms -match [regex]::Escape($EwsScope)) {
            $permFlows['delegated'] = @{ Types = @('nonInteractiveUser', 'interactiveUser')
                                         Ts    = $rec.delegatedClientSignInActivity.lastSignInDateTime }
        }

        # Skip a flow only when the activity report POSITIVELY shows it was last used before
        # the window. A missing per-flow timestamp is common (the report doesn't always fill
        # these in) and must not be read as "unused".
        $flows = @()
        foreach ($pf in $permFlows.Values) {
            $staleByReport = (-not $NoActivityPrefilter) -and $pf.Ts -and ((Get-UtcDate $pf.Ts) -lt $sinceDt)
            if (-not $staleByReport) { $flows += $pf.Types }
        }
        # Safety net: the app passed the overall activity check, so never prune it to nothing
        if ($flows.Count -eq 0) { $flows = @($permFlows.Values | ForEach-Object { $_.Types }) }

        foreach ($t in $flows) { $pending.Add([pscustomobject]@{ AppId = $appId; Type = $t }) }
    }
    Write-Host ("  {0} app/flow quer{1} needed (vs {2} without pruning)" -f $pending.Count,
        $(if ($pending.Count -eq 1) { 'y' } else { 'ies' }), ($candidates.Count * $EventTypes.Count))

    if ($pending.Count -gt 0) {
        # --- (c) probe once for resourceId filter and $select support
        function Test-SignInUrl([string]$Url) {
            try { [void](Invoke-Graph $Url); $true }
            catch { if ($_.Exception.Message -match '400|BadRequest') { $false } else { $true } }
        }
        $p     = $pending[0]
        $pBase = "createdDateTime ge $since and appId eq '$($p.AppId)' and signInEventTypes/any(x:x eq '$($p.Type)')"
        $useResourceFilter = Test-SignInUrl (Format-Filter $signInPath "$pBase and resourceId eq '$ExoAppId'" '&$top=1')
        $pFilter   = if ($useResourceFilter) { "$pBase and resourceId eq '$ExoAppId'" } else { $pBase }
        $useSelect = Test-SignInUrl (Format-Filter $signInPath $pFilter "&`$top=1&`$select=$selFields")

        $top   = if ($useResourceFilter) { 1 } else { 50 }    # client-side filtering needs a few rows
        $extra = "&`$top=$top" + $(if ($useSelect) { "&`$select=$selFields" } else { '' })
        Write-Verbose "resourceId filter: $useResourceFilter; `$select: $useSelect"
        if (-not $useResourceFilter) { Write-Warning 'resourceId filter not supported; filtering sign-ins client-side (slower).' }

        # --- (d) newest-first time slices
        $now        = (Get-Date).ToUniversalTime()
        $boundaries = @(@(1, 7, $Days) | Where-Object { $_ -le $Days } | Sort-Object -Unique)
        $upper      = $now

        foreach ($d in $boundaries) {
            if ($pending.Count -eq 0) { break }
            $lower = $now.AddDays(-$d)
            $range = "createdDateTime ge $($lower.ToString('yyyy-MM-ddTHH:mm:ssZ')) and createdDateTime lt $($upper.ToString('yyyy-MM-ddTHH:mm:ssZ'))"

            $reqs = foreach ($q in $pending) {
                $f = "$range and appId eq '$($q.AppId)' and signInEventTypes/any(x:x eq '$($q.Type)')"
                if ($useResourceFilter) { $f += " and resourceId eq '$ExoAppId'" }
                @{ Id = "$($q.AppId)|$($q.Type)"; Url = (Format-Filter '/auditLogs/signIns' $f $extra) }
            }
            $siResp = Invoke-GraphBatch -Requests @($reqs) -Version beta -Activity "Exchange sign-ins (last $d day(s))"

            $stillPending = [System.Collections.Generic.List[object]]::new()
            foreach ($q in $pending) {
                $r = $siResp["$($q.AppId)|$($q.Type)"]
                if (-not $r -or $r.status -ne 200) {
                    Write-Warning "Sign-in query for $($q.AppId) [$($q.Type)] failed (status $($r.status))"
                    continue
                }
                $s = $r.body.value | Where-Object { $useResourceFilter -or $_.resourceId -eq $ExoAppId } | Select-Object -First 1
                if (-not $s) { $stillPending.Add($q); continue }    # try the next (older) slice

                $cur = $lastExo[$q.AppId]
                if (-not $cur -or (Get-UtcDate $s.createdDateTime) -gt (Get-UtcDate $cur.createdDateTime)) {
                    $s | Add-Member -NotePropertyName EventType -NotePropertyValue $q.Type -Force
                    $lastExo[$q.AppId] = $s
                }
            }
            # An app already found in a newer slice can't get a later hit from an older one,
            # so drop its other flows too (e.g. delegated flows once app-only was found)
            $stillPending = @($stillPending | Where-Object { -not $lastExo.ContainsKey($_.AppId) })
            Write-Host ("  last {0,2} day(s): {1} app(s) found so far, {2} quer{3} carried to next slice" -f $d,
                $lastExo.Count, $stillPending.Count, $(if ($stillPending.Count -eq 1) { 'y' } else { 'ies' }))
            $pending = [System.Collections.Generic.List[object]]::new()
            foreach ($q in $stillPending) { $pending.Add($q) }
            $upper   = $lower
        }
    }
}

# ---------------------------------------------------------------- 8. Declared mailbox Graph perms (batched)

$declaredBySp = @{}
$outputSet = if ($UnusedOnly) { $reportSet } else { $candidates }
$homeApps  = @($outputSet.GetEnumerator() | Where-Object { $_.Value.appOwnerOrganizationId -eq $tenantId })
if ($homeApps.Count -gt 0) {
    $reqs = foreach ($kv in $homeApps) {
        @{ Id = $kv.Key; Url = (Format-Filter '/applications' "appId eq '$($kv.Value.appId)'" '&$select=requiredResourceAccess') }
    }
    $appResp = Invoke-GraphBatch -Requests @($reqs) -Activity 'Reading app registrations'
    foreach ($kv in $homeApps) {
        $appReg = Get-BatchValues $appResp[$kv.Key] | Select-Object -First 1
        $declaredBySp[$kv.Key] = ($appReg.requiredResourceAccess | Where-Object resourceAppId -eq $GraphAppId |
            ForEach-Object resourceAccess | ForEach-Object {
                if ($_.type -eq 'Role') { "$($graphRoles[$_.id]) (Application)" } else { $graphScopes[$_.id] }
            } | Where-Object { $_ -match $MailboxGraphPattern }) -join '; '
    }
}

# ---------------------------------------------------------------- 9. Build results

$results = foreach ($kv in $outputSet.GetEnumerator()) {
    $spId  = $kv.Key; $sp = $kv.Value
    $last  = $lastExo[$sp.appId]
    $anyTs = Get-UtcDate $activity[$sp.appId]
    $activeAnywhere = $anyTs -and $anyTs -ge $sinceDt

    if ($UnusedOnly) {
        if ($last) { continue }                                       # used Exchange in the window
        if ($ActivityReportOnly -and $activeAnywhere) { continue }    # report-only: any activity counts as used
    }
    elseif (-not $ActivityReportOnly -and -not $last -and -not $IncludeNoSignIns) { continue }

    $usage = if ($last)                                       { 'Exchange sign-ins' }
             elseif (-not $activityObj.ContainsKey($sp.appId)) { if ($ActivityReportOnly) { 'Unknown (no activity record)' } else { 'No Exchange sign-ins' } }
             elseif ($activeAnywhere)                          { if ($ActivityReportOnly) { 'Signs in (any resource)' } else { 'Signs in, but not to Exchange' } }
             else                                              { 'No sign-ins' }

    $origin = if ($sp.appOwnerOrganizationId -eq $tenantId) { 'This tenant' }
              elseif ($sp.appOwnerOrganizationId -in $MicrosoftOrgs) { 'Microsoft' }
              elseif ($sp.appOwnerOrganizationId) { 'Third-party' }
              else { 'Unknown' }

    [pscustomobject]@{
        DisplayName          = $sp.displayName
        AppId                = $sp.appId
        Origin               = $origin
        Enabled              = $sp.accountEnabled
        EwsPermissions       = $ewsHolders[$spId] -join '; '
        OtherGraph           = $graphOtherBySp[$spId]
        MailboxGraphDeclared = $declaredBySp[$spId]          # declared on app reg but not consented
        Usage                = $usage
        LastAnySignInUtc     = $anyTs                       # from the activity report; can predate the window
        LastExoSignInUtc     = if ($last) { Get-UtcDate $last.createdDateTime } else { $null }
        LastExoSignInLocal   = if ($last) { (Get-UtcDate $last.createdDateTime).ToLocalTime() } else { $null }
        LastSignInType       = $last.EventType
        LastSignInUser       = $last.userPrincipalName
        LastSignInIP         = $last.ipAddress
        LastSignInErrorCode  = $last.status.errorCode
        InEwsAllowList       = $allowSet.Contains($sp.appId)
    }
}

# ---------------------------------------------------------------- 10. Output

Write-Step 'Done.'
if ($UnusedOnly) {
    # Longest-idle first (never-seen apps at the top)
    $results = @($results | Sort-Object @{ Expression = { if ($_.LastAnySignInUtc) { $_.LastAnySignInUtc } else { [datetime]::MinValue } } }, DisplayName)
    Write-Host ("`n{0} app(s) with EWS but no Graph mailbox permissions and no {1} in the last {2} day(s):" -f $results.Count,
        $(if ($ActivityReportOnly) { 'sign-ins' } else { 'Exchange sign-ins' }), $Days) -ForegroundColor Yellow
    $cols = 'DisplayName', 'AppId', 'Origin', 'EwsPermissions', 'InEwsAllowList', 'Usage', 'LastAnySignInUtc'
}
else {
    $sortKey = if ($ActivityReportOnly) { 'LastAnySignInUtc' } else { 'LastExoSignInUtc' }
    $results = @($results | Sort-Object $sortKey -Descending)
    Write-Host ("`n{0} app(s) with EWS but no Graph mailbox permissions, not in the allow list{1}:" -f $results.Count,
        $(if ($IncludeNoSignIns) { '' }
          elseif ($ActivityReportOnly) { " with any sign-in in the last $Days day(s)" }
          else { " with Exchange sign-ins in the last $Days day(s)" })) -ForegroundColor Yellow
    $cols = if ($ActivityReportOnly) { 'DisplayName', 'AppId', 'Origin', 'EwsPermissions', 'OtherGraph', 'LastAnySignInUtc' }
            else { 'DisplayName', 'AppId', 'Origin', 'EwsPermissions', 'OtherGraph', 'LastExoSignInLocal', 'LastSignInType' }
}
$results | Format-Table $cols -AutoSize -Wrap

if ($CsvPath) {
    $results | Export-Csv -Path $CsvPath -NoTypeInformation
    Write-Host "Exported to $CsvPath"
}