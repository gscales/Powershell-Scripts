<#
.SYNOPSIS
    Creates a meeting in an Exchange Online calendar without sending meeting requests to the
    attendees.

.DESCRIPTION
    Creates the meeting through Microsoft Graph as a draft (isDraft = true), which stops Exchange
    sending meeting requests. It then sets the EventDraft and IntendedEventDraft properties to
    false so Outlook shows it as a normal meeting rather than a draft.

    The attendees are listed on the meeting but receive nothing, and the meeting is not added to
    their calendars. Rooms and equipment added with -ResourceAttendees are not booked, because a
    resource mailbox books itself when it processes a meeting request.

.PARAMETER Subject
    The meeting subject. You are prompted for it if you leave it out.

.PARAMETER StartTime
    When the meeting starts, in the time zone set by -TimeZone. You are prompted for it if you
    leave it out. Enter it as yyyy-MM-dd HH:mm (for example 2026-10-06 09:00): PowerShell reads
    a date such as 6/10/2026 in US month/day order, whatever your regional settings.

.PARAMETER EndTime
    When the meeting ends, in the time zone set by -TimeZone. You are prompted for it if you
    leave it out.

.PARAMETER TimeZone
    The time zone of StartTime and EndTime, as a Windows or IANA name such as
    'E. Australia Standard Time' or 'Australia/Brisbane' (Windows PowerShell 5.1 accepts Windows
    names only). Defaults to this computer's time zone. Date/time values that already carry a
    time zone, such as the output of Get-Date, are converted to this time zone.

.PARAMETER RequiredAttendees
    Email addresses of required attendees.

.PARAMETER OptionalAttendees
    Email addresses of optional attendees.

.PARAMETER ResourceAttendees
    Email addresses of rooms or equipment.

.PARAMETER Location
    The meeting location.

.PARAMETER Body
    The meeting body, as plain text unless -BodyAsHtml is used.

.PARAMETER BodyAsHtml
    Treats -Body as HTML.

.PARAMETER Mailbox
    The organizer's mailbox, as a UPN or user ID. Leave it out to use the signed-in user's
    calendar. Required with app-only authentication.

.EXAMPLE
    .\New-SilentMeeting.ps1 -Subject 'Quarterly planning' -StartTime '2026-10-06 09:00' -EndTime '2026-10-06 10:00' -RequiredAttendees alex@contoso.com, sam@contoso.com -OptionalAttendees jo@contoso.com -ResourceAttendees boardroom@contoso.com

    Creates the meeting in the signed-in user's calendar.

.EXAMPLE
    .\New-SilentMeeting.ps1 -RequiredAttendees alex@contoso.com

    Prompts for Subject, StartTime and EndTime.

.EXAMPLE
    Connect-MgGraph -ClientId $appId -TenantId $tenantId -CertificateThumbprint $thumbprint -NoWelcome
    .\New-SilentMeeting.ps1 -Mailbox organizer@contoso.com -Subject 'Contract review' -StartTime '2026-10-07 14:00' -EndTime '2026-10-07 14:30' -TimeZone 'AUS Eastern Standard Time' -RequiredAttendees alex@contoso.com

    Creates the meeting in another user's calendar using app-only authentication.

.OUTPUTS
    PSCustomObject with the meeting's Id, Subject, StartTime, EndTime, TimeZone, Attendees and
    WebLink.

.NOTES
    Microsoft Graph permissions:
        Delegated, own calendar       Calendars.ReadWrite
        Delegated, another calendar   Calendars.ReadWrite.Shared, plus delegate access to it
        Application                   Calendars.ReadWrite

    If there is no Microsoft Graph connection, the script signs in interactively with the
    delegated permissions above. For unattended use, run Connect-MgGraph first.
#>
[CmdletBinding(SupportsShouldProcess)]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, HelpMessage = 'Meeting subject')]
    [ValidateNotNullOrEmpty()]
    [string]$Subject,

    [Parameter(Mandatory, HelpMessage = 'Start date and time as yyyy-MM-dd HH:mm, for example 2026-10-06 09:00')]
    [datetime]$StartTime,

    [Parameter(Mandatory, HelpMessage = 'End date and time as yyyy-MM-dd HH:mm, for example 2026-10-06 10:00')]
    [datetime]$EndTime,

    [ValidateNotNullOrEmpty()]
    [string]$TimeZone = [TimeZoneInfo]::Local.Id,

    [ValidatePattern('^[^@\s]+@[^@\s]+$')]
    [string[]]$RequiredAttendees,

    [ValidatePattern('^[^@\s]+@[^@\s]+$')]
    [string[]]$OptionalAttendees,

    [ValidatePattern('^[^@\s]+@[^@\s]+$')]
    [string[]]$ResourceAttendees,

    [string]$Location,

    [string]$Body,

    [switch]$BodyAsHtml,

    [string]$Mailbox
)

$ErrorActionPreference = 'Stop'

function ConvertTo-ZoneTime {
    param(
        [datetime]$Value,
        [TimeZoneInfo]$Zone
    )

    # A value without a time zone (typed or prompted for) is already a wall-clock time in $Zone.
    if ($Value.Kind -eq [DateTimeKind]::Unspecified) {
        return $Value
    }
    [TimeZoneInfo]::ConvertTime($Value, $Zone)
}

# --- Meeting times -----------------------------------------------------------------------------
try {
    $zone = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZone)
}
catch {
    throw "Unknown time zone '$TimeZone'. Run [TimeZoneInfo]::GetSystemTimeZones() to list valid names."
}

$start = ConvertTo-ZoneTime -Value $StartTime -Zone $zone
$end   = ConvertTo-ZoneTime -Value $EndTime -Zone $zone
if ($end -le $start) {
    throw 'EndTime must be later than StartTime.'
}

# --- Attendees ---------------------------------------------------------------------------------
# Each address is added once. An address given for more than one type keeps the first of
# required, optional, resource.
$attendees = [System.Collections.Generic.List[hashtable]]::new()
$added     = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$groups    = @(
    @{ Type = 'required'; Addresses = $RequiredAttendees }
    @{ Type = 'optional'; Addresses = $OptionalAttendees }
    @{ Type = 'resource'; Addresses = $ResourceAttendees }
)
foreach ($group in $groups) {
    foreach ($address in $group.Addresses) {
        if ($added.Add($address)) {
            $attendees.Add(@{ emailAddress = @{ address = $address }; type = $group.Type })
        }
    }
}

# --- Request body ------------------------------------------------------------------------------
$meetingRequest = @{
    subject       = $Subject
    start         = @{ dateTime = $start.ToString('s'); timeZone = $zone.Id }
    end           = @{ dateTime = $end.ToString('s'); timeZone = $zone.Id }
    attendees     = $attendees
    isDraft       = $true                          # a draft meeting sends no meeting requests
    transactionId = [guid]::NewGuid().ToString()   # stops a retried POST creating a duplicate
}
if ($Location) {
    $meetingRequest.location = @{ displayName = $Location }
}
if ($Body) {
    $contentType = if ($BodyAsHtml) { 'html' } else { 'text' }
    $meetingRequest.body = @{ contentType = $contentType; content = $Body }
}

$mailboxSegment = if ($Mailbox) { "users/$Mailbox" } else { 'me' }
$eventsUri      = "https://graph.microsoft.com/v1.0/$mailboxSegment/events"
$organizer      = if ($Mailbox) { $Mailbox } else { 'the signed-in user' }

if (-not $PSCmdlet.ShouldProcess("calendar of $organizer", "Create meeting '$Subject' without sending meeting requests")) {
    return
}

# --- Connection --------------------------------------------------------------------------------
$context = Get-MgContext
if (-not $context) {
    $scopes = if ($Mailbox) { 'Calendars.ReadWrite', 'Calendars.ReadWrite.Shared' } else { 'Calendars.ReadWrite' }
    Connect-MgGraph -Scopes $scopes -NoWelcome
    $context = Get-MgContext
}
if ($context.AuthType -eq 'AppOnly' -and -not $Mailbox) {
    throw 'Use -Mailbox to name the organizer. An app-only connection has no signed-in user.'
}

# --- Create the meeting as a draft -------------------------------------------------------------
$createJson = $meetingRequest | ConvertTo-Json -Depth 10
$created    = Invoke-MgGraphRequest -Method POST -Uri $eventsUri -Body $createJson -ContentType 'application/json'

if (-not $created.isDraft) {
    Write-Warning "Meeting '$Subject' was not created as a draft, so meeting requests may have been sent."
}

# --- Clear the draft flags ---------------------------------------------------------------------
# Until these are false, Outlook shows the meeting as a draft. Both are string-named properties
# in PSETID_Appointment.
$psetidAppointment = '{00062002-0000-0000-C000-000000000046}'
$clearDraftJson = @{
    singleValueExtendedProperties = @(
        @{ id = "Boolean $psetidAppointment Name EventDraft"; value = 'false' }
        @{ id = "Boolean $psetidAppointment Name IntendedEventDraft"; value = 'false' }
    )
} | ConvertTo-Json -Depth 5

try {
    $meeting = Invoke-MgGraphRequest -Method PATCH -Uri "$eventsUri/$($created.id)" -Body $clearDraftJson -ContentType 'application/json'
}
catch {
    throw "Meeting '$Subject' was created (id $($created.id)), but its draft flags could not be cleared, so Outlook shows it as a draft. $($_.Exception.Message)"
}

[pscustomobject]@{
    Id        = $meeting.id
    Subject   = $meeting.subject
    StartTime = $start
    EndTime   = $end
    TimeZone  = $zone.Id
    Attendees = @($attendees | ForEach-Object { '{0} ({1})' -f $_.emailAddress.address, $_.type })
    WebLink   = $meeting.webLink
}