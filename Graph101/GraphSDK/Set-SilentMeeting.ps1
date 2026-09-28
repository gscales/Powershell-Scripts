#Requires -Modules Microsoft.Graph.Authentication
<#
.SYNOPSIS
    Changes a meeting's time, or adds attendees to it, through Microsoft Graph without sending
    meeting updates or invitations.

.DESCRIPTION
    Changing a meeting's start or end time normally sends an update to every attendee, and adding
    attendees sends invitations to the new ones. To make the change without notifying anyone,
    the script:

      1. Sets PidLidFInvited to false if meeting requests have been sent, so the meeting is
         treated as never sent.
      2. Removes the attendees.
      3. Adds the attendees back, with any new ones, applies the new times, and saves the meeting
         as a draft (isDraft = true) so no meeting requests are sent.
      4. Sets EventDraft and IntendedEventDraft to false so Outlook shows a normal meeting rather
         than a draft.
      5. With -RestoreInvitedFlag, sets PidLidFInvited back to true.

    Because nothing is sent, attendees' copies of the meeting keep the old time, and new attendees
    don't get the meeting in their calendars. Rooms keep their booking at the old time, and new
    rooms aren't booked. Responses recorded on the organizer's copy may be reset, because the
    attendees are removed and added back.

    Only single-instance meetings are supported, and the time of an all-day meeting can't be
    changed.

.PARAMETER EventId
    The id of the meeting in the organizer's calendar, such as the Id returned by
    New-SilentMeeting.ps1. You are prompted for it if you leave it out.

.PARAMETER StartTime
    The new start, in the time zone set by -TimeZone. If you give only StartTime, the meeting
    keeps its length. Enter it as yyyy-MM-dd HH:mm (for example 2026-10-07 10:00): PowerShell
    reads a date such as 7/10/2026 in US month/day order, whatever your regional settings.

.PARAMETER EndTime
    The new end, in the time zone set by -TimeZone. If you give only EndTime, the start doesn't
    change.

.PARAMETER TimeZone
    The time zone of StartTime and EndTime, as a Windows or IANA name such as
    'E. Australia Standard Time' or 'Australia/Brisbane' (Windows PowerShell 5.1 accepts Windows
    names only). Defaults to this computer's time zone. Date/time values that already carry a
    time zone, such as the output of Get-Date, are converted to this time zone.

.PARAMETER AddRequiredAttendees
    Email addresses to add as required attendees. The existing attendees are kept.

.PARAMETER AddOptionalAttendees
    Email addresses to add as optional attendees.

.PARAMETER AddResourceAttendees
    Email addresses of rooms or equipment to add.

.PARAMETER RestoreInvitedFlag
    Sets PidLidFInvited back to true at the end, if it was true before, so the meeting is treated
    as sent again and later changes send updates as usual. Without it, the meeting stays marked
    as never sent.

.PARAMETER Mailbox
    The organizer's mailbox, as a UPN or user ID. Leave it out to use the signed-in user's
    calendar. Required with app-only authentication.

.EXAMPLE
    .\Set-SilentMeeting.ps1 -EventId $id -StartTime '2026-10-07 10:00'

    Moves the meeting to 10:00 on 7 October and keeps its length.

.EXAMPLE
    .\Set-SilentMeeting.ps1 -EventId $id -AddRequiredAttendees kim@contoso.com -AddOptionalAttendees lee@contoso.com -AddResourceAttendees projector@contoso.com

    Adds three attendees without sending them invitations.

.EXAMPLE
    $meeting = .\New-SilentMeeting.ps1 -Subject 'Handover' -StartTime '2026-10-06 09:00' -EndTime '2026-10-06 10:00' -RequiredAttendees alex@contoso.com
    .\Set-SilentMeeting.ps1 -EventId $meeting.Id -StartTime '2026-10-06 14:00' -EndTime '2026-10-06 15:30'

    Changes the time of a meeting created with New-SilentMeeting.ps1.

.EXAMPLE
    .\Set-SilentMeeting.ps1 -Mailbox organizer@contoso.com -EventId $id -StartTime '2026-10-08 09:00' -EndTime '2026-10-08 09:30' -TimeZone 'AUS Eastern Standard Time' -RestoreInvitedFlag

    Changes the time of a meeting in another user's calendar (for example with app-only
    authentication), then marks the meeting as sent again.

.OUTPUTS
    PSCustomObject with the meeting's Id, Subject, Changes, Attendees, PidLidFInvited and
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
    [Parameter(Mandatory, HelpMessage = 'Id of the meeting in the organizer''s calendar')]
    [ValidateNotNullOrEmpty()]
    [string]$EventId,

    [datetime]$StartTime,

    [datetime]$EndTime,

    [ValidateNotNullOrEmpty()]
    [string]$TimeZone = [TimeZoneInfo]::Local.Id,

    [ValidatePattern('^[^@\s]+@[^@\s]+$')]
    [string[]]$AddRequiredAttendees,

    [ValidatePattern('^[^@\s]+@[^@\s]+$')]
    [string[]]$AddOptionalAttendees,

    [ValidatePattern('^[^@\s]+@[^@\s]+$')]
    [string[]]$AddResourceAttendees,

    [switch]$RestoreInvitedFlag,

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

function ConvertFrom-GraphUtc {
    param([string]$Value)

    $parsed = [datetime]::Parse($Value, [cultureinfo]::InvariantCulture)
    [datetime]::SpecifyKind($parsed, [DateTimeKind]::Utc)
}

function Invoke-EventPatch {
    param(
        [string]$Uri,
        [hashtable]$Changes
    )

    $json = $Changes | ConvertTo-Json -Depth 10
    Invoke-MgGraphRequest -Method PATCH -Uri $Uri -Body $json -ContentType 'application/json'
}

# --- Check the input ---------------------------------------------------------------------------
$changingTimes   = $PSBoundParameters.ContainsKey('StartTime') -or $PSBoundParameters.ContainsKey('EndTime')
$addingAttendees = [bool]($AddRequiredAttendees -or $AddOptionalAttendees -or $AddResourceAttendees)
if (-not ($changingTimes -or $addingAttendees)) {
    throw 'Nothing to change. Use -StartTime, -EndTime or the -Add*Attendees parameters.'
}

try {
    $zone = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZone)
}
catch {
    throw "Unknown time zone '$TimeZone'. Run [TimeZoneInfo]::GetSystemTimeZones() to list valid names."
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

# --- Read the meeting --------------------------------------------------------------------------
$psetidAppointment = '{00062002-0000-0000-C000-000000000046}'
$fInvitedId        = "Boolean $psetidAppointment Id 0x8229"    # PidLidFInvited

$mailboxSegment = if ($Mailbox) { "users/$Mailbox" } else { 'me' }
$eventUri       = "https://graph.microsoft.com/v1.0/$mailboxSegment/events/$EventId"
$organizer      = if ($Mailbox) { $Mailbox } else { 'the signed-in user' }

$expand  = "singleValueExtendedProperties(`$filter=id eq '$fInvitedId')"
$meeting = Invoke-MgGraphRequest -Method GET -Uri ('{0}?$expand={1}' -f $eventUri, $expand) -Headers @{ Prefer = 'outlook.timezone="UTC"' }
$subject = $meeting.subject

if (-not $meeting.isOrganizer) {
    throw "'$subject' isn't the organizer's copy of the meeting. Use the event in the organizer's calendar."
}
if ($meeting.type -ne 'singleInstance') {
    throw "'$subject' is part of a recurring series, which this script doesn't support."
}
if ($meeting.isCancelled) {
    throw "'$subject' has been cancelled."
}
if ($changingTimes -and $meeting.isAllDay) {
    throw "'$subject' is an all-day meeting. This script can't change its time."
}

$fInvitedProp = $meeting.singleValueExtendedProperties | Where-Object { $_.id -match 'Id 0x0*8229$' }
$wasInvited   = $fInvitedProp.value -eq 'true'

# --- New times ---------------------------------------------------------------------------------
if ($changingTimes) {
    $currentStart = ConvertFrom-GraphUtc $meeting.start.dateTime
    $currentEnd   = ConvertFrom-GraphUtc $meeting.end.dateTime

    $start = if ($PSBoundParameters.ContainsKey('StartTime')) {
        ConvertTo-ZoneTime -Value $StartTime -Zone $zone
    }
    else {
        [TimeZoneInfo]::ConvertTimeFromUtc($currentStart, $zone)
    }

    $end = if ($PSBoundParameters.ContainsKey('EndTime')) {
        ConvertTo-ZoneTime -Value $EndTime -Zone $zone
    }
    elseif ($PSBoundParameters.ContainsKey('StartTime')) {
        $start + ($currentEnd - $currentStart)    # keep the meeting's length
    }
    else {
        [TimeZoneInfo]::ConvertTimeFromUtc($currentEnd, $zone)
    }

    if ($end -le $start) {
        throw 'EndTime must be later than StartTime.'
    }
}

# --- Attendees ---------------------------------------------------------------------------------
# The current attendees, then the new ones. An address that's already on the meeting keeps its
# current type.
$attendees = [System.Collections.Generic.List[hashtable]]::new()
$added     = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($attendee in $meeting.attendees) {
    $address = $attendee.emailAddress.address
    if (-not $address) {
        throw "'$subject' has an attendee without an email address ($($attendee.emailAddress.name)), so its attendees can't be removed and added back."
    }
    if ($added.Add($address)) {
        $emailAddress = @{ address = $address }
        if ($attendee.emailAddress.name) {
            $emailAddress.name = $attendee.emailAddress.name
        }
        $attendees.Add(@{ emailAddress = $emailAddress; type = $attendee.type })
    }
}
$currentCount = $attendees.Count

$groups = @(
    @{ Type = 'required'; Addresses = $AddRequiredAttendees }
    @{ Type = 'optional'; Addresses = $AddOptionalAttendees }
    @{ Type = 'resource'; Addresses = $AddResourceAttendees }
)
foreach ($group in $groups) {
    foreach ($address in $group.Addresses) {
        if ($added.Add($address)) {
            $attendees.Add(@{ emailAddress = @{ address = $address }; type = $group.Type })
        }
    }
}
$newCount = $attendees.Count - $currentCount

# --- Confirm -----------------------------------------------------------------------------------
$changes = @()
if ($changingTimes) {
    $changes += 'moved to {0:yyyy-MM-dd HH:mm} - {1:yyyy-MM-dd HH:mm} ({2})' -f $start, $end, $zone.Id
}
if ($newCount) {
    $changes += "added $newCount attendee(s)"
}
if (-not $changes) {
    Write-Warning "Nothing to change: the attendees are already on '$subject'."
    return
}
$summary = $changes -join ', '

if (-not $PSCmdlet.ShouldProcess("'$subject' in the calendar of $organizer", "Update without notifying attendees: $summary")) {
    return
}

# --- Update the meeting ------------------------------------------------------------------------
$timeChanges = @{}
if ($changingTimes) {
    $timeChanges.start = @{ dateTime = $start.ToString('s'); timeZone = $zone.Id }
    $timeChanges.end   = @{ dateTime = $end.ToString('s'); timeZone = $zone.Id }
}

$step             = ''
$attendeesRemoved = $false
try {
    if ($attendees.Count -eq 0) {
        # No attendees before or after the change, so there's no one to notify.
        $step   = 'changing the time'
        $result = Invoke-EventPatch -Uri $eventUri -Changes $timeChanges
    }
    else {
        # 1. Mark the meeting as never sent, so removing the attendees sends no cancellations.
        if ($wasInvited) {
            $step = 'setting PidLidFInvited to false'
            $null = Invoke-EventPatch -Uri $eventUri -Changes @{
                singleValueExtendedProperties = @(@{ id = $fInvitedId; value = 'false' })
            }
        }

        # 2. Remove the attendees.
        if ($currentCount) {
            $step             = 'removing the attendees'
            $null             = Invoke-EventPatch -Uri $eventUri -Changes @{ attendees = @() }
            $attendeesRemoved = $true
        }

        # 3. Add the attendees back with the new times, as a draft so no meeting requests are sent.
        $step                   = 'adding the attendees back as a draft'
        $draftChanges           = $timeChanges.Clone()
        $draftChanges.attendees = $attendees
        $draftChanges.isDraft   = $true
        $result                 = Invoke-EventPatch -Uri $eventUri -Changes $draftChanges
        $attendeesRemoved       = $false
        if (-not $result.isDraft) {
            Write-Warning "'$subject' wasn't saved as a draft, so meeting requests may have been sent."
        }

        # 4. Clear the draft flags so Outlook shows a normal meeting rather than a draft.
        $step   = 'clearing the draft flags'
        $result = Invoke-EventPatch -Uri $eventUri -Changes @{
            singleValueExtendedProperties = @(
                @{ id = "Boolean $psetidAppointment Name EventDraft"; value = 'false' }
                @{ id = "Boolean $psetidAppointment Name IntendedEventDraft"; value = 'false' }
            )
        }

        # 5. Mark the meeting as sent again.
        if ($RestoreInvitedFlag -and $wasInvited) {
            $step   = 'setting PidLidFInvited back to true'
            $result = Invoke-EventPatch -Uri $eventUri -Changes @{
                singleValueExtendedProperties = @(@{ id = $fInvitedId; value = 'true' })
            }
        }
    }
}
catch {
    $message = "Updating '$subject' failed while $step. $($_.Exception.Message)"
    if ($attendeesRemoved) {
        $removed = $attendees |
            Select-Object -First $currentCount |
            ForEach-Object { '{0} ({1})' -f $_.emailAddress.address, $_.type }
        $message += " The attendees were removed and not added back: $($removed -join ', ')."
    }
    throw $message
}

$invitedNow = if ($attendees.Count -eq 0) { $wasInvited } else { $wasInvited -and $RestoreInvitedFlag }

[pscustomobject]@{
    Id             = $result.id
    Subject        = $result.subject
    Changes        = $summary
    Attendees      = @($attendees | ForEach-Object { '{0} ({1})' -f $_.emailAddress.address, $_.type })
    PidLidFInvited = [bool]$invitedNow
    WebLink        = $result.webLink
}
