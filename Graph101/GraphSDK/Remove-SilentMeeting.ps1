#Requires -Modules Microsoft.Graph.Authentication
<#
.SYNOPSIS
    Permanently deletes a meeting through Microsoft Graph without sending cancellations to the
    attendees.

.DESCRIPTION
    Deleting a meeting normally sends a cancellation to every attendee. To avoid that, the script:

      1. Removes the attendees and sets PidLidFInvited to false, in one request, so the meeting
         is treated as never sent.
      2. Permanently deletes the meeting (permanentDelete). It goes to the Purges folder in the
         mailbox's Recoverable Items, where Outlook can't see it, and is removed for good after
         the retention period unless the mailbox is on hold.

    Because nothing is sent, attendees keep the meeting in their calendars, and rooms keep their
    booking.

    Only single-instance meetings are supported. The script asks for confirmation before it
    deletes anything; use -Confirm:$false to skip that in unattended scripts.

.PARAMETER EventId
    The id of the meeting in the organizer's calendar, such as the Id returned by
    New-SilentMeeting.ps1. You are prompted for it if you leave it out.

.PARAMETER Mailbox
    The organizer's mailbox, as a UPN or user ID. Leave it out to use the signed-in user's
    calendar. Required with app-only authentication.

.EXAMPLE
    .\Remove-SilentMeeting.ps1 -EventId $id

    Deletes the meeting from the signed-in user's calendar after asking for confirmation.

.EXAMPLE
    .\Remove-SilentMeeting.ps1 -Mailbox organizer@contoso.com -EventId $id -Confirm:$false

    Deletes the meeting from another user's calendar without asking, for example in an
    unattended script using app-only authentication.

.NOTES
    Microsoft Graph permissions:
        Delegated, own calendar       Calendars.ReadWrite
        Delegated, another calendar   Calendars.ReadWrite.Shared, plus delegate access to it
        Application                   Calendars.ReadWrite

    If there is no Microsoft Graph connection, the script signs in interactively with the
    delegated permissions above. For unattended use, run Connect-MgGraph first.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory, HelpMessage = 'Id of the meeting in the organizer''s calendar')]
    [ValidateNotNullOrEmpty()]
    [string]$EventId,

    [string]$Mailbox
)

$ErrorActionPreference = 'Stop'

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
$mailboxSegment = if ($Mailbox) { "users/$Mailbox" } else { 'me' }
$eventUri       = "https://graph.microsoft.com/v1.0/$mailboxSegment/events/$EventId"
$organizer      = if ($Mailbox) { $Mailbox } else { 'the signed-in user' }

$meeting = Invoke-MgGraphRequest -Method GET -Uri ('{0}?$select=subject,isOrganizer,type,attendees' -f $eventUri)
$subject = $meeting.subject

if (-not $meeting.isOrganizer) {
    throw "'$subject' isn't the organizer's copy of the meeting. Use the event in the organizer's calendar."
}
if ($meeting.type -ne 'singleInstance') {
    throw "'$subject' is part of a recurring series, which this script doesn't support."
}

if (-not $PSCmdlet.ShouldProcess("'$subject' in the calendar of $organizer", 'Permanently delete without sending cancellations')) {
    return
}

# --- Delete the meeting ------------------------------------------------------------------------
$fInvitedId = 'Boolean {00062002-0000-0000-C000-000000000046} Id 0x8229'    # PidLidFInvited

# 1. Remove the attendees and mark the meeting as never sent, so no cancellations are sent.
$clearJson = @{
    attendees                     = @()
    singleValueExtendedProperties = @(@{ id = $fInvitedId; value = 'false' })
} | ConvertTo-Json -Depth 5
$null = Invoke-MgGraphRequest -Method PATCH -Uri $eventUri -Body $clearJson -ContentType 'application/json'

# 2. Permanently delete it.
try {
    $null = Invoke-MgGraphRequest -Method POST -Uri "$eventUri/permanentDelete"
}
catch {
    $message = "'$subject' wasn't deleted. $($_.Exception.Message)"
    $removed = @($meeting.attendees | ForEach-Object { '{0} ({1})' -f $_.emailAddress.address, $_.type })
    if ($removed) {
        $message += " Its attendees were already removed and PidLidFInvited set to false. The attendees were: $($removed -join ', ')."
    }
    throw $message
}