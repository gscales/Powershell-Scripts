// Change a meeting's time, or add attendees to it, with the Microsoft Graph .NET SDK without
// sending meeting updates or invitations.
//
// Changing a meeting's start or end time normally sends an update to every attendee, and adding
// attendees sends invitations to the new ones. To avoid that, the sample:
//   1. Removes the attendees and sets PidLidFInvited to false, so the meeting is treated as
//      never sent.
//   2. Adds the attendees back, with any new ones, as a draft (IsDraft = true), so no meeting
//      requests are sent.
//   3. Sets EventDraft to false so Outlook shows a normal meeting, and applies the new times.
//   4. Optionally sets PidLidFInvited back to true, so later changes send updates as usual.
//
// Attendees' copies of the meeting keep the old time, and new attendees don't get the meeting
// in their calendars.
//
// A complete Program.cs for a .NET 8 (or later) console app:
//   dotnet add package Microsoft.Graph
//   dotnet add package Azure.Identity
// Permission: Calendars.ReadWrite (application)

using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

var credential = new ClientSecretCredential("tenant-id", "client-id", "client-secret");
var graphClient = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);

var meeting = await SilentMeeting.UpdateAsync(
    graphClient,
    mailbox: "organizer@contoso.com",
    eventId: "AAMkAG...",
    start: new DateTime(2026, 10, 7, 10, 0, 0),
    end: new DateTime(2026, 10, 7, 11, 0, 0),
    timeZone: "AUS Eastern Standard Time",
    addRequiredAttendees: ["kim@contoso.com"],
    addResourceAttendees: ["projector@contoso.com"]);

Console.WriteLine($"Updated meeting {meeting?.Id}");

static class SilentMeeting
{
    const string PsetidAppointment = "{00062002-0000-0000-C000-000000000046}";
    const string FInvited = $"Boolean {PsetidAppointment} Id 0x8229"; // PidLidFInvited
    const string EventDraft = $"Boolean {PsetidAppointment} Name EventDraft";

    /// <summary>Changes a meeting's time and/or adds attendees without notifying anyone.</summary>
    /// <param name="mailbox">The organizer's UPN or user ID.</param>
    /// <param name="start">The new start, as a wall-clock time in <paramref name="timeZone"/>.</param>
    /// <param name="end">The new end, as a wall-clock time in <paramref name="timeZone"/>.</param>
    /// <param name="timeZone">A Windows or IANA time zone name.</param>
    /// <param name="restoreInvitedFlag">Mark the meeting as sent again afterwards, if it was sent before.</param>
    public static async Task<Event?> UpdateAsync(
        GraphServiceClient graphClient,
        string mailbox,
        string eventId,
        DateTime? start = null,
        DateTime? end = null,
        string timeZone = "UTC",
        IEnumerable<string>? addRequiredAttendees = null,
        IEnumerable<string>? addOptionalAttendees = null,
        IEnumerable<string>? addResourceAttendees = null,
        bool restoreInvitedFlag = false)
    {
        var meeting = graphClient.Users[mailbox].Events[eventId];

        var current = await meeting.GetAsync(config =>
            config.QueryParameters.Expand = [$"singleValueExtendedProperties($filter=id eq '{FInvited}')"])
            ?? throw new InvalidOperationException($"Meeting {eventId} wasn't found.");
        var wasInvited = string.Equals(
            current.SingleValueExtendedProperties?.FirstOrDefault()?.Value, "true", StringComparison.OrdinalIgnoreCase);

        List<Attendee> attendees =
        [
            .. (current.Attendees ?? []).Select(Copy),
            .. ToAttendees(addRequiredAttendees, AttendeeType.Required),
            .. ToAttendees(addOptionalAttendees, AttendeeType.Optional),
            .. ToAttendees(addResourceAttendees, AttendeeType.Resource),
        ];

        // 1. Remove the attendees and mark the meeting as never sent, so no cancellations are sent.
        await meeting.PatchAsync(new Event
        {
            Attendees = [],
            SingleValueExtendedProperties = [new() { Id = FInvited, Value = "false" }],
        });

        // 2. Add the attendees back, with the new ones, as a draft so no meeting requests are sent.
        await meeting.PatchAsync(new Event
        {
            IsDraft = true,
            Attendees = attendees,
            SingleValueExtendedProperties = [new() { Id = FInvited, Value = "false" }],
        });

        // 3. Clear the draft flag so Outlook shows a normal meeting, and apply the new times.
        //    Only set the times you're changing: the SDK sends every property you set, even null.
        var changes = new Event
        {
            SingleValueExtendedProperties = [new() { Id = EventDraft, Value = "false" }],
        };
        if (start is not null)
        {
            changes.Start = new DateTimeTimeZone { DateTime = start.Value.ToString("s"), TimeZone = timeZone };
        }
        if (end is not null)
        {
            changes.End = new DateTimeTimeZone { DateTime = end.Value.ToString("s"), TimeZone = timeZone };
        }
        var updated = await meeting.PatchAsync(changes);

        // 4. Optionally mark the meeting as sent again.
        if (restoreInvitedFlag && wasInvited)
        {
            updated = await meeting.PatchAsync(new Event
            {
                SingleValueExtendedProperties = [new() { Id = FInvited, Value = "true" }],
            });
        }

        return updated;
    }

    // Copies an attendee from the GET into a new object. The SDK only sends properties that have
    // been set on an object, so build new ones rather than reusing the objects the GET returned.
    static Attendee Copy(Attendee attendee)
    {
        var copy = new Attendee
        {
            Type = attendee.Type,
            EmailAddress = new EmailAddress
            {
                Name = attendee.EmailAddress?.Name,
                Address = attendee.EmailAddress?.Address,
            },
        };
        if (attendee.Status is not null)
        {
            copy.Status = new ResponseStatus { Response = attendee.Status.Response, Time = attendee.Status.Time };
        }
        return copy;
    }

    static IEnumerable<Attendee> ToAttendees(IEnumerable<string>? addresses, AttendeeType type) =>
        (addresses ?? []).Select(address => new Attendee
        {
            EmailAddress = new EmailAddress { Address = address },
            Type = type,
        });
}
