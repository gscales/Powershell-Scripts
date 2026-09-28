// Create a meeting in Exchange Online with the Microsoft Graph .NET SDK without sending
// meeting requests to the attendees.
//
// The meeting is created as a draft (IsDraft = true), which stops Exchange sending meeting
// requests. The EventDraft and IntendedEventDraft properties are then set to false so Outlook
// shows it as a normal meeting instead of a draft.
//
// Attendees receive nothing and the meeting isn't added to their calendars. Rooms added as
// resource attendees aren't booked, because a room books itself when it processes a meeting
// request.
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

var meeting = await SilentMeeting.CreateAsync(
    graphClient,
    mailbox: "organizer@contoso.com",
    subject: "Quarterly planning",
    start: new DateTime(2026, 10, 6, 9, 0, 0),
    end: new DateTime(2026, 10, 6, 10, 0, 0),
    timeZone: "AUS Eastern Standard Time",
    requiredAttendees: ["alex@contoso.com", "sam@contoso.com"],
    optionalAttendees: ["jo@contoso.com"],
    resourceAttendees: ["boardroom@contoso.com"]);

Console.WriteLine($"Created meeting {meeting?.Id}");

static class SilentMeeting
{
    // EventDraft and IntendedEventDraft are string-named properties in PSETID_Appointment.
    const string PsetidAppointment = "{00062002-0000-0000-C000-000000000046}";

    /// <summary>Creates a meeting without sending meeting requests to the attendees.</summary>
    /// <param name="mailbox">The organizer's UPN or user ID.</param>
    /// <param name="start">The start, as a wall-clock time in <paramref name="timeZone"/>.</param>
    /// <param name="end">The end, as a wall-clock time in <paramref name="timeZone"/>.</param>
    /// <param name="timeZone">A Windows or IANA time zone name.</param>
    public static async Task<Event?> CreateAsync(
        GraphServiceClient graphClient,
        string mailbox,
        string subject,
        DateTime start,
        DateTime end,
        string timeZone,
        IEnumerable<string>? requiredAttendees = null,
        IEnumerable<string>? optionalAttendees = null,
        IEnumerable<string>? resourceAttendees = null)
    {
        var events = graphClient.Users[mailbox].Events;

        // 1. Create the meeting as a draft, so no meeting requests are sent.
        var draft = await events.PostAsync(new Event
        {
            Subject = subject,
            Start = new DateTimeTimeZone { DateTime = start.ToString("s"), TimeZone = timeZone },
            End = new DateTimeTimeZone { DateTime = end.ToString("s"), TimeZone = timeZone },
            Attendees =
            [
                .. ToAttendees(requiredAttendees, AttendeeType.Required),
                .. ToAttendees(optionalAttendees, AttendeeType.Optional),
                .. ToAttendees(resourceAttendees, AttendeeType.Resource),
            ],
            IsDraft = true,
            TransactionId = Guid.NewGuid().ToString(), // a retried POST won't create a duplicate
        });
        var draftId = draft?.Id ?? throw new InvalidOperationException("Graph did not return the new meeting.");

        // 2. Clear the draft flags so Outlook shows a normal meeting rather than a draft.
        return await events[draftId].PatchAsync(new Event
        {
            SingleValueExtendedProperties =
            [
                new() { Id = $"Boolean {PsetidAppointment} Name EventDraft", Value = "false" },
                new() { Id = $"Boolean {PsetidAppointment} Name IntendedEventDraft", Value = "false" },
            ],
        });
    }

    static IEnumerable<Attendee> ToAttendees(IEnumerable<string>? addresses, AttendeeType type) =>
        (addresses ?? []).Select(address => new Attendee
        {
            EmailAddress = new EmailAddress { Address = address },
            Type = type,
        });
}
