// Delete a meeting with the Microsoft Graph .NET SDK without sending cancellations to the
// attendees.
//
// Deleting a meeting normally sends a cancellation to every attendee. To avoid that, the sample
// removes the attendees and sets PidLidFInvited to false, so the meeting is treated as never
// sent, then permanently deletes it. The meeting goes to the Purges folder in the mailbox's
// Recoverable Items, where Outlook can't see it.
//
// Attendees keep the meeting in their calendars, and rooms keep their booking.
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

await SilentMeeting.DeleteAsync(graphClient, mailbox: "organizer@contoso.com", eventId: "AAMkAG...");

Console.WriteLine("Meeting deleted");

static class SilentMeeting
{
    // PidLidFInvited: whether meeting requests have been sent for the meeting.
    const string FInvited = "Boolean {00062002-0000-0000-C000-000000000046} Id 0x8229";

    /// <summary>Permanently deletes a meeting without sending cancellations.</summary>
    /// <param name="mailbox">The organizer's UPN or user ID.</param>
    public static async Task DeleteAsync(GraphServiceClient graphClient, string mailbox, string eventId)
    {
        var meeting = graphClient.Users[mailbox].Events[eventId];

        // 1. Remove the attendees and mark the meeting as never sent, so no cancellations are sent.
        await meeting.PatchAsync(new Event
        {
            Attendees = [],
            SingleValueExtendedProperties = [new() { Id = FInvited, Value = "false" }],
        });

        // 2. Permanently delete the meeting.
        await meeting.PermanentDelete.PostAsync();
    }
}
