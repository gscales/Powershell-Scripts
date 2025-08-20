/// <summary>
/// Performs the Autodiscover request to get user settings.
/// </summary>
using Azure.Core;
using Azure.Identity;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Text;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;

var builder = WebApplication.CreateBuilder(args);
// This class represents a single alternate mailbox from the response.

// Add services to the container.
builder.Services.AddEndpointsApiExplorer();
// SwaggerGen is not available by default in .NET 9 preview, so comment out or remove the following line:
// builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
// Swagger is not available by default in .NET 9 preview, so comment out or remove the following lines:
/*
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}
*/

app.UseHttpsRedirection();

// The new REST endpoint
// This endpoint accepts a user's UPN (email) as a query parameter.
app.MapGet("/alternate-mailboxes", async ([FromQuery] string upn) =>
{
    // Retrieve credentials from a more secure location than hardcoded.
    // In a real application, you'd use a configuration manager.
    var clientId = "55105516-ccf7-45c2-b121-xxxx";
    var clientSecret = "xxx";
    var tenantId = "13af9f3c-b494-4795-bb19-xxx";

    if (string.IsNullOrEmpty(upn))
    {
        return Results.BadRequest("The 'upn' query parameter is required.");
    }

    try
    {
        // 1. Get an access token for Autodiscover
        var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        var easToken = await clientSecretCredential.GetTokenAsync(
            new TokenRequestContext(new[] { "https://outlook.office365.com/.default" })
        );

        // 2. Perform Autodiscover and parse the response
        var alternateMailboxes = await PerformAutodiscoverAsync(upn, easToken.Token);

        if (alternateMailboxes == null || !alternateMailboxes.Any())
        {
            return Results.Ok(alternateMailboxes);
        }

        // 3. Get a separate token for Microsoft Graph
        var graphServiceClient = new GraphServiceClient(clientSecretCredential);

        // 4. Iterate through alternate mailboxes and get the UserPurpose
        foreach (var mailbox in alternateMailboxes)
        {
            var userPurpose = await GetUserPurpose(graphServiceClient, mailbox.SmtpAddress ?? "");
            mailbox.UserPurpose = userPurpose; // Add UserPurpose to the object
        }

        // Return the final list as a JSON response
        return Results.Ok(alternateMailboxes);
    }
    catch (Exception ex)
    {
        // Log the exception and return a generic server error
        Console.WriteLine($"An error occurred: {ex.Message}");
        return Results.StatusCode(500);
    }
})
.WithName("GetAlternateMailboxes");
// .WithOpenApi(); // Remove or comment out if OpenAPI/Swagger is not available

app.Run();


/// <summary>
/// Parses the XML response from the Autodiscover GetUserSettings operation.
/// </summary>
static List<AlternateMailbox> ParseAutodiscoverResponse(string responseXml)
{
    var mailboxes = new List<AlternateMailbox>();
    try
    {
        var doc = XDocument.Parse(responseXml);
        XNamespace autodiscoverNs = "http://schemas.microsoft.com/exchange/2010/Autodiscover";
        var alternateMailboxElements = doc.Descendants(autodiscoverNs + "AlternateMailbox");

        foreach (var mailboxElement in alternateMailboxElements)
        {
            mailboxes.Add(new AlternateMailbox
            {
                Type = mailboxElement.Element(autodiscoverNs + "Type")?.Value,
                DisplayName = mailboxElement.Element(autodiscoverNs + "DisplayName")?.Value,
                LegacyDN = mailboxElement.Element(autodiscoverNs + "LegacyDN")?.Value,
                SmtpAddress = mailboxElement.Element(autodiscoverNs + "SmtpAddress")?.Value
            });
        }
    }
    catch (System.Xml.XmlException ex)
    {
        Console.WriteLine($"Error parsing XML: {ex.Message}");
    }
    return mailboxes;
}

/// <summary>
/// Performs the Autodiscover request to get user settings.
/// </summary>
static async Task<List<AlternateMailbox>?> PerformAutodiscoverAsync(string userEmail, string accessToken)
{
    string autodiscoverUrl = "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc";
    string soapXml = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:a=""http://schemas.microsoft.com/exchange/2010/Autodiscover"" xmlns:wsa=""http://www.w3.org/2005/08/addressing"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">
  <soap:Header>
    <a:RequestedServerVersion>Exchange2013</a:RequestedServerVersion>
    <wsa:Action>http://schemas.microsoft.com/exchange/2010/Autodiscover/Autodiscover/GetUserSettings</wsa:Action>
    <wsa:To>https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc</wsa:To>
  </soap:Header>
  <soap:Body>
    <a:GetUserSettingsRequestMessage xmlns:a=""http://schemas.microsoft.com/exchange/2010/Autodiscover"">
      <a:Request>
        <a:Users>
          <a:User>
            <a:Mailbox>{userEmail}</a:Mailbox>
          </a:User>
        </a:Users>
        <a:RequestedSettings>
          <a:Setting>AlternateMailboxes</a:Setting>
        </a:RequestedSettings>
      </a:Request>
    </a:GetUserSettingsRequestMessage>
  </soap:Body>
</soap:Envelope>";

    using (var client = new HttpClient())
    {
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        client.DefaultRequestHeaders.UserAgent.ParseAdd("MyApplication/1.0 (Windows NT 10.0)");
        var requestContent = new StringContent(soapXml, Encoding.UTF8, "text/xml");
        var response = await client.PostAsync(autodiscoverUrl, requestContent);

        if (response.IsSuccessStatusCode)
        {
            string responseXml = await response.Content.ReadAsStringAsync();
            return ParseAutodiscoverResponse(responseXml);
        }
        else
        {
            string errorContent = await response.Content.ReadAsStringAsync();
            Console.WriteLine($"Autodiscover failed with status code: {response.StatusCode}. Error: {errorContent}");
            return null;
        }
    }
}

/// <summary>
/// Retrieves the MailboxSettings.UserPurpose for a given user via Microsoft Graph.
/// </summary>
static async Task<string?> GetUserPurpose(GraphServiceClient graphServiceClient, string userId)
{
    try
    {
        var mailboxSettings = await graphServiceClient.Users[userId].MailboxSettings.GetAsync();
        return mailboxSettings?.UserPurpose?.ToString();
    }
    catch (ServiceException ex)
    {
        Console.WriteLine($"Error retrieving user purpose for '{userId}': {ex.Message}");
        return null;
    }
}

public class AlternateMailbox
{
    public string? Type { get; set; }
    public string? DisplayName { get; set; }
    public string? LegacyDN { get; set; }
    public string? SmtpAddress { get; set; }
    public string? UserPurpose { get; set; }
}

