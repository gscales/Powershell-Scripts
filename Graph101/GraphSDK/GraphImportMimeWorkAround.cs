using System.Text;
using System.Text.Json;
using System.Net;
using System.Text.Json.Nodes;
using Microsoft.Graph.Beta;
using Microsoft.Graph.Beta.Models;
using Microsoft.Graph.Beta.Admin.Exchange.Mailboxes.Item.ExportItems;
using Azure.Identity;
using Microsoft.Kiota.Abstractions;


public class Program
{
    // Entry point
    public static async Task Main(string[] args)
    {
        // App registration and user configuration
        var clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx";
        var scopes = new[] { "Mail.ReadWrite", "MailboxItem.Read", "MailboxItem.ImportExport" };
        string emlPath = @"C:\Users\gscal\Downloads\message5.eml";
        string userId = "user@domain.com";

        // Set up Azure.Identity interactive authentication
        var interactiveCredential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
        {
            ClientId = clientId,
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            RedirectUri = new Uri("http://localhost")
        });

        // Set up GraphServiceClient
        var graphClient = new GraphServiceClient(interactiveCredential);

        try
        {
            // Read EML file and convert to base64
            byte[] emlBytes = await File.ReadAllBytesAsync(emlPath);
            string base64Eml = Convert.ToBase64String(emlBytes);
            var contentBytes = Encoding.UTF8.GetBytes(base64Eml);
            using var base64Content = new MemoryStream(contentBytes);

            // Create a new message via Graph API
            var requestInfo = new RequestInformation
            {
                HttpMethod = Method.POST,
                URI = new Uri($"https://graph.microsoft.com/v1.0/users/{userId}/messages")
            };
            requestInfo.Headers.Clear();
            requestInfo.SetStreamContent(base64Content, "text/plain");
            var createdMessage = await graphClient.RequestAdapter.SendAsync<Message>(
                requestInfo,
                Message.CreateFromDiscriminatorValue,
                errorMapping: null
            );

            if (createdMessage == null)
            {
                Console.WriteLine("Message creation failed.");
                return;
            }
            // Export message item
            var exportedItem = await GraphImportExport.ExportItemAsync(userId, graphClient, createdMessage.Id);
            if (exportedItem?.Data == null)
            {
                Console.WriteLine("Export failed.");
                return;
            }

            // Mark exported message as sent
            GraphImportExport.MarkMessageAsSent(exportedItem.Data);

            // Delete the original message
            await graphClient.Users[userId].Messages[createdMessage.Id].PermanentDelete.PostAsync();

            // Import the modified message
            var finalMessageId = await GraphImportExport.ImportMessageAsync(userId, graphClient, exportedItem.Data);

            Console.WriteLine($"Imported message ID: {finalMessageId}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Graph API Error: {ex.Message}");
        }
    }

    public static class GraphImportExport
    {
        public static async Task<string> ImportMessageAsync(string userId, GraphServiceClient graphServiceClient, byte[] data)
        {
            var mailboxSettings = await graphServiceClient.Users[userId].Settings.Exchange.GetAsync();
            var importSession = await graphServiceClient.Admin.Exchange.Mailboxes[mailboxSettings.PrimaryMailboxId]
                .CreateImportSession.PostAsync();
            var folder = await graphServiceClient.Users[userId].MailFolders["inbox"].GetAsync();
            var folderId = folder.Id;

            var requestDict = new Dictionary<string, object>
            {
                ["FolderId"] = folderId,
                ["Mode"] = "create",
                ["Data"] = Convert.ToBase64String(data)
            };

            string jsonBody = JsonSerializer.Serialize(requestDict);

            var handler = new HttpClientHandler { CookieContainer = new CookieContainer() };
            using var client = new HttpClient(handler);
            using var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
            var response = await client.PostAsync(importSession.ImportUrl, content);

            if (!response.IsSuccessStatusCode)
            {
                Console.WriteLine($"Import failed: {response.StatusCode} {response.ReasonPhrase}");
            }
            var responseJson = await response.Content.ReadAsStringAsync();
            try
            {
                var jsonNode = JsonNode.Parse(responseJson);
                var id = jsonNode?["itemId"]?.ToString();
                Console.WriteLine($"Parsed id: {id}");
                return id;
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"Failed to parse JSON: {ex.Message}");
                Console.WriteLine(responseJson);
                return null;
            }
        }

        public static async Task<ExportItemResponse> ExportItemAsync(string userId, GraphServiceClient graphServiceClient, string itemId)
        {
            var mailboxSettings = await graphServiceClient.Users[userId].Settings.Exchange.GetAsync();
            var requestBody = new ExportItemsPostRequestBody
            {
                ItemIds = new List<string> { itemId }
            };
            var result = await graphServiceClient.Admin.Exchange.Mailboxes[mailboxSettings.PrimaryMailboxId]
                .ExportItems.PostAsExportItemsPostResponseAsync(requestBody);

            return result.Value?.Count > 0 ? result.Value[0] : null;
        }

        // Finds the index of the PidTagMessageFlags property tag (0x0E070003, little-endian) in a byte array.
        public static int FindPidTagMessageFlags(byte[] data)
        {
            byte[] pattern = { 0x03, 0x00, 0x07, 0x0E };
            for (int i = 0; i <= data.Length - pattern.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < pattern.Length; j++)
                {
                    if (data[i + j] != pattern[j])
                    {
                        match = false;
                        break;
                    }
                }
                if (match)
                    return i;
            }
            return -1;
        }

        public static bool MarkMessageAsSent(byte[] data)
        {
            const int UnsentFlag = 0x0008;
            const int SubmitFlag = 0x0004;
            int tagIndex = FindPidTagMessageFlags(data);
            if (tagIndex < 0 || tagIndex + 8 > data.Length)
                return false;

            int valueOffset = tagIndex + 4;
            int oldFlags = BitConverter.ToInt32(data, valueOffset);

            // Remove UNSENT, add SUBMIT (preserve other flags)
            int newFlags = (oldFlags & ~UnsentFlag) | SubmitFlag;

            // Write new value back (little-endian)
            byte[] newFlagBytes = BitConverter.GetBytes(newFlags);
            Buffer.BlockCopy(newFlagBytes, 0, data, valueOffset, 4);
            return true;
        }

        public static void PrintPidTagMessageFlagsValue(byte[] data)
        {
            int index = FindPidTagMessageFlags(data);
            if (index < 0)
            {
                Console.WriteLine("PidTagMessageFlags property not found.");
                return;
            }
            int valueOffset = index + 4;
            if (valueOffset + 4 > data.Length)
            {
                Console.WriteLine("Not enough data for the value.");
                return;
            }
            int value = BitConverter.ToInt32(data, valueOffset);
            Console.WriteLine($"PidTagMessageFlags value: 0x{value:X8} ({value})");
        }
    }
}
