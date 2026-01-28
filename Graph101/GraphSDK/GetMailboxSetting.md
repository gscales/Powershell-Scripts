**How to Retrieve Mailbox Settings via Microsoft Graph**

To utilize the MailboxItems, MailboxFolder, or Import/Export endpoints, you must first discover the mailbox settings.

**Using the Microsoft.Graph PowerShell Module**

Because these endpoints are currently in the Beta stage, dedicated cmdlets for discovery are not yet available. However, you can use Invoke-MgGraphRequest to query the API directly.

**Method 1: Detailed Request**

Use this method to store the full response in a variable. This is useful if you need to inspect multiple properties within the Exchange settings.

PowerShell

    # Define the target user
    $Upn = "yourmailbox@domain.com"
    
    # Construct the Beta Request URL
    $RequestURL = "https://graph.microsoft.com/beta/users/$Upn/settings/exchange"
    
    # Execute the request
    $Results = Invoke-MgGraphRequest -Method Get -Uri $RequestURL

**Method 2: One-Liner (Primary Mailbox ID)**

Use this streamlined version if you only need the primaryMailboxId returned immediately.

PowerShell

    (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/users/yourmailbox@domain.com/settings/exchange").primaryMailboxId

[!IMPORTANT] Prerequisites:

Ensure you have authenticated via Connect-MgGraph.

Required permissions typically include MailboxSettings.Read or User.Read.All.

As these are Beta endpoints, the API schema may change.
