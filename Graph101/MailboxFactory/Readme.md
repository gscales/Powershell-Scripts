## Mailbox Factory PowerShell Sample : Allow switching between the Microsoft Graph PowerShell SDK and EWS Managed API seamlessly when using Hybrid Modern Authentication and the Client Credentials OAuth Flow

With the recent depreciation of the Graph hybrid endpoint https://learn.microsoft.com/en-us/graph/hybrid-rest-support there is no longer a way to have one API call to access mailboxes in a Hybrid enviroment. Therefore you need to use EWS for your onPrem mailboxes while the Microsoft Graph should be used for anything that is hosted in Exchange Online. I recently gave a talk at the MEC Airlift event https://www.youtube.com/watch?v=a8ECVg_0hQI where i went through the challenges associated with building modern apps against Exchange no matter where its hosted. As part of that i created some sample c# code for discovery and abstraction which is shared https://github.com/gscales/MEC-Talk-2022. 

In PowerShell the way in which you may want to handle this can vary eg one approach could be to create two separate scripts or modules and then call the functions in each module based on your Autodiscover or other discovery method. Or with the introduction of Classes in PowerShell 5 what can be done is to use a Factory pattern https://en.wikipedia.org/wiki/Factory_(object-oriented_programming) and interfaces (abstract classes).  In this example I create a MailboxClientFactory that has an abstract base class MailboxClient and two classes that implement this, EWSClient and GraphClient that the factory will return relative to the mailbox that is being accessed. To determine which client to return the Factory uses multiple discovery methods such as the OpenId Discovery document and Autodiscoverv2. So once the factory has determined that a Mailbox is OnPrem or Online it will create or return an Instance of the EWSClient or GraphClient class. Everything in this example uses the client credentials flow and MSAL, the OnPrem code expects Hybrid Modern Authentication is enabled and the same Azure App registration is being used.

In this example the MailboxClient class has only one method GetInboxItemCount which returns the count of the number of items in the Inbox. In the class that implements MailboxClient this method is overridden with code relative to that API (or data provider) eg in the PowerShell Graph SDK its

         $Folder = Get-MgUserMailFolder -MailFolderId Inbox -UserId $MailboxName
         return $Folder.TotalItemCount

in the EWS Managed API its

       $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MailboxName)
       $InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($this.Service, $folderid) 
       return $InboxFolder.TotalCount;
       
## Using the factory

Pre-Reqs is for the Microsoft Graph you need to have the PowerShell Graph SDK installed https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0

For EWS the EWS Managed API and MSAL libraries are needed which are included with in the GitHub repo, a custom version of the EWS Managed API from https://github.com/gscales/EWS-BasicToOAuth-Info/blob/main/EWA%20Managed%20API%20MASL%20Token%20Refresh.md is used which support using the client credentials flow and implements token refresh/caching which is important if the script is going to be long running. 

Because of the dependencies and the way, the classes and modules work in PowerShell using Import-Module won't work for this module to use it you need to make using of using module eg to load the classes and dependencies 

    using module .\MailboxFactory.psd1

Unlike some factory implementations this doesn't use Static methods so you need to first create an instance of the factory where you pass the clientId of your appication registration and tenantId and the certificate you obtain either via

    $cert = Get-PfxCertificate -FilePath C:\temp\hbc.pfx 
    
    or
    
    $cert = Get-ChildItem -Path cert:\currentUser\My\$certThumbprint

    $mbFactory = [MailboxClientFactory]::new("d66f79ab-9457-46ba-b544-xxxxx","13af9f3c-b494-4795-bb19-xxxx",$cert)
    
 Then get an instance of the MailboxClient for the mailbox you want to access from the Factory
 
    $mbClient = $mbFactory.GetMailboxClient("gscales@48blah.domain.com")
    
Then you can use the MailboxClient methods eg

    $mbClient.GetInboxItemCount("gscales@48blah.domain.com")
  
Note this is version 1.0 so there are some improvments needed (eg the way in which the certificate is handled)

    
    
    
