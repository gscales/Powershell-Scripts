 class MailboxClient
 {
     MailboxClient ([string] $EndPoint, [string] $ClientId, [string] $Tenantid, [PsObject] $Certificate)
     {
         $type = $this.GetType()
 
         if ($type -eq [MailboxClient])
         {
             throw("Class $type must be inherited")
         }          
     }
 
     [INT64] GetInboxItemCount([string] $MailboxName)
     {
         throw("Must Override Method")
     }
 
 }
 
 class MSALAppToken : Microsoft.Exchange.WebServices.Data.Credentials.CustomTokenCredentials
 {
    [string] $EndPoint 
    [Microsoft.Identity.Client.IConfidentialClientApplication] $App
     MSALAppToken ([string] $EndPoint, [string] $ClientId, [string] $Tenantid, [PsObject] $Certificate){
         $this.App = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($ClientId).WithCertificate($Certificate).WithTenantId($Tenantid).Build();
         $this.EndPoint = $EndPoint
     }
     
     [string]GetCustomToken(){
        [string[]] $scopes = $this.EndPoint + "/.default"
        $authenticationResult = $this.App.AcquireTokenForClient($scopes).ExecuteAsync().Result;
        return "Bearer " + $authenticationResult.AccessToken;
     }
 }

 class EWSClient : MailboxClient
 {
    [PSObject] $Service
    EWSClient ([string] $EndPoint, [string] $ClientId, [string] $Tenantid, [PsObject] $Certificate) : base ($EndPoint, $ClientId, $Tenantid, $Certificate)
    {
      $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016_SP1      
      $this.Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
      $this.Service.credentials = [MSALAppToken]::new($EndPoint,$ClientId,$Tenantid,$Certificate)
    }
 
     # @Override
     [INT64] GetInboxItemCount([string] $MailboxName)
     {
       $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MailboxName)
       $InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($this.Service, $folderid) 
       return $InboxFolder.TotalCount;
     }
 }

 class GraphClient : MailboxClient
 {
     GraphClient ([string] $EndPoint, [string] $ClientId, [string] $Tenantid, [PsObject] $Certificate) : base ($EndPoint, $ClientId, $Tenantid, $Certificate)
     {
       connect-mggraph -ClientId $ClientId -Certificate $Certificate -Tenantid $Tenantid
     }
 
     [INT64] GetInboxItemCount([string] $MailboxName)
     {
         $Folder = Get-MgUserMailFolder -MailFolderId Inbox -UserId $MailboxName
         return $Folder.TotalItemCount
     }
 }
 

 class MailboxClientFactory
 {  
     [MailboxClient] $GraphClient
     [MailboxClient] $EWSClient
     [string] $ClientId
     [string] $Tenantid
     [PsObject] $Certificate

     MailboxClientFactory([string] $ClientId, [string] $Tenantid, [PsObject] $Certificate){
        
        $this.ClientId = $ClientId
        $this.Tenantid = $Tenantid
        $this.Certificate = $Certificate        
     }
     #Create an instance
     [MailboxClient] GetMailboxClient([string] $MailboxName)
     { 
        $graphEndPoint = $this.GraphOpenIdDiscovery($MailboxName.Split('@')[1])
        $outlookEndPoint =  $this.GraphToO365Endpoint($graphEndPoint)
        $adUrl = $this.AutoDiscoverV2($MailboxName,$outlookEndPoint)
        $host = ([uri]$adUrl).Host
        if($host  -eq "outlook.office365.com"){
            if($this.GraphClient -eq $null){
                $this.GraphClient = [GraphClient]::new("https://$graphEndPoint", $this.ClientId,$this.Tenantid,$this.Certificate)
            }
            return $this.GraphClient
        }else{
            if($this.EWSClient -eq $null){
                [Uri]::($adUrl).host
                $this.EWSClient = [EWSClient]::new("https://$host", $this.ClientId,$this.Tenantid,$this.Certificate)
                $uri=[system.URI] $adUrl
                $this.EWSClient.Service.Url = $uri  
            }
            $this.EWSClient.Service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
            $this.EWSClient.Service.HttpHeaders.Remove("X-AnchorMailbox");
            $this.EWSClient.Service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);
            return $this.EWSClient
        }
     }

     [string] AutoDiscoverV2([string] $emailAddress, [string] $serverEndPoint)
     {
        $autoDiscoverEndpoint = "https://$serverEndPoint/autodiscover/autodiscover.json?Email=" + [Uri]::EscapeDataString($emailAddress) + "&Protocol=EWS&RedirectCount=3";
        $adresponse = Invoke-RestMethod -Uri $autoDiscoverEndpoint
        if($adresponse.Url){
            return $adresponse.Url
        }else{
            return $null
        }
     }
     [string] GraphToO365Endpoint([string] $graphEndPoint)
     {
         switch ($graphEndPoint)
         {
             "microsoftgraph.chinacloudapi.cn" {return "partner.outlook.cn"}
             "graph.microsoft.de" {return "outlook.office.de"}
             "dod-graph.microsoft.com" {return "outlook-dod.office365.us"}
             "graph.microsoft.us" {return "outlook.office365.us"}            
         }
         return "outlook.office365.com"
     }

     [string] GraphOpenIdDiscovery([string] $domainName)
     {
         $odicEndpoint = "https://login.microsoftonline.com/$domainName/.well-known/openid-configuration";
         $odResponse = Invoke-RestMethod -uri $odicEndpoint
         if($odResponse.msgraph_host){
             return $odResponse.msgraph_host             
         }
         return $null
     }

 }
