function Connect-Exchange{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [string]$url
    )  
 	Begin
		 {
		Load-EWSManagedAPI
		
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
		  
		## Create Exchange Service Object  
		$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
		  
		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		  
		#Credentials Option 1 using UPN for the windows Account  
		#$psCred = Get-Credential  
		$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString())  
		$service.Credentials = $creds      
		#Credentials Option 2  
		#service.UseDefaultCredentials = $true  
		 #$service.TraceEnabled = $true
		## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
		  
		Handle-SSL	
		  
		## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
		  
		#CAS URL Option 1 Autodiscover  
		if($url){
			$uri=[system.URI] $url
			$service.Url = $uri    
		}
		else{
			$service.AutodiscoverUrl($MailboxName,{$true})  
		}
		Write-host ("Using CAS Server : " + $Service.url)   
		   
		#CAS URL Option 2 Hardcoded  
		  
		#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
		#$service.Url = $uri    
		  
		## Optional section for Exchange Impersonation  
		  
		#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		if(!$service.URL){
			throw "Error connecting to EWS"
		}
		else
		{		
			return $service
		}
	}
}

function Load-EWSManagedAPI{
    param( 
    )  
 	Begin
	{
		## Load Managed API dll  
		###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
		$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
		if (Test-Path $EWSDLL)
		    {
		    Import-Module $EWSDLL
		    }
		else
		    {
		    "$(get-date -format yyyyMMddHHmmss):"
		    "This script requires the EWS Managed API 1.2 or later."
		    "Please download and install the current version of the EWS Managed API from"
		    "http://go.microsoft.com/fwlink/?LinkId=255472"
		    ""
		    "Exiting Script."
		    exit
		    } 
  	}
}

function Handle-SSL{
    param( 
    )  
 	Begin
	{
		## Code From http://poshcode.org/624
		## Create a compilation environment
		$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
		$Compiler=$Provider.CreateCompiler()
		$Params=New-Object System.CodeDom.Compiler.CompilerParameters
		$Params.GenerateExecutable=$False
		$Params.GenerateInMemory=$True
		$Params.IncludeDebugInformation=$False
		$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
		$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
		$TAAssembly=$TAResults.CompiledAssembly

		## We now create an instance of the TrustAll and attach it to the ServicePointManager
		$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
		[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

		## end code from http://poshcode.org/624

	}
}


function Like-EWSMessage  {
	    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [string]$url,
		[Parameter(Position=4, Mandatory=$true)] [String]$Subject,
        [Parameter(Position=5, Mandatory=$false)] [switch]$Unlike
    )  
 	Begin
	{
		if($url){
			$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url 
		}
		else{
			$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		}
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
		$AQSSearch = "Subject:" + $Subject 
		$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)    
        $fiItems = $service.FindItems($folderid,$AQSSearch,$ivItemView)
		if($fiItems.Items.Count -eq 1){
            Write-Host ("Processing Item " + $fiItems.Items[0].Subject)   
            if($Unlike.IsPresent){
                UnLike-Operation -Message $fiItems.Items[0] -Credentials $Credentials
            }
            else{
                Like-Operation -Message $fiItems.Items[0] -Credentials $Credentials
            }       
            
        }
	}
}

function Like-Operation
{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.EmailMessage]$Message,
        [Parameter(Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
    )  
 	Begin
	{
        $ItemId = $Message.Id.UniqueId
        $ItemChangeKey = $Message.Id.ChangeKey
        $request = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header>
    <t:RequestServerVersion Version="V2015_10_05"/>
  </soap:Header>
  <soap:Body xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
    <LikeItem>
      <ItemId Id="$ItemId" ChangeKey="$ItemChangeKey"/>
    </LikeItem>
  </soap:Body>
</soap:Envelope>
"@
        $LikeRequest = [System.Net.HttpWebRequest]::Create($Message.Service.url.ToString());
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($request);
        $LikeRequest.ContentLength = $bytes.Length;
        $LikeRequest.ContentType = "text/xml";
        $LikeRequest.UserAgent = "EWS Liker";            
        $LikeRequest.Headers.Add("Translate", "F");
        $LikeRequest.Method = "POST";
        $LikeRequest.Credentials =  New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString())  
        $RequestStream = $LikeRequest.GetRequestStream();
        $RequestStream.Write($bytes, 0, $bytes.Length);
        $RequestStream.Close();
        $LikeRequest.AllowAutoRedirect = $true;
        $Response = $LikeRequest.GetResponse().GetResponseStream()
        $sr = New-Object System.IO.StreamReader($Response)
        [XML]$xmlReposne = $sr.ReadToEnd()
        if($xmlReposne.Envelope.Body.LikeItemResponse.ResponseClass -eq "Success"){
            Write-Host("Item Liked")
        }
        else
        {
            Write-Host  $sr.ReadToEnd()
        } 
	}
	
}
function UnLike-Operation
{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.EmailMessage]$Message,
        [Parameter(Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
    )  
 	Begin
	{
        $ItemId = $Message.Id.UniqueId
        $ItemChangeKey = $Message.Id.ChangeKey
        $request = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header>
    <t:RequestServerVersion Version="V2015_10_05"/>
  </soap:Header>
  <soap:Body xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
    <LikeItem>
      <ItemId Id="$ItemId" ChangeKey="$ItemChangeKey"/>
      <IsUnlike>true</IsUnlike>
    </LikeItem>
  </soap:Body>
</soap:Envelope>
"@
        $LikeRequest = [System.Net.HttpWebRequest]::Create($Message.Service.url.ToString());
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($request);
        $LikeRequest.ContentLength = $bytes.Length;
        $LikeRequest.ContentType = "text/xml";
        $LikeRequest.UserAgent = "EWS Liker";            
        $LikeRequest.Headers.Add("Translate", "F");
        $LikeRequest.Method = "POST";
        $LikeRequest.Credentials =  New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString())  
        $RequestStream = $LikeRequest.GetRequestStream();
        $RequestStream.Write($bytes, 0, $bytes.Length);
        $RequestStream.Close();
        $LikeRequest.AllowAutoRedirect = $true;
        $Response = $LikeRequest.GetResponse().GetResponseStream()
        $sr = New-Object System.IO.StreamReader($Response)
        [XML]$xmlReposne = $sr.ReadToEnd()
        if($xmlReposne.Envelope.Body.LikeItemResponse.ResponseClass -eq "Success"){
            Write-Host("Item Unliked")
        }
        else
        {
            Write-Host  $sr.ReadToEnd()
        } 
	}
	
}