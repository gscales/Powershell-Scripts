####################### 
<# 
.SYNOPSIS 
 Enumerates Items in a Public folder in a Exchange Public folder using the  Exchange Web Services API 
 
.DESCRIPTION 
  Enumerates Items in a  Public folder in a Exchange Public folder  using the  Exchange Web Services API 
  
 Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
.EXAMPLE
 PS C:\>Get-FolderItems -MailboxName user.name@domain.com  -FolderPath '\Reporttest'

#> 
function Get-PublicFolderItemsAge
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$true)] [string]$PublicFolderPath,
		[Parameter(Position=3, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=4, Mandatory=$false)] [string]$url
    )  
 	Begin
		 {
		if($url)
		{
			$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url 
		}
		else{
			$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		}
		if($useImpersonation.IsPresent)
		{
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		}
        Get-PublicFolderRoutingHeader -service $service -Credentials $Credentials -MailboxName $MailboxName -Header "X-AnchorMailbox"
		$fldId = PublicFolderIdFromPath -FolderPath $PublicFolderPath  -SmtpAddress $MailboxName -service $service		
        $SubFolderId =  new-object Microsoft.Exchange.WebServices.Data.FolderId($fldId)
        $minYear = (Get-Date).Year
        $maxYear = (Get-Date).Year
		$rptCollection = @{}
        
		#Define ItemView to retrive just 1000 Items    
		$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000) 
		$ItemPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Size)
        $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
        $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
		$ivItemView.PropertySet = $ItemPropset
		$rptCollection = @{}
		$fiItems = $null    
		do{    
		    $fiItems = $service.FindItems($SubFolderId,$ivItemView)   
			Write-Host ("Processed " + $fiItems.Items.Count)
		    #[Void]$service.LoadPropertiesForItems($fiItems,$ItemPropset)  
		    foreach($Item in $fiItems.Items){      
				#Process Item
				$dateVal = $null
				if($Item.TryGetProperty([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[ref]$dateVal )-eq $false){
					$dateVal = $Item.DateTimeCreated
				}
				if($rptCollection.ContainsKey($dateVal.Year)){
					$rptCollection[$dateVal.Year].TotalNumber += 1
					$rptCollection[$dateVal.Year].TotalSize += [Int64]$Item.Size
				}
				else{
					$rptObj = "" | Select Year,TotalNumber,TotalSize
					$rptObj.TotalNumber = 1
                    $rptObj.Year = $dateVal.Year
					$rptObj.TotalSize = [Int64]$Item.Size
					$rptCollection.add($dateVal.Year,$rptObj)
					if($dateVal.Year -lt $minYear){$minYear = $dateVal.Year}
				}
		    }    
		    $ivItemView.Offset += $fiItems.Items.Count    
		}while($fiItems.MoreAvailable -eq $true) 
		Write-Output $rptCollection.Values | Sort-Object -Property NumberOfItems -Descending
		}
}

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
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
		  
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
function PublicFolderIdFromPath{
	param (
            [Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
            [Parameter(Position=1, Mandatory=$true)] [String]$FolderPath,
            [Parameter(Position=2, Mandatory=$true)] [String]$SmtpAddress
		  )
	process{
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root  
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)   
        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        $PR_REPLICA_LIST = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6698,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary); 
        $psPropset.Add($PR_REPLICA_LIST)
		$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)  
        $PR_REPLICA_LIST_Value = $null
        if($tfTargetFolder.TryGetProperty($PR_REPLICA_LIST,[ref]$PR_REPLICA_LIST_Value)){
                 $GuidAsString = [System.Text.Encoding]::ASCII.GetString($PR_REPLICA_LIST_Value, 0, 36);
                 $HeaderAddress = new-object System.Net.Mail.MailAddress($service.HttpHeaders["X-AnchorMailbox"])
                 $pfHeader = $GuidAsString + "@" + $HeaderAddress.Host
                 write-host ("Root Public Folder Routing Information Header : " + $pfHeader )  
                 $service.HttpHeaders.Add("X-PublicFolderMailbox", $pfHeader)    
        }
		#Split the Search path into an array  
		$fldArray = $FolderPath.Split("\") 
		 #Loop through the Split Array and do a Search for each level of folder 
		for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
	        #Perform search based on the displayname of each folder level 
	        $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
            $fvFolderView.PropertySet = $psPropset
	        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$fldArray[$lint]) 
	        $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView) 
	        if ($findFolderResults.TotalCount -gt 0){ 
	            foreach($folder in $findFolderResults.Folders){ 
	                $tfTargetFolder = $folder                
	            } 
	        } 
	        else{ 
	            "Error Folder Not Found"  
	            $tfTargetFolder = $null  
	            break  
	        }     
	    }  
		if($tfTargetFolder -ne $null){
            $PR_REPLICA_LIST_Value = $null
            if($tfTargetFolder.TryGetProperty($PR_REPLICA_LIST,[ref]$PR_REPLICA_LIST_Value)){
                    $GuidAsString = [System.Text.Encoding]::ASCII.GetString($PR_REPLICA_LIST_Value, 0, 36);
                    $HeaderAddress = new-object System.Net.Mail.MailAddress($service.HttpHeaders["X-AnchorMailbox"])
                    $pfHeader = $GuidAsString + "@" + $HeaderAddress.Host
                    write-host ("Target Public Folder Routing Information Header : " + $pfHeader )  
                    Get-PublicFolderContentRoutingHeader -service $service -Credentials $Credentials -MailboxName $SmtpAddress -pfAddress $pfHeader
            }            
			return $tfTargetFolder.Id.UniqueId.ToString()
		}
		else{
			throw "Folder not found"
		}
	}
}

function Get-PublicFolderRoutingHeader
{
    param (
	        [Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
            [Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
		    [Parameter(Position=2, Mandatory=$true)] [string]$MailboxName,
            [Parameter(Position=3, Mandatory=$true)] [string]$Header
          )
	process
    {
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
        $AutoDiscoverService =  New-Object  Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($ExchangeVersion);
        $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString()) 
        $AutoDiscoverService.Credentials = $creds
        $AutoDiscoverService.EnableScpLookup = $false;
        $AutoDiscoverService.RedirectionUrlValidationCallback = {$true};
        $AutoDiscoverService.PreAuthenticate = $true;
        $AutoDiscoverService.KeepAlive = $false;      
        if($Header -eq "X-AnchorMailbox")
        {
            $gsp = $AutoDiscoverService.GetUserSettings($MailboxName,[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation);
            $PublicFolderInformation = $null
            if ($gsp.Settings.TryGetValue([Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation, [ref] $PublicFolderInformation))
            {
                write-host ("Public Folder Routing Information Header : " + $PublicFolderInformation)  
                $service.HttpHeaders.Add($Header, $PublicFolderInformation)           
                        
            } 
            
        }

       
    }
    
}
function Get-PublicFolderContentRoutingHeader
{
    param (
	        [Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
            [Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
		    [Parameter(Position=2, Mandatory=$true)] [string]$MailboxName,
            [Parameter(Position=3, Mandatory=$true)] [string]$pfAddress
     )
	process
    {
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
        $AutoDiscoverService =  New-Object  Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($ExchangeVersion);
        $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString()) 
        $AutoDiscoverService.Credentials = $creds
        $AutoDiscoverService.EnableScpLookup = $false;
        $AutoDiscoverService.RedirectionUrlValidationCallback = {$true};
        $AutoDiscoverService.PreAuthenticate = $true;
        $AutoDiscoverService.KeepAlive = $false;      
        $gsp = $AutoDiscoverService.GetUserSettings($MailboxName,[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::AutoDiscoverSMTPAddress);
        #Write-Host $AutoDiscoverService.url
        $auDisXML = "<Autodiscover xmlns=`"http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006`"><Request>`r`n" +
        "<EMailAddress>" + $pfAddress + "</EMailAddress>`r`n" +
        "<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>`r`n" +
        "</Request>`r`n" +
        "</Autodiscover>`r`n";
        $AutoDiscoverRequest = [System.Net.HttpWebRequest]::Create($AutoDiscoverService.url.ToString().replace(".svc",".xml"));
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($auDisXML);
        $AutoDiscoverRequest.ContentLength = $bytes.Length;
        $AutoDiscoverRequest.ContentType = "text/xml";
        $AutoDiscoverRequest.UserAgent = "Microsoft Office/16.0 (Windows NT 6.3; Microsoft Outlook 16.0.6001; Pro)";            
        $AutoDiscoverRequest.Headers.Add("Translate", "F");
        $AutoDiscoverRequest.Method = "POST";
        $AutoDiscoverRequest.Credentials = $creds;
        $RequestStream = $AutoDiscoverRequest.GetRequestStream();
        $RequestStream.Write($bytes, 0, $bytes.Length);
        $RequestStream.Close();
        $AutoDiscoverRequest.AllowAutoRedirect = $truee;
        $Response = $AutoDiscoverRequest.GetResponse().GetResponseStream()
        $sr = New-Object System.IO.StreamReader($Response)
        [XML]$xmlReposne = $sr.ReadToEnd()
        if($xmlReposne.Autodiscover.Response.User.AutoDiscoverSMTPAddress -ne $null)
        {
            write-host ("Public Folder Content Routing Information Header : " + $xmlReposne.Autodiscover.Response.User.AutoDiscoverSMTPAddress)  
            $service.HttpHeaders["X-AnchorMailbox"] = $xmlReposne.Autodiscover.Response.User.AutoDiscoverSMTPAddress    
            $service.HttpHeaders["X-PublicFolderMailbox"] = $xmlReposne.Autodiscover.Response.User.AutoDiscoverSMTPAddress              
        }

    }
    
}