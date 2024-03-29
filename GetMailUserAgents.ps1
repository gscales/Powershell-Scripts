﻿function Connect-Exchange{ 
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
		if (Test-Path ($PSScriptRoot + "/Microsoft.Exchange.WebServices.OauthMod.dll")) {
			Import-Module ($PSScriptRoot + "/Microsoft.Exchange.WebServices.OauthMod.dll")			
			write-verbose ("Using EWS dll from Local Directory")
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

function Get-MailUserAgents{
    param(
	    [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [string]$url
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
		# Bind to the SentItems Folder		
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$MailboxName)   
		$SentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		#Define ItemView to retrive just 1000 Items    
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$ClientInfo = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Guid]::Parse("41F28F13-83F4-4114-A584-EEDB5A6B0BFF"), "ClientInfo",[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
		$psPropset.Add($ClientInfo)
		$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
		$ivItemView.PropertySet = $psPropset
		$reportHash = @{}
		$fiItems = $null    
		do{    
		    $fiItems = $service.FindItems($SentItems.Id,$ivItemView)    
		    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
		    foreach($Item in $fiItems.Items){      
				$ClientInfoVal = $null
				if($Item.TryGetProperty($ClientInfo,[ref]$ClientInfoVal)){
					$unknownClient = $false
					if([string]::IsNullOrEmpty($ClientInfoVal)){
						$unknownClient = $true
					}
					else{
						if($ClientInfoVal.Contains("Client=")){
							$ValArray = $ClientInfoVal.Split(';')
							if($ValArray[0].Length -gt 7){
								$clientValParsed = $ValArray[0].Substring(7)
								if($reportHash.ContainsKey($clientValParsed))
								{
									$reportHash[$clientValParsed].ItemCount ++
									$reportHash[$clientValParsed].ItemSize += $Item.Size
								}
								else
								{
									$rptobj = "" | select Client,ItemCount,ItemSize
									$rptobj.Client = $clientValParsed
									$rptobj.ItemCount = 1
									$rptobj.ItemSize = $Item.Size
									$reportHash.Add($clientValParsed,$rptobj)
								}
							}
						}
						else{
							$unknownClient = $true
						}
					}
					if($unknownClient){
						if($reportHash.ContainsKey("Unknown"))
						{
							$reportHash[$clientValParsed].ItemCount ++
							$reportHash[$clientValParsed].ItemSize += $Item.Size
						}
						else
						{
							$rptobj = "" | select Client,ItemCount,ItemSize
							$rptobj.Client = "Unknown"
							$rptobj.ItemCount = 1
							$rptobj.ItemSize = $Item.Size
							$reportHash.Add("Unknown",$rptobj)
						}
					}
				}
		    }    
		    $ivItemView.Offset += $fiItems.Items.Count    
		}while($fiItems.MoreAvailable -eq $true) 
		$reportHash.GetEnumerator() | % { Write-Output $_.Value }
	}
	
}
