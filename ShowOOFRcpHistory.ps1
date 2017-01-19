function Connect-Exchange{ 
    param( 
    		[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
    )  
 	Begin
		 {
		Load-EWSManagedAPI
		
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
		  
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
		$service.AutodiscoverUrl($MailboxName,{$true})  
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

function ConvertId{    
 param (
         $HexId = "$( throw 'HexId is a mandatory Parameter' )"
    )
 process{
     $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId      
     $aiItem.Mailbox = $MailboxName      
     $aiItem.UniqueId = $HexId   
     $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId      
     $convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId) 
     return $convertedId.UniqueId
 }
}

function ConvertToString($ipInputString){  
    $Val1Text = ""  
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
            $clInt++  
    }  
    return $Val1Text  
}  
####################### 
<# 
.SYNOPSIS 
 Dumps Email Addresses for the Out of Office History property in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Dumps Email Addresses for the Out of Office History property in a Mailbox x using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
.EXAMPLE
	Example 1 To parsers the Email address from the FreeBusy Folder Mapi property 
    
	Get-OOFRcpHistory -MailboxName user@domain.com
#> 
########################
function Get-OOFRcpHistory{
	param (
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation		  )
	process{
        
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
		$EWSRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)
		$sf1 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"Freebusy Data")    
		$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)  
		$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;  
		$PR_DELEGATED_BY_RULE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3FE3, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		
		$fiResult = $EWSRootFolder.FindFolders($sf1,$fvFolderView)  
		if($fiResult.Folders.Count -eq 1){  
			$psPropset.Add($PR_DELEGATED_BY_RULE)
    		write-host "Found FreeBusy Folder"  
    		$Freebusyfld = $fiResult.Folders[0]  
    		$Freebusyfld.Load($psPropset) 
			$PropVal = $null
			if($Freebusyfld.TryGetProperty($PR_DELEGATED_BY_RULE,[ref]$PropVal)){
				$hexVal = [System.BitConverter]::ToString($PropVal).Replace("-","");
				$Sval = 0;  
				while(($Sval+12) -lt $hexVal.Length){  
					$Parse = $hexVal.SubString($Sval,12)  
					if($Parse -eq "00003349C843")
					{
						$Sval+=16
						$Parse = $hexVal.SubString(($Sval+2),2) + $hexVal.SubString($Sval,2)  
					    $ValLength = [System.Convert]::ToInt32($Parse,16)
						$Sval+=4
						$Parse = $hexVal.SubString($Sval,($ValLength*2))  
						$EmailAddress = ConvertToString($Parse)
						Write-Output $EmailAddress
						$Sval+=($ValLength*2)
					}
					else
					{
					#	Write-Host $Parse	
						$Sval+=2				
					}	
				    				
				}     
			}   		
		}  	
	    }
}

