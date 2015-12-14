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
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
		  
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
####################### 
<# 
.SYNOPSIS 
 Exports in the Calendars setting from a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Exports in the Calendars setting from a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Export the Calendar Settng from a  Mailbox
	 how-CalendarSettings -Mailboxname mailbox@domain.com 
	

#> 
########################
function Show-CalendarSettings
{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [string]$url
    )  
 	Begin
	{
		$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		# Bind to the Calendar Folder
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)   
		$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		$UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, "Calendar", $Calendar.Id, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)  
		#$UsrConfig.Dictionary
		$rptCollection = @()
		foreach($cnfObj in $UsrConfig.Dictionary.Key){
			$rptObj = "" | Select Property,Value,Valid
			$rptObj.Property = $cnfObj
			$rptObj.Valid = "Ok"
			if($UsrConfig.Dictionary[$cnfObj] -is [String[]])
			{
				$rptObj.Value = [String]::Join(",",$UsrConfig.Dictionary[$cnfObj])
			}
			else
			{
				$rptObj.Value = $UsrConfig.Dictionary[$cnfObj]
			}
			$rptCollection +=$rptObj
			if($cnfObj -eq "BookInPolicyLegDN")
			{
				foreach($LegDn in $UsrConfig.Dictionary["BookInPolicyLegDN"])
				{
				    $ncCol = $service.ResolveName($LegDn, [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly, $false);
	       			if($ncCol.Count -gt 0){
						#Write-output ("Found " + $ncCol[0].Mailbox.Address)
						$rptObj = "" | Select Property,Value,Valid
						$rptObj.Property = "BookInPolicyValue"
						$rptObj.Value = $ncCol[0].Mailbox.Address
						$rptObj.Valid = "Ok"
						$rptCollection +=$rptObj
	       			}
					else
					{	
						#Write-Output "Couldn't resolve " + $LegDn
						$rptObj = "" | Select Property,Value,Valid
						$rptObj.Property = "BookInPolicyValue"
						$rptObj.Value = $LegDn
						$rptObj.Valid = "Couldn't resolve"
						$rptCollection +=$rptObj
					}
					
				}	
			}
			if($cnfObj -eq "RequestInPolicyLegDN")
			{
				foreach($LegDn in $UsrConfig.Dictionary["RequestInPolicyLegDN"])
				{
				    $ncCol = $service.ResolveName($LegDn, [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly, $false);
	       			if($ncCol.Count -gt 0){
						#Write-output ("Found " + $ncCol[0].Mailbox.Address)
						$rptObj = "" | Select Property,Value,Valid
						$rptObj.Property = "RequestInPolicyValue"
						$rptObj.Value = $ncCol[0].Mailbox.Address
						$rptObj.Valid = "Ok"
						$rptCollection +=$rptObj
	       			}
					else
					{	
						#Write-Output "Couldn't resolve " + $LegDn
						$rptObj = "" | Select Property,Value,Valid
						$rptObj.Property = "RequestInPolicyValue"
						$rptObj.Value = $LegDn
						$rptObj.Valid = "Couldn't resolve"
						$rptCollection +=$rptObj
					}
					
				}			
			}
		}
		Write-Output $rptCollection
	 	
		$rptCollection | Export-Csv -Path ("$MailboxName-CalendarSetting.csv") -NoTypeInformation
	}
}