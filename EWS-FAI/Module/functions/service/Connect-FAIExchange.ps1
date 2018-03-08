function Connect-FAIExchange
{
<#
	.SYNOPSIS
		A brief description of the Connect-FAIExchange function.
	
	.DESCRIPTION
		A detailed description of the Connect-FAIExchange function.
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.EXAMPLE
		PS C:\> Connect-FAIExchange -MailboxName 'value1' -Credentials $Credentials
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[System.Management.Automation.PSCredential]
		$Credentials
	)
	Begin
	{
		## Load Managed API dll  
		###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
		$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services' | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
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
			$exception = New-Object System.Exception ("Managed Api missing")
			throw $exception
		}
		
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
		
		## Create Exchange Service Object  
		$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
		
		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		
		#Credentials Option 1 using UPN for the windows Account  
		#$psCred = Get-Credential  
		$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString())
		$service.Credentials = $creds
		#Credentials Option 2  
		#service.UseDefaultCredentials = $true  
		#$service.TraceEnabled = $true
		## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
		
		## Code From http://poshcode.org/624
		## Create a compilation environment
		$Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
		$Compiler = $Provider.CreateCompiler()
		$Params = New-Object System.CodeDom.Compiler.CompilerParameters
		$Params.GenerateExecutable = $False
		$Params.GenerateInMemory = $True
		$Params.IncludeDebugInformation = $False
		$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null
		
		$TASource = @'
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
		$TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
		$TAAssembly = $TAResults.CompiledAssembly
		
		## We now create an instance of the TrustAll and attach it to the ServicePointManager
		$TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
		[System.Net.ServicePointManager]::CertificatePolicy = $TrustAll
		
		## end code from http://poshcode.org/624
		
		## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
		
		#CAS URL Option 1 Autodiscover  
		$service.AutodiscoverUrl($MailboxName, { $true })
		Write-host ("Using CAS Server : " + $Service.url)
		
		#CAS URL Option 2 Hardcoded  
		
		#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
		#$service.Url = $uri    
		
		## Optional section for Exchange Impersonation  
		
		#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		if (!$service.URL)
		{
			throw "Error connecting to EWS"
		}
		else
		{
			return $service
		}
	}
}
