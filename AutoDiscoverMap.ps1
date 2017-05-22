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

function Get-AutoDiscoverMailboxMap{
	param (
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [psObject]$Mailboxes
				  )
	process{
            Load-EWSManagedAPI
            $MbList = @{}
            $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
            $autodiscover = New-Object Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($ExchangeVersion);
            $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString())  
		    $autodiscover.Credentials = $creds  
            $autodiscover.RedirectionUrlValidationCallback = {$true};
 	        foreach ($Mailbox in $Mailboxes) {
               write-host ("Processing Mailbox " + $Mailbox)
               if(!$MbList.ContainsKey($Mailbox)){
                   $mbComp = "" | Select MailboxName,AlternateMailboxCount,ReverseAlternateMailboxCount,AlternateMailboxList,ReverseAlternateMailboxList
                   $mbComp.MailboxName = $Mailbox
                   $mbComp.ReverseAlternateMailboxCount = 0
                   $mbComp.ReverseAlternateMailboxList = @()
                   $mbList.Add($Mailbox,$mbComp)
               }
               $Result =  $autodiscover.GetUserSettings($Mailbox,[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::AlternateMailboxes)
               $Alternates = $Result.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::AlternateMailboxes];
               $alternateList = @()
               foreach($Alternate in $Alternates.Entries){
                   if($Alternate.Type -eq "Delegate"){
                        $alternateList += $Alternate.OwnerSmtpAddress
                        if($mbList.ContainsKey($Alternate.OwnerSmtpAddress)){
                            $mbList[$Alternate.OwnerSmtpAddress].ReverseAlternateMailboxCount++
                            $mbList[$Alternate.OwnerSmtpAddress].ReverseAlternateMailboxList += $Mailbox
                        }
                        else {
                            $mbComp = "" | Select MailboxName,AlternateMailboxCount,ReverseAlternateMailboxCount,AlternateMailboxList,ReverseAlternateMailboxList
                            $mbComp.MailboxName = $Alternate.OwnerSmtpAddress
                            $mbComp.ReverseAlternateMailboxCount = 1
                            $mbComp.AlternateMailboxCount = 0
                            $mbComp.ReverseAlternateMailboxList = @()
                            $mbComp.ReverseAlternateMailboxList += $Mailbox
                            $mbList.Add($Alternate.OwnerSmtpAddress,$mbComp)    
                        }
                   }
               } 
               $mbList[$Mailbox].AlternateMailboxList = $alternateList
               $mbList[$Mailbox].AlternateMailboxCount = $alternateList.Count
           }
           $MbList.Values | Select MailboxName,AlternateMailboxCount,ReverseAlternateMailboxCount,@{Name=’AlternateMailboxList’;Expression={[string]::join(";", ($_.AlternateMailboxList))}},@{Name=’ReverseAlternateMailboxList’;Expression={[string]::join(";", ($_.ReverseAlternateMailboxList))}} | Export-Csv -NoTypeInformation -Path "c:\temp\AutoDiscoverMap.csv"
	    }
}

