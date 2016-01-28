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

function ConvertToString($ipInputString){  
    $Val1Text = ""  
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
            $clInt++  
    }  
    return $Val1Text  
} 

function Enum-SearchFolders{ 
    param( 
		[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
		[Parameter(Position=2, Mandatory=$false)] [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId
    )  
 	Begin
	{
		$rptCollection = @()
		#Define Extended properties  
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);				
	    $PidTagMessageSizeExtended = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0xe08,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
		$psPropset.Add($PR_Folder_Path)
		$psPropset.Add($PidTagMessageSizeExtended)
		$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
		#Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
		$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
		#Deep Transval will ensure all folders in the search path are returned  
		$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
     	#Add Properties to the  Property Set  
		$fvFolderView.PropertySet = $psPropset;  
		#The Search filter will exclude any Search Folders  
		$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"2")  
		$fiResult = $null  
		#The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
		do {  
		    $fiResult = $Service.FindFolders($folderId,$sfSearchFilter,$fvFolderView)  
		    foreach($ffFolder in $fiResult.Folders){  
		        $foldpathval = $null  
		        #Try to get the FolderPath Value and then covert it to a usable String   
		        if ($ffFolder.TryGetProperty($PR_Folder_Path,[ref] $foldpathval))  
		        {  
		            $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
		            $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
		            $hexString = $hexArr -join ''  
		            $hexString = $hexString.Replace("FEFF", "5C00")  
		            $fpath = ConvertToString($hexString)  
		        }  
				$FolderSizeVal = $null
				[Void]$ffFolder.TryGetProperty($PidTagMessageSizeExtended,[ref]$FolderSizeVal)
				$ffFolder | Add-Member -Name "FolderPath" -Value $fpath -MemberType NoteProperty
				$rptobj = "" | Select Name,FolderPath,TotalItemCount,FolderItemSizeMB
				$rptobj.Name = $ffFolder.DisplayName
				$rptobj.FolderPath = $ffFolder.FolderPath
				$rptobj.TotalItemCount = $ffFolder.TotalCount
				$rptobj.FolderItemSizeMB =  [Math]::Round($FolderSizeVal/1024/1024,2)
				$rptCollection += $rptobj
		    } 
		    $fvFolderView.Offset += $fiResult.Folders.Count
		}while($fiResult.MoreAvailable -eq $true)  
		return $rptCollection	
	}
}
####################### 
<# 
.SYNOPSIS 
 Enumerates the Search folder in a Mailbox using the Exchange Web Services API on Exchange 2013,Office365 and Exchange 2016 
 
.DESCRIPTION 
  Enumerates the Search folder in a Mailbox using the Exchange Web Services API on Exchange 2013,Office365 and Exchange 2016 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 Get all the Search Folder in a Mailbox
	
	Get-SearchFolders -MailboxName Mailbox@domain.com
	
	Example 2 Get all the Searchfolder from an Archive

	Get-SearchFolders -MailboxName exporttest@datarumble.com -Archive

	Example 3 to use EWS Impersonation

	Get-SearchFolders -MailboxName Mailbox@domain.com -useImpersonation

#> 
########################
function Get-SearchFolders{

    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [switch]$Archive
    )  
 	Begin
	{
		$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
			$service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
		}
		# Bind to the NON_IPM_ROOT Root folder 
		if($Archive.IsPresent)
		{
			$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRoot,$MailboxName)		
		}
		else
		{
			$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)
		}
		$rptCollection = Enum-SearchFolders -service $service -FolderId $folderid
		write-output $rptCollection
	}

}