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
		#Write-host ("Using CAS Server : " + $Service.url)   
		   
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

function Enum-PublicFolders
{
	param (
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [String]$url
		  )
	process
	{
		$Script:rptCollection = @()
		$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
			$service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
		}
		$publicFolderRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
		Process-Folder -Folder $publicFolderRoot
		$Script:rptCollection | Export-Csv -Path c:\Temp\pfReport.csv -NoTypeInformation
	}
}

function Enum-MailPublicFolders
{
	param (
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [String]$url
		  )
	process
	{
		$Script:rptCollection = @()
		$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
			$service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
		}
		$publicFolderRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
		Process-FolderMailEnabled -Folder $publicFolderRoot
		Write-Output $Script:rptCollection
	}
}


function Process-FolderMailEnabled
{
	param
	(
    	[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.Folder]$Folder
		
	)
	process
	{
		$PidTagMessageSizeExtended = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0xe08,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
		$PidTagLocalCommitTimeMax = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x670A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime);  
		$PR_HAS_RULES = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x663A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean);  
		$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
		$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
		$PR_PF_PROXY = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x671D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
		$PR_ENTRYID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FFF,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)  
		#Add Properties to the  Property Set  
		$psPropertySet.Add($PR_Folder_Path); 
		$psPropertySet.Add($PR_HAS_RULES);
		$psPropertySet.Add($PidTagLocalCommitTimeMax);
		$psPropertySet.Add($PidTagMessageSizeExtended);
		$psPropertySet.Add($PR_PF_PROXY);
		$psPropertySet.Add($PR_ENTRYID);
		$fvFolderView.PropertySet = $psPropertySet; 
		#Deep Transval will ensure all folders in the search path are returned  
		$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;  
		do {  
		    $fiResult = $Service.FindFolders($Folder.Id,$fvFolderView)  
		    foreach($ffFolder in $fiResult.Folders){  
				$proxyguid = $null
				if($ffFolder.TryGetProperty($PR_PF_PROXY,[ref]$proxyguid)){
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
					$EntryIdVal = $null
					[Void]$ffFolder.TryGetProperty($PR_ENTRYID,[ref]$EntryIdVal)
					Write-Host ("Found Mail-Eanbled Folder " + $fpath)
					$Guidval = new-object -TypeName System.Guid -ArgumentList (,$proxyguid)
					$rptObj = "" | Select FolderPath,Guid,EntryId
					$rptObj.FolderPath = $fpath
					$rptObj.Guid = $Guidval.ToString("D")
					$rptObj.EntryId =  [System.BitConverter]::ToString($EntryIdVal).Replace("-","")
					$Script:rptCollection += $rptObj
				}
				if($ffFolder.ChildFolderCount -gt 0)
				{
					Process-FolderMailEnabled -Folder $ffFolder 
				}
				
		    } 
		    $fvFolderView.Offset += $fiResult.Folders.Count
		}while($fiResult.MoreAvailable -eq $true)  
		
	}
	
}


function Process-Folder
{
	param
	(
    	[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.Folder]$Folder
		
	)
	process
	{
		$PidTagMessageSizeExtended = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0xe08,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
		$PidTagLocalCommitTimeMax = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x670A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime);  
		$PR_HAS_RULES = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x663A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean);  
		$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
		$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
		#Add Properties to the  Property Set  
		$psPropertySet.Add($PR_Folder_Path); 
		$psPropertySet.Add($PR_HAS_RULES);
		$psPropertySet.Add($PidTagLocalCommitTimeMax);
		$psPropertySet.Add($PidTagMessageSizeExtended);
		$fvFolderView.PropertySet = $psPropertySet; 
		#Deep Transval will ensure all folders in the search path are returned  
		$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;  
		do {  
		    $fiResult = $Service.FindFolders($Folder.Id,$fvFolderView)  
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
		       # write-host ("Processing FolderPath : " + $fpath)
				$LocalCommitTimeMax = $null
				[Void]$ffFolder.TryGetProperty($PidTagLocalCommitTimeMax,[ref] $LocalCommitTimeMax)
				$hasRules = $null;				
		        [Void]$ffFolder.TryGetProperty($PR_HAS_RULES,[ref] $hasRules)  
				$folderSize = $null;
				[Void]$ffFolder.TryGetProperty($PidTagMessageSizeExtended,[ref] $folderSize)  
				$rptObj = "" | Select FolderPath,FolderClass,HasRules,ActiveRules,DisabledRules,LastCommitTime,TotalItemCount,FolderSize
				$rptObj.FolderPath = $fpath
				$rptObj.FolderClass = $ffFolder.FolderClass
				$rptObj.HasRules = $hasRules
				$rptObj.ActiveRules = 0
				$rptObj.DisabledRules = 0
				$rptObj.LastCommitTime = $LocalCommitTimeMax
				$rptObj.FolderSize = [Math]::Round($folderSize /1024/1024,2)
				$rptObj.TotalItemCount = $ffFolder.TotalCount
				if($ffFolder.ChildFolderCount -gt 0)
				{
					Process-Folder -Folder $ffFolder 
				}
				if($hasRules)
				{
					Write-Host ($fpath + " : " + $hasRules)
					$PR_RULE_MSG_STATE = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65E9,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer); 
					$PR_EXTENDED_RULE_ACTIONS = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E99,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary); 
					$PR_EXTENDED_RULE_CONDITION = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E9A,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary); 
					$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
					$psPropertySet.Add($PR_RULE_MSG_STATE)
					$psPropertySet.Add($PR_EXTENDED_RULE_ACTIONS)
					$psPropertySet.Add($PR_EXTENDED_RULE_CONDITION)
					#Define ItemView to retrive just 1000 Items    
					$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000) 
					$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
					$fiItems = $null    
					do{    
					    $fiItems = $service.FindItems($ffFolder.Id,$ivItemView)    
					    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
					    foreach($Item in $fiItems.Items){      
							if($Item.ItemClass -eq "IPM.Rule.Version2.Message")
							{
								$Item.Load($psPropertySet)
								$RuleStateValue = $null
								[Void]$Item.TryGetProperty($PR_RULE_MSG_STATE,[ref]$RuleStateValue)
								if($RuleStateValue -eq 0)
								{
									Write-Host -ForegroundColor Red "Rule Disabled"
									$rptObj.DisabledRules++
									
								}
								else
								{
									$rptObj.ActiveRules++
									Write-Host -ForegroundColor Green ("Rule State " + $RuleStateValue)
								}
								$PR_Action_val = $null
								[Void]$Item.TryGetProperty($PR_EXTENDED_RULE_ACTIONS,[ref]$PR_Action_val)
							}       
					    }    
					    $ivItemView.Offset += $fiItems.Items.Count    
					}while($fiItems.MoreAvailable -eq $true) 
				}
				$Script:rptCollection += $rptObj
		    } 
		    $fvFolderView.Offset += $fiResult.Folders.Count
		}while($fiResult.MoreAvailable -eq $true)  
		
	}
	
}

#Define Function to convert String to FolderPath  
function ConvertToString($ipInputString){  
    $Val1Text = ""  
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
            $clInt++  
    }  
    return $Val1Text  
} 

function Enum-RuleObjects
{
	param
	(
    	[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.Folder]$Folder		
	)
	process
	{

	}
}

function Parse-ActionRule
{
	param
	(
    	[Parameter(Position=0, Mandatory=$true)] [Byte[]]$PropVal		
	)
	process
	{
		$Stream = new-object System.IO.MemoryStream (,$PropVal)
		$NoOfNamedProps = New-Object Byte[] 2
        $Stream.Read($NoOfNamedProps, 0, 2);
		Write-host ("Number of Named Props : " + [System.BitConverter]::ToInt16($NoOfNamedProps,0))
	}
}



