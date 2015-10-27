function Connect-Exchange
{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials
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

function GetFolderPaths
{
	param (
	    	[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.FolderId]$rootFolderId,
			[Parameter(Position=1, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
			[Parameter(Position=2, Mandatory=$true)] [PSObject]$FolderCache,
			[Parameter(Position=3, Mandatory=$false)] [String]$FolderPrefix
		  )
	process{
	#Define Extended properties  
	$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
	$PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
	#Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
	$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
	#Deep Transval will ensure all folders in the search path are returned  
	$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
	$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
	$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
	#Add Properties to the  Property Set  
	$psPropertySet.Add($PR_Folder_Path);  
	$psPropertySet.Add($PR_MESSAGE_SIZE_EXTENDED)
	$fvFolderView.PropertySet = $psPropertySet;  
	#The Search filter will exclude any Search Folders  
	$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")  
	$fiResult = $null  
	#The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
	do 
	{  
	    $fiResult = $service.FindFolders($rootFolderId,$sfSearchFilter,$fvFolderView)  
	    foreach($ffFolder in $fiResult.Folders){
	        #Try to get the FolderPath Value and then covert it to a usable String 
			$foldpathval = $null
	        if ($ffFolder.TryGetProperty($PR_Folder_Path,[ref] $foldpathval))  
	        {  
	            $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
	            $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
	            $hexString = $hexArr -join ''  
	            $hexString = $hexString.Replace("FEFF", "5C00")  
	            $fpath = ConvertToString($hexString)  
	        }
			if($FolderCache.ContainsKey($ffFolder.Id.UniqueId) -eq $false)
			{
				if ([string]::IsNullOrEmpty($FolderPrefix)){
					$FolderCache.Add($ffFolder.Id.UniqueId,($fpath))	
				}
				else
				{
					$FolderCache.Add($ffFolder.Id.UniqueId,("\" + $FolderPrefix + $fpath))	
				}
			}
	    } 
	    $fvFolderView.Offset += $fiResult.Folders.Count
	}while($fiResult.MoreAvailable -eq $true)  
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

function Get-MailboxItemStats 
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [string]$KQL,
		[Parameter(Position=3, Mandatory=$false)] [DateTime]$Start,
		[Parameter(Position=4, Mandatory=$false)] [DateTime]$End,
		[Parameter(Position=5, Mandatory=$false)] [Switch]$FolderList
		
    )  
 	Begin
		{
			$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials			
			if((![string]::IsNullOrEmpty($Start)) -band (![string]::IsNullOrEmpty($End))){
				$KQL = "Received:" + $Start.ToString("yyyy-MM-dd") + ".." + $End.ToString("yyyy-MM-dd")
			}
			if(!$FolderList.IsPresent){
				Exec-eDiscoveryKeyWordStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName -Prefix "kind:"
			}
			else{
				Exec-eDiscoveryPreviewItemsStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName
			}
		}
}

function Get-MailboxItemTypeStats 
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
		[Parameter(Position=1, Mandatory=$true)] [String]$ItemType,
		[Parameter(Position=2, Mandatory=$false)] [Switch]$FolderList
    )  
 	Begin
		{
			$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
			$KQL = ""
			if([string]::IsNullOrEmpty($ItemType))
			{
				$KQL = "kind:email OR kind:meetings OR kind:contacts OR kind:tasks OR kind:notes OR kind:IM OR kind:rssfeeds OR kind:voicemail";
			}
			else{
				$KQL = "kind:" + $ItemType
			}	
			
			if(!$FolderList.IsPresent){
				Exec-eDiscoveryKeyWordStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName -Prefix "kind:"
			}
			else{
				Exec-eDiscoveryPreviewItemsStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName
			}
		}
}
function Get-MailboxConversationStats 
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,		
		[Parameter(Position=3, Mandatory=$true)] [PSObject]$ParticipantList,
		[Parameter(Position=4, Mandatory=$false)] [Switch]$FolderList
    )  
 	Begin
		{
			$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
			$KQL = ""
			foreach($Item in $ParticipantList){
				if($KQL -eq ""){
					$KQL = "Participants:" + $Item
				}
				else{
					$KQL += " OR Participants:" + $Item
				}
			}
			if(!$FolderList.IsPresent){
				Exec-eDiscoveryKeyWordStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName -Prefix "kind:"
			}
			else{
				Exec-eDiscoveryPreviewItemsStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName
			}
		}
}
function Get-AttachmentTypeMailboxStats 
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [PSObject]$AttachmentList,
		[Parameter(Position=3, Mandatory=$false)] [string]$AttachmentType,
		[Parameter(Position=4, Mandatory=$false)] [string]$AttachmentName,
		[Parameter(Position=5, Mandatory=$false )] [Switch]$FolderList
    )  
 	Begin
		{
			$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
			$KQL = ""
			if([string]::IsNullOrEmpty($AttachmentList))
			{			
				$AttachmentList = @()
				$AttachmentList += "xlsx"
				$AttachmentList += "docx"
				$AttachmentList += "doc"
				$AttachmentList += "xls"
				$AttachmentList += "pptx"
				$AttachmentList += "ppt"
				$AttachmentList += "txt"
				$AttachmentList += "mp3"
				$AttachmentList += "zip"
				$AttachmentList += "txt"
				$AttachmentList += "wma"
				$AttachmentList += "pdf"
			}
			if(![string]::IsNullOrEmpty($AttachmentType))
			{
				$AttachmentList = @()
				$AttachmentList += $AttachmentType
			}
			else
			{
				if(![string]::IsNullOrEmpty($AttachmentName))
				{
					$KQL = "Attachment:" + $AttachmentName
				}
			}			
			if($KQL -eq ""){
				foreach($Item in $AttachmentList){
					if($KQL -eq ""){
						$KQL = "Attachment:." + $Item
					}
					else{
						$KQL += " OR Attachment:." + $Item
					}
				}			
			}
			if(!$FolderList.IsPresent)
			{
				Exec-eDiscoveryKeyWordStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName -Prefix "Attachment:."
			}
			else
			{			
				Exec-eDiscoveryPreviewItemsStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName
			}
		}
}
function Exec-eDiscoveryKeyWordStats
{
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
		[Parameter(Position=1, Mandatory=$true)] [String]$KQL,
		[Parameter(Position=2, Mandatory=$true)] [String]$SearchableMailboxString,
		[Parameter(Position=3, Mandatory=$true)] [String]$Prefix
	)
	Begin 
		{
			$gsMBResponse = $service.GetSearchableMailboxes($SearchableMailboxString, $false);
			$msbScope = New-Object  Microsoft.Exchange.WebServices.Data.MailboxSearchScope[] $gsMBResponse.SearchableMailboxes.Length
			$mbCount = 0;
			foreach ($sbMailbox in $gsMBResponse.SearchableMailboxes)
			{
			    $msbScope[$mbCount] = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope($sbMailbox.ReferenceId, [Microsoft.Exchange.WebServices.Data.MailboxSearchLocation]::All);
			    $mbCount++;
			}
			$smSearchMailbox = New-Object Microsoft.Exchange.WebServices.Data.SearchMailboxesParameters
			$mbq =  New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery($KQL, $msbScope);
			$mbqa = New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery[] 1
			$mbqa[0] = $mbq
			$smSearchMailbox.SearchQueries = $mbqa;
			$smSearchMailbox.PageSize = 100;
			$smSearchMailbox.PageDirection = [Microsoft.Exchange.WebServices.Data.SearchPageDirection]::Next;
			$smSearchMailbox.PerformDeduplication = $false;           
			$smSearchMailbox.ResultType = [Microsoft.Exchange.WebServices.Data.SearchResultType]::StatisticsOnly;
			$srCol = $service.SearchMailboxes($smSearchMailbox);
			$rptCollection = @()
			if ($srCol[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
			{
				foreach($KeyWorkdStat in $srCol[0].SearchResult.KeywordStats){
					if($KeyWorkdStat.Keyword.Contains(" OR ") -eq $false){
						$rptObj = "" | Select Name,ItemHits,Size
						$rptObj.Name = $KeyWorkdStat.Keyword.Replace($Prefix,"")
						$rptObj.Name = $rptObj.Name.Replace($Prefix.ToLower(),"")
						$rptObj.ItemHits = $KeyWorkdStat.ItemHits
						$rptObj.Size = [System.Math]::Round($KeyWorkdStat.Size /1024/1024,2)
						$rptCollection += $rptObj
					}
				}   
			}
			Write-Output $rptCollection
		
		}
}
function Exec-eDiscoveryKeyWordStatsMultiMailbox
{
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
		[Parameter(Position=1, Mandatory=$true)] [String]$KQL,
		[Parameter(Position=2, Mandatory=$true)] [psObject]$Mailboxes
	)
	Begin 
		{
			$searchArray = @()
			foreach($mailbox in $Mailboxes){
			$gsMBResponse = $service.GetSearchableMailboxes($mailbox, $false);
				foreach ($sbMailbox in $gsMBResponse.SearchableMailboxes)
				{					
					if($sbMailbox.SMTPAddress.ToLower() -eq $mailbox.ToLower()){
						$searchArray +=$sbMailbox
					}										
				}
			}
			$msbScope = New-Object  Microsoft.Exchange.WebServices.Data.MailboxSearchScope[] $searchArray.Count
			$mbCount = 0;
			foreach ($sbMailbox in $searchArray)
			{
			    $msbScope[$mbCount] = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope($sbMailbox.ReferenceId, [Microsoft.Exchange.WebServices.Data.MailboxSearchLocation]::All);
			    $mbCount++;
			}
			$smSearchMailbox = New-Object Microsoft.Exchange.WebServices.Data.SearchMailboxesParameters
			$mbq =  New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery($KQL, $msbScope);
			$mbqa = New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery[] 1
			$mbqa[0] = $mbq
			$smSearchMailbox.SearchQueries = $mbqa;
			$smSearchMailbox.PageSize = 100;
			$smSearchMailbox.PageDirection = [Microsoft.Exchange.WebServices.Data.SearchPageDirection]::Next;
			$smSearchMailbox.PerformDeduplication = $false;           
			$smSearchMailbox.ResultType = [Microsoft.Exchange.WebServices.Data.SearchResultType]::StatisticsOnly;
			$srCol = $service.SearchMailboxes($smSearchMailbox);
			$rptCollection = @()
			if ($srCol[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
			{
				foreach($KeyWorkdStat in $srCol[0].SearchResult.KeywordStats){
						$rptObj = "" | Select Name,ItemHits,Size
						$rptObj.Name = $KeyWorkdStat.Keyword
						$rptObj.ItemHits = $KeyWorkdStat.ItemHits
						$rptObj.Size = [System.Math]::Round($KeyWorkdStat.Size /1024/1024,2)
						$rptCollection += $rptObj
				}   
			}
			Write-Output $rptCollection
		
		}
}
function Exec-eDiscoveryPreviewItemsStats
{
    [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
		[Parameter(Position=1, Mandatory=$true)] [String]$KQL,
		[Parameter(Position=2, Mandatory=$true)] [String]$SearchableMailboxString
		
    )  
 	Begin
	{
		
		$FolderCache = @{}
		# Bind to the MsgFolderRoot folder  
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)   
		$MsgRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		GetFolderPaths -FolderCache $FolderCache -service $service -rootFolderId $MsgRoot.Id
		try
		{
			# Bind to the ArchiveMsgFolderRoot folder  
			$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot,$MailboxName)   
			$ArchiveRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
			GetFolderPaths -FolderCache $FolderCache -service $service -rootFolderId $ArchiveRoot.Id -FolderPrefix "Archive"
			Write-Host ("Mailbox has Archive")
		}
		catch
		{
		
		}
		$gsMBResponse = $service.GetSearchableMailboxes($SearchableMailboxString, $false);
		$msbScope = New-Object  Microsoft.Exchange.WebServices.Data.MailboxSearchScope[] $gsMBResponse.SearchableMailboxes.Length
		$mbCount = 0;
		foreach ($sbMailbox in $gsMBResponse.SearchableMailboxes)
		{
		    $msbScope[$mbCount] = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope($sbMailbox.ReferenceId, [Microsoft.Exchange.WebServices.Data.MailboxSearchLocation]::All);
		    $mbCount++;
		}
		$smSearchMailbox = New-Object Microsoft.Exchange.WebServices.Data.SearchMailboxesParameters
		$mbq =  New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery($KQL, $msbScope);
		$mbqa = New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery[] 1
		$mbqa[0] = $mbq
		$smSearchMailbox.SearchQueries = $mbqa;
		$smSearchMailbox.PageSize = 1000;
		$smSearchMailbox.PageDirection = [Microsoft.Exchange.WebServices.Data.SearchPageDirection]::Next;
		$smSearchMailbox.PerformDeduplication = $false;           
		$smSearchMailbox.ResultType = [Microsoft.Exchange.WebServices.Data.SearchResultType]::PreviewOnly;
		$srCol = $service.SearchMailboxes($smSearchMailbox);
		$rptCollection = @{}

		if ($srCol[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
		{
			Write-Host ("Items Found " + $srCol[0].SearchResult.ItemCount)
		    if ($srCol[0].SearchResult.ItemCount -gt 0)
		    {                  
		        do
		        {
		            $smSearchMailbox.PageItemReference = $srCol[0].SearchResult.PreviewItems[$srCol[0].SearchResult.PreviewItems.Length - 1].SortValue;
		            foreach ($PvItem in $srCol[0].SearchResult.PreviewItems) {
						$rptObj = "" | select FolderPath,TotalItemNumbers,Size
		                if($FolderCache.ContainsKey($PvItem.ParentId.UniqueId)){
							if($rptCollection.ContainsKey($PvItem.ParentId.UniqueId) -eq $false){
								$rptObj = "" | Select FolderPath,TotalItemNumbers,TotalSize
								$rptObj.FolderPath = $FolderCache[$PvItem.ParentId.UniqueId]
								$rptCollection.Add($PvItem.ParentId.UniqueId,$rptObj)
							}
							$rptCollection[$PvItem.ParentId.UniqueId].TotalSize += $PvItem.Size
							$rptCollection[$PvItem.ParentId.UniqueId].TotalItemNumbers++
						}
						else{
						#$ItemId = new-object Microsoft.Exchange.WebServices.Data.ItemId($PvItem.Id)   
							$Item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$PvItem.Id)
							if($FolderCache.ContainsKey($Item.ParentFolderId.UniqueId)){
								$FolderCache.Add($PvItem.ParentId.UniqueId,$FolderCache[$Item.ParentFolderId.UniqueId])
								if($rptCollection.ContainsKey($PvItem.ParentId.UniqueId) -eq $false){
									$rptObj = "" | Select FolderPath,TotalItemNumbers,TotalSize
									$rptObj.FolderPath = $FolderCache[$PvItem.ParentId.UniqueId]
									$rptCollection.Add($PvItem.ParentId.UniqueId,$rptObj)
								}
								$rptCollection[$PvItem.ParentId.UniqueId].TotalSize += $PvItem.Size
								$rptCollection[$PvItem.ParentId.UniqueId].TotalItemNumbers++
							}
						}
		            }                        
		            $srCol = $service.SearchMailboxes($smSearchMailbox);
					Write-Host("Items Remaining : " + $srCol[0].SearchResult.ItemCount);
		        } while ($srCol[0].SearchResult.ItemCount-gt 0 );
		        
		    }		    
		}
		Write-Output $rptCollection.Values 
	}
}
function Exec-eDiscoveryPreviewItemsStatsMultiMailbox
{
    [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
		[Parameter(Position=1, Mandatory=$true)] [String]$KQL,
		[Parameter(Position=2, Mandatory=$true)] [psObject]$Mailboxes
		
    )  
 	Begin
	{
		
		$FolderCache = @{}
		$searchArray = @()
		foreach($mailbox in $Mailboxes){
		$gsMBResponse = $service.GetSearchableMailboxes($mailbox, $false);
			foreach ($sbMailbox in $gsMBResponse.SearchableMailboxes)
			{					
				if($sbMailbox.SMTPAddress.ToLower() -eq $mailbox.ToLower()){
					$searchArray +=$sbMailbox
				}										
			}
		}
		$mbCount =0
		$msbScope = New-Object  Microsoft.Exchange.WebServices.Data.MailboxSearchScope[] $searchArray.Count
		foreach ($sbMailbox in $searchArray)
		{
		    $msbScope[$mbCount] = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope($sbMailbox.ReferenceId, [Microsoft.Exchange.WebServices.Data.MailboxSearchLocation]::All);
		    $mbCount++;
		}
		$smSearchMailbox = New-Object Microsoft.Exchange.WebServices.Data.SearchMailboxesParameters
		$mbq =  New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery($KQL, $msbScope);
		$mbqa = New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery[] 1
		$mbqa[0] = $mbq
		$smSearchMailbox.SearchQueries = $mbqa;
		$smSearchMailbox.PageSize = 1;
		$smSearchMailbox.PageDirection = [Microsoft.Exchange.WebServices.Data.SearchPageDirection]::Next;
		$smSearchMailbox.PerformDeduplication = $false;           
		$smSearchMailbox.ResultType = [Microsoft.Exchange.WebServices.Data.SearchResultType]::PreviewOnly;
		$srCol = $service.SearchMailboxes($smSearchMailbox);
		$rptCollection = @{}

		if ($srCol[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
		{
			Write-Host ("Total Number of Items Found " + $srCol[0].SearchResult.ItemCount)
		    if ($srCol[0].SearchResult.ItemCount -gt 0)
		    {
				write-output $srCol[0].SearchResult.MailboxStats
		        
		    }		    
		}
		Write-Output $rptCollection.Values 
	}
}
function Exec-eDiscoveryPreviewItems
{
    [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
		[Parameter(Position=1, Mandatory=$true)] [String]$KQL,
		[Parameter(Position=2, Mandatory=$true)] [String]$SearchableMailboxString
		
    )  
 	Begin
	{
		
		$FolderCache = @{}
		# Bind to the MsgFolderRoot folder  
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)   
		$MsgRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		GetFolderPaths -FolderCache $FolderCache -service $service -rootFolderId $MsgRoot.Id
		try
		{
			# Bind to the ArchiveMsgFolderRoot folder  
			$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot,$MailboxName)   
			$ArchiveRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
			GetFolderPaths -FolderCache $FolderCache -service $service -rootFolderId $ArchiveRoot.Id -FolderPrefix "Archive"
			Write-Host ("Mailbox has Archive")
		}
		catch
		{
		
		}
		$gsMBResponse = $service.GetSearchableMailboxes($SearchableMailboxString, $false);
		$msbScope = New-Object  Microsoft.Exchange.WebServices.Data.MailboxSearchScope[] $gsMBResponse.SearchableMailboxes.Length
		$mbCount = 0;
		foreach ($sbMailbox in $gsMBResponse.SearchableMailboxes)
		{
		    $msbScope[$mbCount] = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope($sbMailbox.ReferenceId, [Microsoft.Exchange.WebServices.Data.MailboxSearchLocation]::All);
		    $mbCount++;
		}
		$smSearchMailbox = New-Object Microsoft.Exchange.WebServices.Data.SearchMailboxesParameters
		$mbq =  New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery($KQL, $msbScope);
		$mbqa = New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery[] 1
		$mbqa[0] = $mbq
		$smSearchMailbox.SearchQueries = $mbqa;
		$smSearchMailbox.PageSize = 100;
		$smSearchMailbox.PageDirection = [Microsoft.Exchange.WebServices.Data.SearchPageDirection]::Next;
		$smSearchMailbox.PerformDeduplication = $false;           
		$smSearchMailbox.ResultType = [Microsoft.Exchange.WebServices.Data.SearchResultType]::PreviewOnly;
		$srCol = $service.SearchMailboxes($smSearchMailbox);
		$rptCollection = @{}

		if ($srCol[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
		{
			Write-Host ("Items Found " + $srCol[0].SearchResult.ItemCount)
		    if ($srCol[0].SearchResult.ItemCount -gt 0)
		    {                  
		        do
		        {
		            $smSearchMailbox.PageItemReference = $srCol[0].SearchResult.PreviewItems[$srCol[0].SearchResult.PreviewItems.Length - 1].SortValue;
		            foreach ($PvItem in $srCol[0].SearchResult.PreviewItems) {
						if($FolderCache.ContainsKey($PvItem.ParentId.UniqueId)){
							$PvItem | Add-Member -NotePropertyName FolderPath -NotePropertyValue $FolderCache[$PvItem.ParentId.UniqueId] 
							Write-Output $PvItem
						}
						else
						{
							$Item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$PvItem.Id)
							if($FolderCache.ContainsKey($Item.ParentFolderId.UniqueId)){
								$FolderCache.Add($PvItem.ParentId.UniqueId,$FolderCache[$Item.ParentFolderId.UniqueId])
								$PvItem | Add-Member -NotePropertyName FolderPath -NotePropertyValue $FolderCache[$PvItem.ParentId.UniqueId]
							}
							else{
								$PvItem | Add-Member -NotePropertyName FolderPath -NotePropertyValue ""
							}
							Write-Output $PvItem
						}
		            }                        
		            $srCol = $service.SearchMailboxes($smSearchMailbox);
					Write-Host("Items Remaining : " + $srCol[0].SearchResult.ItemCount);
		        } while ($srCol[0].SearchResult.ItemCount-gt 0 );
		        
		    }		    
		}
	}
}
function Make-UniqueFileName{
    param(
		[Parameter(Position=0, Mandatory=$true)] [string]$FileName
	)
	Begin
	{
	
	$directoryName = [System.IO.Path]::GetDirectoryName($FileName)
    $FileDisplayName = [System.IO.Path]::GetFileNameWithoutExtension($FileName);
    $FileExtension = [System.IO.Path]::GetExtension($FileName);
    for ($i = 1; ; $i++){
            
            if (![System.IO.File]::Exists($FileName)){
				return($FileName)
			}
			else{
					$FileName = [System.IO.Path]::Combine($directoryName, $FileDisplayName + "(" + $i + ")" + $FileExtension);
			}                
            
			if($i -eq 10000){throw "Out of Range"}
        }
	}
}
function Get-MailboxAttachments
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [PSObject]$AttachmentList,
		[Parameter(Position=3, Mandatory=$false)] [string]$AttachmentType,
		[Parameter(Position=4, Mandatory=$false)] [string]$AttachmentName,
		[Parameter(Position=5, Mandatory=$false)] [switch]$Download,
		[Parameter(Position=6, Mandatory=$false)] [string]$DownloadDirectory
    )  
 	Begin
		{
			$KQL = ""
			if([string]::IsNullOrEmpty($DownloadDirectory)){
				$DownloadDirectory = (Get-Location).Path
			}
			$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
			if([string]::IsNullOrEmpty($AttachmentType) -band [string]::IsNullOrEmpty($AttachmentName)){
				if([string]::IsNullOrEmpty($AttachmentList))
				{
					$AttachmentList = @()
					$AttachmentList += "xlsx"
					$AttachmentList += "docx"
					$AttachmentList += "doc"
					$AttachmentList += "xls"
					$AttachmentList += "pptx"
					$AttachmentList += "ppt"
					$AttachmentList += "txt"
					$AttachmentList += "mp3"
					$AttachmentList += "zip"
					$AttachmentList += "txt"
					$AttachmentList += "wma"
					$AttachmentList += "pdf"
				}
			}
			else{
				$AttachmentList = @()
				if(![string]::IsNullOrEmpty($AttachmentType)){
					$AttachmentList += $AttachmentType
				}
				else{
					$KQL = "Attachment:" + $AttachmentName
				}
			}	
			if($KQL -eq ""){
				foreach($Item in $AttachmentList){
					if($KQL -eq ""){
						$KQL = "Attachment:." + $Item
					}
					else{
						$KQL += " OR Attachment:." + $Item
					}
				}			
			}
			if($Download)
			{
				$attachemtItems = Exec-eDiscoveryPreviewItems -service $service -KQL $KQL -SearchableMailboxString $MailboxName 
				foreach($AttachmentItem in $attachemtItems){
					Write-Host ("Processing Item " + $AttachmentItem.Subject)
					$Item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$AttachmentItem.Id)
					foreach($attach in $Item.Attachments){
						if($attach -is [Microsoft.Exchange.WebServices.Data.FileAttachment]){
							if(![string]::IsNullOrEmpty($attach.Name) -band $attach.IsInline -eq $false){
								$attach.Load()	
								$FileName = Make-UniqueFileName -FileName ($DownloadDirectory + “\” + $attach.Name.ToString())
								$fiFile = new-object System.IO.FileStream($FileName, [System.IO.FileMode]::Create)
								$fiFile.Write($attach.Content, 0, $attach.Content.Length)
								$fiFile.Close()
								write-host ("Downloaded Attachment : " + $FileName)
							}
						}
						if ("Microsoft.Exchange.WebServices.Data.ReferenceAttachment" -as [type]) {
							if($attach -is [Microsoft.Exchange.WebServices.Data.ReferenceAttachment]){
								$SharePointClientDll = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SharePoint Client Components\'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Location') + "ISAPI\Microsoft.SharePoint.Client.dll")
								Add-Type -Path $SharePointClientDll 
								$DownloadURI = New-Object System.Uri($attach.AttachLongPathName);
 								$SharepointHost = "https://" + $DownloadURI.Host
  								$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($SharepointHost)
								$soCredentials =  New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.UserName.ToString(),$Credentials.password)
  								$clientContext.Credentials = $soCredentials;
  								$FileName = Make-UniqueFileName -FileName ($DownloadDirectory + “\” + $attach.Name.ToString())
  								$fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($clientContext, $DownloadURI.LocalPath);
 								$fstream = New-Object System.IO.FileStream($FileName, [System.IO.FileMode]::Create);
 								$fileInfo.Stream.CopyTo($fstream)
 								$fstream.Flush()
  								$fstream.Close()
 								Write-Host ("File downloaded to " + ($FileName))
							}
						}
					}
				}
			}
			else
			{
				$attachemtItems = Exec-eDiscoveryPreviewItems -service $service -KQL $KQL -SearchableMailboxString $MailboxName 
				foreach($AttachmentItem in $attachemtItems){
					Write-Host ("Processing Item " + $AttachmentItem.Subject)
					$Item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$AttachmentItem.Id)
					foreach($attach in $Item.Attachments){
						Write-Host $attach
					}
				}
			}
		}
}
####################### 
<# 
.SYNOPSIS 
 Peforms a Multi Mailbox KeyWordStats Search using eDiscovery and the Exchange Web Services API on Exchange 2013,Office365 and Exchange 2016 
 
.DESCRIPTION 
  Peforms a Multi Mailbox KeyWordStats Search using eDiscovery and the Exchange Web Services API on Exchange 2013,Office365 and Exchange 2016 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 To search for the number of Contact in a group of Mailbox
	 Search-MultiMailboxesKeyWordStats -Mailboxes 

#> 
########################
function Search-MultiMailboxesKeyWordStats{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [PSObject]$Mailboxes,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$true)] [String]$QueryString
    )  
 	Begin
	 {
	 	$service = Connect-Exchange -MailboxName $Mailboxes[0] -Credentials $Credentials
		Exec-eDiscoveryKeyWordStatsMultiMailbox -service $service -Mailboxes $Mailboxes -KQL $QueryString
	 }
}
####################### 
<# 
.SYNOPSIS 
 Peforms a Multi Mailbox Statistics Search using eDiscovery and the Exchange Web Services API on Exchange 2013,Office365 and Exchange 2016 
 
.DESCRIPTION 
  Peforms a Multi Mailbox Statistics Search using eDiscovery and the Exchange Web Services API on Exchange 2013,Office365 and Exchange 2016 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 To search for the number of Contact in a group of Mailbox
	 Search-MultiMailboxesItemStats -Mailboxes @('Mailbox@domain.com','mailbox2@domain.com')

#> 
########################
function Search-MultiMailboxesItemStats{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [PSObject]$Mailboxes,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$true)] [String]$QueryString
    )  
 	Begin
	 {
	 	$service = Connect-Exchange -MailboxName $Mailboxes[0] -Credentials $Credentials
		Exec-eDiscoveryPreviewItemsStatsMultiMailbox -service $service -Mailboxes $Mailboxes -KQL $QueryString
	 }
}