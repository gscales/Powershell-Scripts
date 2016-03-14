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
function Get-FolderFromPath{
	param (
			[Parameter(Position=0, Mandatory=$true)] [string]$FolderPath,
			[Parameter(Position=1, Mandatory=$true)] [string]$MailboxName,
			[Parameter(Position=2, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
			[Parameter(Position=3, Mandatory=$false)] [Microsoft.Exchange.WebServices.Data.PropertySet]$PropertySet
		  )
	process{
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root  
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)   
		$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
		#Split the Search path into an array  
		$fldArray = $FolderPath.Split("\") 
		 #Loop through the Split Array and do a Search for each level of folder 
		for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
	        #Perform search based on the displayname of each folder level 
	        $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
			if(![string]::IsNullOrEmpty($PropertySet)){
				$fvFolderView.PropertySet = $PropertySet
			}
	        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$fldArray[$lint]) 
	        $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView) 
	        if ($findFolderResults.TotalCount -gt 0){ 
	            foreach($folder in $findFolderResults.Folders){ 
	                $tfTargetFolder = $folder                
	            } 
	        } 
	        else{ 
	            Write-host ("Error Folder Not Found check path and try again")  
	            $tfTargetFolder = $null  
	            break  
	        }     
	    }  
		if($tfTargetFolder -ne $null){
			return [Microsoft.Exchange.WebServices.Data.Folder]$tfTargetFolder
		}
		else{
			throw ("Folder Not found")
		}
	}
}
function Get-ConversationStats
{
    param(
	    [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [string]$url,
		[Parameter(Position=4, Mandatory=$true)] [Int32]$Period,
		[Parameter(Position=5, Mandatory=$false)] [String]$FolderPath
    )  
 	Begin
	{
		$Script:rptcollection = @{}
		$Script:cnvrptcollection = @{}
		$Script:lmouth = @{}
		if($url){
			$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url 
		}
		else{
			$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		}
		if($useImpersonation.IsPresent)
		{
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		}
		if($FolderPath)
		{
			$FolderToProcess = Get-FolderFromPath -MailboxName $MailboxName -FolderPath $FolderPath 
		}
		else
		{
			$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
			$FolderToProcess = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		$cnvItemView = New-Object Microsoft.Exchange.WebServices.Data.ConversationIndexedItemView(1000)
		$exit = $false;
		$PeriodDate = (Get-Date).AddDays(-$Period)				
		do
		{
			$cnvs = $service.FindConversation($cnvItemView, $FolderToProcess.Id);
			Write-Host ("Number of Conversation items returned " + $cnvs.Count)
			foreach($cnv in $cnvs){
				if($cnv.GlobalMessageCount -gt 2 -band $cnv.UniqueSenders.Count -gt 2)
				{
					if($cnv.LastDeliveryTime -gt $PeriodDate)
					{
						if($Script:rptcollection.Contains($cnv.Id)-eq $false)
						{					
							$Script:rptcollection.Add($cnv.Id,$cnv)
						}
					}					
				}
			}			
			$cnvItemView.Offset += $cnvs.Count
		}while($cnvs.Count -gt 0)
		#Process tracked conversation
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)  
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ConversationId)
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Attachments)
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender)
		$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
		$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.ItemId" -as "Type")
		$BatchItems = [Activator]::CreateInstance($type)
		foreach($tracked in $Script:rptcollection.Values)
		{
			Write-Host ("Processing Topic " + $tracked.Topic)
			$rptObj = "" | Select Subject,Started,LastMessage,Messages,TotalSize,DurationHours,NumberofAttachments,Initiator,InitiatorResponses,Participants,LoudMouth,LoudMouthResponses
			$rptObj.Subject = $tracked.Topic
			$rptObj.Started = $tracked.LastDeliveryTime
			$rptObj.LastMessage = $tracked.LastDeliveryTime
			$rptObj.Messages = $tracked.GlobalMessageCount
			$rptObj.Participants = $tracked.UniqueSenders.Count 
			$rptObj.TotalSize = $tracked.GlobalSize	
			$rptObj.LoudMouthResponses = 0
			$Script:cnvrptcollection.Add($tracked.Id.UniqueId,$rptObj)
			foreach($ItemId in $tracked.GlobalItemIds)
			{
				$BatchItems.Add($ItemId)
			}
			if($BatchItems.Count -gt 500)
			{

				$cnvItems = $service.BindToItems($BatchItems,$psPropset)
				foreach($cnvItem in $cnvItems)
				{
					if($cnvItem.Item.ConversationId -ne $null)
					{
						if($cnvItem.Item.Subject -eq $Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].Subject)
						{
							$Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].Initiator = $cnvItem.Item.Sender.Address
						}
						if($Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].Started -gt $cnvItem.Item.DateTimeReceived)
						{
							$Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].NumberofAttachments = $cnvItem.Item.Attachments.Count
							$Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].Started = $cnvItem.Item.DateTimeReceived
							$dur = $Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].LastMessage - $Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].Started
							$Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].DurationHours = [System.Math]::Round($dur.TotalHours,0)
						}
						if($Script:lmouth.ContainsKey($cnvItem.Item.ConversationId.UniqueId))
						{
							if($Script:lmouth[$cnvItem.Item.ConversationId.UniqueId].ContainsKey($cnvItem.Item.Sender.Address))
							{
								$Script:lmouth[$cnvItem.Item.ConversationId.UniqueId][$cnvItem.Item.Sender.Address]++
							}
							else
							{
								$Script:lmouth[$cnvItem.Item.ConversationId.UniqueId].Add($cnvItem.Item.Sender.Address,1)
							}
						}
						else
						{
							$cnvHash = @{}
							$cnvHash.Add($cnvItem.Item.Sender.Address,1)
							$Script:lmouth.Add($cnvItem.Item.ConversationId.UniqueId,$cnvHash)
						}
					}
				}	
				$BatchItems.Clear()
			}
		
		}
		if($BatchItems.Count -gt 0)
		{
			$cnvItems = $service.BindToItems($BatchItems,$psPropset)
			foreach($cnvItem in $cnvItems)
			{
					if($cnvItem.Item.ConversationId -ne $null)
					{
						if($cnvItem.Item.Subject -eq $Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].Subject)
						{
							$Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].Initiator = $cnvItem.Item.Sender.Address
						}
						if($Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].Started -gt $cnvItem.Item.DateTimeReceived)
						{
							$Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].NumberofAttachments = $cnvItem.Item.Attachments.Count
							$Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].Started = $cnvItem.Item.DateTimeReceived
							$dur = $Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].LastMessage - $Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].Started
							$Script:cnvrptcollection[$cnvItem.Item.ConversationId.UniqueId].DurationHours = [System.Math]::Round($dur.TotalHours,0)
						}
						if($Script:lmouth.ContainsKey($cnvItem.Item.ConversationId.UniqueId))
						{
							if($Script:lmouth[$cnvItem.Item.ConversationId.UniqueId].ContainsKey($cnvItem.Item.Sender.Address))
							{
								$Script:lmouth[$cnvItem.Item.ConversationId.UniqueId][$cnvItem.Item.Sender.Address]++
							}
							else
							{
								$Script:lmouth[$cnvItem.Item.ConversationId.UniqueId].Add($cnvItem.Item.Sender.Address,1)
							}
						}
						else
						{
							$cnvHash = @{}
							$cnvHash.Add($cnvItem.Item.Sender.Address,1)
							$Script:lmouth.Add($cnvItem.Item.ConversationId.UniqueId,$cnvHash)
						}
					}
			}	
			$BatchItems.Clear()
		}
		foreach($key in $Script:cnvrptcollection.Keys)
		{
			if($Script:lmouth.ContainsKey($key))
			{
			
				foreach($lmonthVal in $Script:lmouth[$key].Keys)
				{
					if($Script:cnvrptcollection[$key].LoudMouthResponses -lt $Script:lmouth[$key][$lmonthVal])
					{
						if($lmonthVal -ne $Script:cnvrptcollection[$key].Initiator)
						{
							$Script:cnvrptcollection[$key].LoudMouthResponses = $Script:lmouth[$key][$lmonthVal]
							$Script:cnvrptcollection[$key].LoudMouth = $lmonthVal
						}
					}
					if($lmonthVal -eq $Script:cnvrptcollection[$key].Initiator)
					{
						$Script:cnvrptcollection[$key].InitiatorResponses = $Script:lmouth[$key][$lmonthVal]
					}
				}
			}
		}
		$fileName =  "ConversationReport" + (Get-Date).ToString("yyyy-MM-dd-hh-mmm") + ".csv"
		$Script:cnvrptcollection.Values | Export-Csv -NoTypeInformation -Path ("c:\temp\" + $fileName)
		Write-Host ("Report written to " + ("c:\temp\" + $fileName))
	}

}



