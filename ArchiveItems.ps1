function Connect-Exchange{ 
    param( 
    		[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
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
function Invoke-ArchiveItems{
	param (
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=3, Mandatory=$false)] [String]$FolderPath,
		[Parameter(Position=4, Mandatory=$true)] [DateTime]$queryTime 
		 )
	process{
        
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRoot,$MailboxName)
		$ArchiveFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		if($ArchiveFolder -is [Microsoft.Exchange.WebServices.Data.Folder]){
			$AQSString = "System.Message.DateReceived:<"+$queryTime.ToString("yyyy-MM-dd")  
			$TargetFolder = Get-FolderFromPath -MailboxName $MailboxName -FolderPath $FolderPath -Service $service
			if($TargetFolder -is [Microsoft.Exchange.WebServices.Data.Folder]){
				$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)   
				$fiItems = $null
				$type = ("System.Collections.Generic.List" + '`' + 1) -as "Type"
				$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.ItemId" -as "Type")
				$archiveItems = [Activator]::CreateInstance($type)
				do{   
					$fiItems = $service.FindItems($TargetFolder.Id,$AQSString,$ivItemView)   
					#[Void] $service.LoadPropertiesForItems($fiItems,$psPropset)
					foreach($Item in $fiItems.Items){
							$archiveItems.Add($Item.Id)      
					}
					$ivItemView.Offset += $fiItems.Items.Count
				}
				while($fiItems.MoreAvailable -eq $true)
				$type = ("System.Collections.Generic.List" + '`' + 1) -as "Type"
				$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.ItemId" -as "Type")
				$batchArchive = [Activator]::CreateInstance($type)
				foreach($itItemID in $archiveItems){
					$batchArchive.add($itItemID)
					if($batchArchive.Count -eq 100){
						$Rcount = 0;
						$ArchiveResponses = $service.ArchiveItems($batchArchive,$TargetFolder.Id)
						foreach ($ArchiveResponse in $ArchiveResponses) {
							if ($ArchiveResponse.Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success){  
								$Rcount++  
							}  
							else{
								write-host $ArchiveResponse.ErrorMessage
							}
							Write-Host -ForegroundColor Green ("Sucessfully Archived " + $Rcount) 
						}
							$batchArchive.clear()
					}
				}
				if($batchArchive.Count -gt 0){
					$Rcount = 0;
					$ArchiveResponses = $service.ArchiveItems($batchArchive,$TargetFolder.Id)
					foreach ($ArchiveResponse in $ArchiveResponses) {
						if ($ArchiveResponse.Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success){  
							$Rcount++  
						}  
						else{
						write-host $ArchiveResponse.ErrorMessage
						}
					}
					Write-Host -ForegroundColor Green ("Sucessfully Archived " + $Rcount) 
					$batchArchive.clear()
				} 	
			}
			else
			{
				Write-Host "Target Folder not Found"			
			}
		}
		else
		{
			Write-Host "No Archive Found"			
		}
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
			return ,$tfTargetFolder
		}
		else{
			throw ("Folder Not found")
		}
	}
}




function Get-ArchiveFolder{
    param( 
     [Parameter(Position=0, Mandatory=$true)] [String]$MailboxName,
     [Parameter(Position=1, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )
 Begin
 {
   # Bind to the RootFolder Folder
   $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
   $RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
   $sfFolderSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM.Configuration.OWA.UserOptions") 
   #Define ItemView to retrive just 1000 Items    
   $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
   $ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
   $fiItems = $service.FindItems($RootFolder.Id,$sfFolderSearchFilter,$ivItemView) 
   if($fiItems.Items.Count -eq 1)
   {
    $UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, "OWA.UserOptions", $RootFolder.Id, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)  
    if($UsrConfig.Dictionary.ContainsKey("ArchiveFolderId"))
    {
         $idVal = $UsrConfig.Dictionary["ArchiveFolderId"]
         $ArchivefolderId= new-object Microsoft.Exchange.WebServices.Data.FolderId($idVal)   
         $ArchiveFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$ArchivefolderId)
    	 return,$ArchiveFolder 
    }
    else
    {
        write-host ("No Arcive Folder Mailbox")
        return,$null
    }  
    
   }
   else
   {
        write-host ("No Arcive Folder Mailbox")
        return,$null
   }
 }
}