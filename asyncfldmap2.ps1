function Connect-Exchange { 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url
    )  
    Begin {
        Load-EWSManagedAPI
		
        ## Set Exchange Version  
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
		  
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
		  
        Handle-SSL	
		  
        ## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
		  
        #CAS URL Option 1 Autodiscover  
        if ($url) {
            $uri = [system.URI] $url
            $service.Url = $uri    
        }
        else {
            $service.AutodiscoverUrl($MailboxName, {$true})  
        }
        Write-host ("Using CAS Server : " + $Service.url)   
		   
        #CAS URL Option 2 Hardcoded  
		  
        #$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
        #$service.Url = $uri    
		  
        ## Optional section for Exchange Impersonation  
		  
        #$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
        if (!$service.URL) {
            throw "Error connecting to EWS"
        }
        else {		
            return $service
        }
    }
}

function Load-EWSManagedAPI {
    param( 
    )  
    Begin {
        ## Load Managed API dll  
        ###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
        $EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
        if (Test-Path $EWSDLL) {
            Import-Module $EWSDLL
        }
        else {
            "$(get-date -format yyyyMMddHHmmss):"
            "This script requires the EWS Managed API 1.2 or later."
            "Please download and install the current version of the EWS Managed API from"
            "http://go.microsoft.com/fwlink/?LinkId=255472"
            ""
            "Exiting Script."
            #exit
        } 
    }
}

function Handle-SSL {
    param( 
    )  
    Begin {
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

    }
}

function Get-AirSyncFolderMappings  {
	    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [string]$url,
        [Parameter(Position=4, Mandatory=$false)] [string]$FolderPath
    )  
 	Process
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
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
        $MsgRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
        #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
        $fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)  
        #Deep Transval will ensure all folders in the search path are returned  
        $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;  
        #The Search filter will exclude any Search Folders  
        $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"ExchangeSyncData")  
        $asFolderRoot = $Service.FindFolders($MsgRoot.Id,$sfSearchFilter,$fvFolderView)  
        if($asFolderRoot.Folders.Count -eq 1){
            #Define Function to convert String to FolderPath  
            function ConvertToString($ipInputString){  
                $Val1Text = ""  
                for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
                        $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
                        $clInt++  
                }  
                return $Val1Text  
            } 
             #Define Extended properties1  
            $PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
            #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
            $fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
            #Deep Transval will ensure all folders in the search path are returned  
            $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
            $psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
            $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
            $CollectionIdProp = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x7C03, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
            $LastModifiedTime = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3008, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
            $AirSyncLastSyncTimeGuid = [Guid]("71035549-0739-4DCB-9163-00F0580DBBDF")
            $AirSyncLastSyncTime = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($AirSyncLastSyncTimeGuid, "AirSync:AirSyncLastSyncTime", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
            #Add Properties to the  Property Set  
            $psPropertySet.Add($PR_Folder_Path);  
            $psPropertySet.Add($CollectionIdProp);
            $psPropertySet.Add($LastModifiedTime);
            $psPropertySet.Add($AirSyncLastSyncTime)	
            $fvFolderView.PropertySet = $psPropertySet;  
            #The Search filter will exclude any Search Folders  
            $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")  
            $fiResult = $null  
            $fldMappingHash =@{}
            #The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
            do {  
                $fiResult = $Service.FindFolders($asFolderRoot.Folders[0].Id,$sfSearchFilter,$fvFolderView)  
                foreach($ffFolder in $fiResult.Folders){ 
                    $asFolderPath = ""
                    $asFolderPath = (GetFolderPath -EWSFolder $ffFolder)
                    "FolderPath : " + $asFolderPath
                    $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
                    $ivItemView.PropertySet =$psPropertySet
                    $fiItems = $null
                    do{ 
                        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Exists($CollectionIdProp)
                        $fiItems = $ffFolder.findItems($SfSearchFilter,$ivItemView)  
                        foreach($itItem in $fiItems.Items){
                  			$collectVal = $null
                            if($itItem.TryGetProperty($CollectionIdProp,[ref]$collectVal)){
                                $HexEntryId = [System.BitConverter]::ToString($collectVal).Replace("-","").Substring(2)
                                $ewsFolderId = ConvertId -HexId ($HexEntryId.SubString(0,($HexEntryId.Length-2)))
                                try{
                                    $fldReport = "" | Select Mailbox,Device,AsFolderPath,MailboxFolderPath,MailboxItem,LastModified,AirSyncLastSyncTime
                                    $fldReport.Mailbox = $MailboxName
                                    $fldReport.Device = $ffFolder.DisplayName
                                    $fldReport.AsFolderPath = $asFolderPath
                                    $fldReport.MailboxItem = $itItem.Subject
                                    $folderMapId= new-object Microsoft.Exchange.WebServices.Data.FolderId($ewsFolderId)   
                                    $MappedFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderMapId,$psPropertySet)
                                    $MappedFolderPath = (GetFolderPath -EWSFolder $MappedFolder)
                                    $fldReport.MailboxFolderPath = $MappedFolderPath
                                    $LastModifiedVal = $null
                                    if($ffFolder.TryGetProperty($LastModifiedTime,[ref]$LastModifiedVal)){
                                        Write-Host ("Last-Modified : " +  $LastModifiedVal.ToLocalTime().ToString())
                                        $fldReport.LastModified = $LastModifiedVal.ToLocalTime().ToString()
                                    }
                                    $AirSyncLastSyncTimeValue = $null
                                    if($itItem.TryGetProperty($AirSyncLastSyncTime,[ref]$AirSyncLastSyncTimeValue)){
                                        $fldReport.AirSyncLastSyncTime = [DateTime]::FromBinary($AirSyncLastSyncTimeValue)
                                    }
                                    Write-Output $fldReport
                                    #$AsFolderReport += $fldReport
                                }
                                catch{
                                    
                                }
                            }
                        }
                        $ivItemView.Offset += $fiItems.Items.Count    
                    }while($fiItems.MoreAvailable -eq $true)                  
                } 
                $fvFolderView.Offset += $fiResult.Folders.Count
            }while($fiResult.MoreAvailable -eq $true)
        }
    }
}

function Get-FolderFromPath{
	param (
	        [Parameter(Position=0, Mandatory=$true)] [string]$FolderPath,
			[Parameter(Position=1, Mandatory=$true)] [string]$SmtpAddress,
			[Parameter(Position=2, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
		  )
	process{
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root  
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$SmtpAddress)   
		$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
		#Split the Search path into an array  
		$fldArray = $FolderPath.Split("\") 
		 #Loop through the Split Array and do a Search for each level of folder 
		for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
	        #Perform search based on the displayname of each folder level 
	        $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
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
			return [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$tfTargetFolder.Id)
		}
		else{
			throw ("Folder Not found")
		}
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

function GetFolderPath{
	param (
		$EWSFolder = "$( throw 'Folder is a mandatory Parameter' )"
	)
	process{
		$foldpathval = $null  
		$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
		if ($EWSFolder.TryGetProperty($PR_Folder_Path,[ref] $foldpathval))  
        {  
            $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
            $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
            $hexString = $hexArr -join ''  
	    $hexString = $hexString.Replace("EFBFBE", "5C")  
            $fpath = ConvertToString($hexString) 
	    return $fpath
        }  
	}
}