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

function Get-FolderPath{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.Folder]$Folder
    )  
 	Begin
	{
		$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
		$fpath = ""
		$foldpathval = $null
		if ($Folder.TryGetProperty($PR_Folder_Path,[ref] $foldpathval))  
		{  
		    $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
		    $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
		    $hexString = $hexArr -join ''  
		    $hexString = $hexString.Replace("FEFF", "5C00")  
		    $fpath = ConvertToString($hexString)  
			
        } 
		return $fpath
	}
}

function Report-TaggedFolders
{
    param( 
	    [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [string]$url
    )  
 	Begin
	{	
	    #GetRetentionTag
		$retentionTags = @{}
		Get-RetentionPolicyTag | ForEach-Object{
			$Cval = $_
			$retentionTagGUID = "{$($Cval.RetentionId.ToString())}"
			$policyTagGUID = new-Object Guid($retentionTagGUID)
			$retentionTags.Add($policyTagGUID.ToString().ToLower(),$Cval.Name)
		}
		$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		#PR_POLICY_TAG 0x3019
        $PolicyTag = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3019,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        #PR_RETENTION_FLAGS 0x301D   
        $RetentionFlags = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        #PR_RETENTION_PERIOD 0x301A
        $RetentionPeriod = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301A,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
		$PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);  

		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)   
		$MsgRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		$Folders = Get-Folders -RootFolder $MsgRoot -service $service -IsTagged:$true
		$rptCollection = @()
		foreach($folder in $Folders){
			$rptObj = "" | Select FolderName,FolderClass,Path,ItemCount,FolderTagItemCount,FolderItemSize,PolicyTag,RetentionFlags,RetentionPeriod
			$rptObj.FolderName = $folder.DisplayName
			$rptObj.FolderClass = $Folder.FolderClass
			$rptObj.Path = (Get-FolderPath -Folder $folder)
			$rptObj.ItemCount = $folder.TotalCount
			$folderSize = $null
			if($folder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED,[ref] $folderSize)){
				$rptObj.FolderItemSize = [Int64][System.Math]::Round(($folderSize/1024/1024),2)  
			}
			$prop1Val = $null
			if($folder.TryGetProperty($PolicyTag,[ref] $prop1Val))
			{
				$rtnStringVal = ""
				if($prop1Val -ne $null){
				   $rtnStringVal =	[System.BitConverter]::ToString($prop1Val).Replace("-","");
				   $rtnStringVal = $rtnStringVal.Substring(6,2) + $rtnStringVal.Substring(4,2) + $rtnStringVal.Substring(2,2) + $rtnStringVal.Substring(0,2) + "-" + $rtnStringVal.Substring(10,2) + $rtnStringVal.Substring(8,2) + "-" + $rtnStringVal.Substring(14,2) + $rtnStringVal.Substring(12,2) + "-" + $rtnStringVal.Substring(16,2) + $rtnStringVal.Substring(18,2) + "-" +$rtnStringVal.Substring(20,12)
				}
				if($retentionTags.ContainsKey($rtnStringVal.ToLower())){
					$rptObj.PolicyTag = $retentionTags[$rtnStringVal.ToLower()]
				}
				#Get Tagged Item Count
				$sfItemSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PolicyTag,[System.Convert]::ToBase64String($prop1Val))
				$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)    
				$fiItems = $service.FindItems($Folder.Id,$sfItemSearchFilter,$ivItemView)
				$rptObj.FolderTagItemCount = $fiItems.TotalCount

			}
			$prop2Val = $null
			if($folder.TryGetProperty($RetentionFlags,[ref] $prop2Val))
			{
				$rptObj.RetentionFlags = $prop2Val
			}
			$prop3Val = $null
			if($folder.TryGetProperty($RetentionPeriod,[ref] $prop3Val))
			{
				$rptObj.RetentionPeriod = $prop3Val
			}
			$rptCollection += $rptObj

		}
		$ReportFileName = "c:\temp\" + $MailboxName + "-TaggedFolders.csv"
		$rptCollection | Export-Csv -NoTypeInformation -Path $ReportFileName
		Write-Host ("Report written to " + $ReportFileName)
	}
}

function Report-UnTaggedFolders
{
    param( 
	    [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$false)] [string]$url
    )  
 	Begin
	{	
	    #GetRetentionTag
		$retentionTags = @{}
		Get-RetentionPolicyTag | ForEach-Object{
			$Cval = $_
			$retentionTagGUID = "{$($Cval.RetentionId.ToString())}"
			$policyTagGUID = new-Object Guid($retentionTagGUID)
			$retentionTags.Add($policyTagGUID.ToString().ToLower(),$Cval.Name)
		}
		$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)   
		$MsgRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		$Folders = Get-Folders -RootFolder $MsgRoot -service $service -IsTagged:$false
		$PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);  
		$rptCollection = @()
		foreach($folder in $Folders){
			$rptObj = "" | Select FolderName,FolderClass,Path,ItemCount,FolderItemSize
			$rptObj.FolderName = $folder.DisplayName
			$rptObj.FolderClass = $Folder.FolderClass
			$rptObj.Path = (Get-FolderPath -Folder $folder)
			$rptObj.ItemCount = $folder.TotalCount
			$folderSize = $null
			if($folder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED,[ref] $folderSize)){
				$rptObj.FolderItemSize = [Int64][System.Math]::Round(($folderSize/1024/1024),2)  
			}
			$rptCollection +=$rptObj
		}
		$ReportFileName = "c:\temp\" + $MailboxName + "-UnTaggedFolders.csv"
		$rptCollection | Export-Csv -NoTypeInformation -Path $ReportFileName
		Write-Host ("Report written to " + $ReportFileName)
	}
}

function Get-Folders{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.Folder]$RootFolder,
		[Parameter(Position=1, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
		[Parameter(Position=2, Mandatory=$false)] [Microsoft.Exchange.WebServices.Data.PropertySet]$PropertySet,
		[Parameter(Position=3, Mandatory=$true)] [bool]$IsTagged
		
    )  
 	Begin
	{
		if([string]::IsNullOrEmpty($PropertySet)){
			$PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
			#PR_POLICY_TAG 0x3019
        	$PolicyTag = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3019,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        	#PR_RETENTION_FLAGS 0x301D   
        	$RetentionFlags = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        	#PR_RETENTION_PERIOD 0x301A
        	$RetentionPeriod = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301A,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
			$PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);  
			$PropertySet.Add($PolicyTag)
			$PropertySet.Add($RetentionFlags)
			$PropertySet.Add($RetentionPeriod)
			$PropertySet.Add($PR_MESSAGE_SIZE_EXTENDED)
		}	
		$Folders = @()
		#Define Extended properties  
		$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
		#Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
		$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
		#Deep Transval will ensure all folders in the search path are returned  
		$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
		$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
		#Add Properties to the  Property Set  
		$PropertySet.Add($PR_Folder_Path);  
		$fvFolderView.PropertySet = $PropertySet;  
		#The Search filter will exclude any Search Folders
		
		$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Exists($PolicyTag) 
		
		$fiResult = $null  
		#The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
		do {  
			if($IsTagged)
			{
			   	$fiResult = $Service.FindFolders($RootFolder.Id,$sfSearchFilter,$fvFolderView)  
			}
			else
			{
				$sfNotSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($sfSearchFilter) 
				$fiResult = $Service.FindFolders($RootFolder.Id,$sfNotSearchFilter,$fvFolderView) 
			}
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
				$ffFolder | Add-Member -Name "FolderPath" -Value $fpath -MemberType NoteProperty
				$Folders += $ffFolder
		    } 
		    $fvFolderView.Offset += $fiResult.Folders.Count
		}while($fiResult.MoreAvailable -eq $true)  
		return $Folders	
	}
}
