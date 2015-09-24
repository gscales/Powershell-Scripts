function Connect-Exchange{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
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

if (-not ([System.Management.Automation.PSTypeName]'CalendarColor').Type)
{
Add-Type -TypeDefinition @"
   public enum CalendarColor
   {
   		Automatch = -1,
		Blue = 0,
		Green = 1,
		Peach = 2,
		Gray = 3,
		Teal = 4,
		Pink = 5,
		Olive = 6,
		Red = 7,
		Orange = 8,
		Purple = 9,
		Tan = 10,
		Light_Green = 12,
		Yellow = 13,
		Eton_Blue = 14
   }
"@
}

function Get-CommonViews{
	param (
			[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
			[Parameter(Position=1, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
		  )
	process{
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)     
		$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)    
		$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)   
		$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"Common Views")   
		$findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView)   
		if ($findFolderResults.TotalCount -gt 0){   
			return $findFolderResults.Folders[0]
		}
		else{
			return $null
		}
	}
}


function Exec-FindCalendarShortCut{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )  
 	Begin
	 {
		# Bind to the Calendar Folder
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
		$pidTagEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FFF, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684C, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkGroupName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6851, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
		$PidTagWlinkCalendarColor = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6853, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$psPropset.Add($pidTagEntryId)
		$psPropset.Add($PidTagWlinkGroupName)
		$psPropset.Add($PidTagWlinkCalendarColor)
		$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)
		$CalendarEntryId = $null
		[Void]$Calendar.TryGetProperty($pidTagEntryId,[ref]$CalendarEntryId)
		$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PidTagWlinkEntryId,[System.Convert]::ToBase64String($CalendarEntryId)) 
		#Get the HexId		
		$CommonViewFolder = Get-CommonViews -MailboxName $MailboxName -service $service
		$retrunItem = @()
		if($CommonViewFolder -ne $null){
			#Define ItemView to retrive just 1000 Items    
			$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)   
			$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
			$ivItemView.PropertySet = $psPropset
			$fiItems = $null    
			do{    
			    $fiItems = $service.FindItems($CommonViewFolder.Id,$SfSearchFilter,$ivItemView)    
			    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
			    foreach($Item in $fiItems.Items){  
					$retrunItem += $Item          
			    }    
			    $ivItemView.Offset += $fiItems.Items.Count    
			}while($fiItems.MoreAvailable -eq $true) 
		}
		return $retrunItem
	}
}
####################### 
<# 
.SYNOPSIS 
 Gets the Calendar WunderBar ShortCut for the Default Calendar Folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Gets the Calendar WunderBar ShortCut for the Default Calendar Folder using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 To create a Folder named test in the Root of the Mailbox
	 Get-DefaultCalendarFolderShortCut -Mailboxname mailbox@domain.com
	#> 
########################
function Get-DefaultCalendarFolderShortCut{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation
    )  
 	Begin
	 {
		$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$PidTagWlinkGroupName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6851, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
		$PidTagWlinkCalendarColor = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6853, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$ShortCutItems = Exec-FindCalendarShortCut -MailboxName $MailboxName -service $service
		Write-Output ("Number of ShortCuts Founds " + $ShortCutItems.Count)
		foreach($Item in $ShortCutItems){
			$GroupNameVal = $null
			$CalendarColor = $null
			[Void]$Item.TryGetProperty($PidTagWlinkCalendarColor,[ref]$CalendarColor)
			if($Item.TryGetProperty($PidTagWlinkGroupName,[ref]$GroupNameVal))
			{
				Write-Host ("Group Name " + $GroupNameVal + " Color " + [CalendarColor]$CalendarColor)
				Write-Output $Item
			}
		}
		
		

	 }
}
####################### 
<# 
.SYNOPSIS 
 Sets the Color of the Calendar WunderBar ShortCut for the Default Calendar Folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Sets the Color of the Calendar WunderBar ShortCut for the Default Calendar Folder using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 To create a Folder named test in the Root of the Mailbox
	 Set-DefaultCalendarFolderShortCut -Mailboxname mailbox@domain.com -CalendarColor Red

#> 
########################
function Set-DefaultCalendarFolderShortCut{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$true)] [CalendarColor]$CalendarColor
    )  
 	Begin
	 {
	 	$intVal = $CalendarColor -as [int]
		$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		$PidTagWlinkGroupName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6851, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
		$PidTagWlinkCalendarColor = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6853, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$ShortCutItems = Exec-FindCalendarShortCut -MailboxName $MailboxName -service $service
		Write-Output ("Number of ShortCuts Founds " + $ShortCutItems.Count)
		foreach($Item in $ShortCutItems){
			$GroupNameVal = $null
			$CalendarColorValue = $null
			[Void]$Item.TryGetProperty($PidTagWlinkCalendarColor,[ref]$CalendarColorValue)
			if($Item.TryGetProperty($PidTagWlinkGroupName,[ref]$GroupNameVal))
			{
				Write-Host ("Group Name " + $GroupNameVal + " Current Color " + [CalendarColor]$CalendarColorValue)
				$Item.SetExtendedProperty($PidTagWlinkCalendarColor,$intVal)
				$Item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
			}
		}
		
		

	 }
}
####################### 
<# 
.SYNOPSIS 
 Gets the Calendar WunderBar ShortCut for a Shared Calendar Folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Gets the Calendar WunderBar ShortCut for a Shared Calendar Folder  using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 To create a Folder named test in the Root of the Mailbox
	 Get-DefaultCalendarFolderShortCut -Mailboxname mailbox@domain.com -$SharedCalendarMailboxName sharedcalender@domain.com 
	#> 
########################
function Get-SharedCalendarFolderShortCut{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$true)] [string]$SharedCalendarMailboxName
    )  
 	Begin
	 {
		$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}	 
		$PidTagWlinkAddressBookEID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6854,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkFolderType = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684F, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkCalendarColor = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6853, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$PidTagWlinkGroupName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6851, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
	 	$CommonViewFolder = Get-CommonViews -MailboxName $MailboxName -service $service

		Write-Host ("Getting Autodiscover Settings Target")
		Write-Host ("Getting Autodiscover Settings Mailbox")
		$adset = GetAutoDiscoverSettings -adEmailAddress $MailboxName -Credentials $Credentials
		$storeID = ""
		if($adset -is [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverResponse]){
			Write-Host ("Get StoreId")
			$storeID = GetStoreId -AutoDiscoverSettings $adset
		}
		$adset = $null
		$abTargetABEntryId = ""
		$adset = GetAutoDiscoverSettings -adEmailAddress $SharedCalendarMailboxName -Credentials $Credentials
		if($adset -is [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverResponse]){
			Write-Host ("Get AB Id")
			$abTargetABEntryId = GetAddressBookId -AutoDiscoverSettings $adset
			$SharedUserDisplayName =  $adset.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDisplayName]
		}
		$ExistingShortCut = $false
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$psPropset.add($PidTagWlinkAddressBookEID)
		$psPropset.add($PidTagWlinkFolderType)
		$psPropset.add($PidTagWlinkCalendarColor)
		$psPropset.add($PidTagWlinkGroupName)
		$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)   
		$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
		$ivItemView.PropertySet = $psPropset
		$fiItems = $service.FindItems($CommonViewFolder.Id,$ivItemView)    
	    foreach($Item in $fiItems.Items){
			$aeidVal = $null
			$CalendarColorValue = $null
			$GroupNameVal = $null
			[Void]$Item.TryGetProperty($PidTagWlinkCalendarColor,[ref]$CalendarColorValue)
			[Void]$Item.TryGetProperty($PidTagWlinkGroupName,[ref]$GroupNameVal)
			if($Item.TryGetProperty($PidTagWlinkAddressBookEID,[ref]$aeidVal)){
					$fldType = $null
					if($Item.TryGetProperty($PidTagWlinkFolderType,[ref]$fldType)){
						if([System.BitConverter]::ToString($fldType).Replace("-","") -eq "0278060000000000C000000000000046"){
							if([System.BitConverter]::ToString($aeidVal).Replace("-","") -eq $abTargetABEntryId){
									Write-Host ("Group Name " + $GroupNameVal + " Current Color " + [CalendarColor]$CalendarColorValue)
									Write-Output $Item
							}
						}
					}
			}						      
		}
		
		

	 }
}
####################### 
<# 
.SYNOPSIS 
 Sets the Color of the Shared Calendar WunderBar ShortCut for the Shared Calendar Folder Shortcut in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Sets the Color of the Shared Calendar WunderBar ShortCut for the Shared Calendar Folder Shortcut in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
	Example 1 To create a Folder named test in the Root of the Mailbox
	 Set-DefaultCalendarFolderShortCut -Mailboxname mailbox@domain.com -$SharedCalendarMailboxName sharedcalender@domain.com -CalendarColor Red

#> 
########################
function Set-SharedCalendarFolderShortCut{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$true)] [string]$SharedCalendarMailboxName,
		[Parameter(Position=4, Mandatory=$true)] [CalendarColor]$CalendarColor
    )  
 	Begin
	 {
		$service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
		$intVal = $CalendarColor -as [int]
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}	 
		$PidTagWlinkAddressBookEID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6854,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkFolderType = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684F, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkCalendarColor = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6853, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$PidTagWlinkGroupName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6851, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
	 	$CommonViewFolder = Get-CommonViews -MailboxName $MailboxName -service $service

		Write-Host ("Getting Autodiscover Settings Target")
		Write-Host ("Getting Autodiscover Settings Mailbox")
		$adset = GetAutoDiscoverSettings -adEmailAddress $MailboxName -Credentials $Credentials
		$storeID = ""
		if($adset -is [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverResponse]){
			Write-Host ("Get StoreId")
			$storeID = GetStoreId -AutoDiscoverSettings $adset
		}
		$adset = $null
		$abTargetABEntryId = ""
		$adset = GetAutoDiscoverSettings -adEmailAddress $SharedCalendarMailboxName -Credentials $Credentials
		if($adset -is [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverResponse]){
			Write-Host ("Get AB Id")
			$abTargetABEntryId = GetAddressBookId -AutoDiscoverSettings $adset
			$SharedUserDisplayName =  $adset.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDisplayName]
		}
		$ExistingShortCut = $false
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$psPropset.add($PidTagWlinkAddressBookEID)
		$psPropset.add($PidTagWlinkFolderType)
		$psPropset.add($PidTagWlinkCalendarColor)
		$psPropset.add($PidTagWlinkGroupName)
		$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)   
		$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
		$ivItemView.PropertySet = $psPropset
		$fiItems = $service.FindItems($CommonViewFolder.Id,$ivItemView)    
	    foreach($Item in $fiItems.Items){
			$aeidVal = $null
			if($Item.TryGetProperty($PidTagWlinkAddressBookEID,[ref]$aeidVal)){
					$fldType = $null
					$CalendarColorValue = $null
					$GroupNameVal = $null
					[Void]$Item.TryGetProperty($PidTagWlinkCalendarColor,[ref]$CalendarColorValue)
					[Void]$Item.TryGetProperty($PidTagWlinkGroupName,[ref]$GroupNameVal)
					if($Item.TryGetProperty($PidTagWlinkFolderType,[ref]$fldType)){
						if([System.BitConverter]::ToString($fldType).Replace("-","") -eq "0278060000000000C000000000000046"){
							if([System.BitConverter]::ToString($aeidVal).Replace("-","") -eq $abTargetABEntryId){
								Write-Host ("Group Name " + $GroupNameVal + " Current Color " + [CalendarColor]$CalendarColorValue)
								$Item.SetExtendedProperty($PidTagWlinkCalendarColor,$intVal)
								$Item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
							}
						}
					}
			}						      
		}
		
		

	 }
}


function GetAutoDiscoverSettings{
	param (
	        $adEmailAddress = "$( throw 'emailaddress is a mandatory Parameter' )",
			[System.Management.Automation.PSCredential]$Credentials = "$( throw 'Credentials is a mandatory Parameter' )"
		  )
	process{
		$adService = New-Object Microsoft.Exchange.WebServices.AutoDiscover.AutodiscoverService($ExchangeVersion);
		$adService.Credentials = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString())  
		$adService.EnableScpLookup = $false;
		$adService.RedirectionUrlValidationCallback = {$true}
		$UserSettings = new-object Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[] 3
		$UserSettings[0] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDN
		$UserSettings[1] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::InternalRpcClientServer
		$UserSettings[2] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDisplayName
		$adResponse = $adService.GetUserSettings($adEmailAddress , $UserSettings);
		return $adResponse
	}
}
function GetAddressBookId{
	param (
	        $AutoDiscoverSettings = "$( throw 'AutoDiscoverSettings is a mandatory Parameter' )"
		  )
	process{
		$userdnString = $AutoDiscoverSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDN]
		$userdnHexChar = $userdnString.ToCharArray();
		foreach ($element in $userdnHexChar) {$userdnStringHex = $userdnStringHex + [System.String]::Format("{0:X}", [System.Convert]::ToUInt32($element))}
		$Provider = "00000000DCA740C8C042101AB4B908002B2FE1820100000000000000"
		$userdnStringHex = $Provider + $userdnStringHex + "00"
		return $userdnStringHex
	}
}
function GetStoreId{
	param (
	        $AutoDiscoverSettings = "$( throw 'AutoDiscoverSettings is a mandatory Parameter' )"
		  )
	process{
		$userdnString = $AutoDiscoverSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDN]
		$userdnHexChar = $userdnString.ToCharArray();
		foreach ($element in $userdnHexChar) {$userdnStringHex = $userdnStringHex + [System.String]::Format("{0:X}", [System.Convert]::ToUInt32($element))}	
		$serverNameString = $AutoDiscoverSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::InternalRpcClientServer]
		$serverNameHexChar = $serverNameString.ToCharArray();
		foreach ($element in $serverNameHexChar) {$serverNameStringHex = $serverNameStringHex + [System.String]::Format("{0:X}", [System.Convert]::ToUInt32($element))}
		$flags = "00000000"
		$ProviderUID = "38A1BB1005E5101AA1BB08002B2A56C2"
		$versionFlag = "0000"
		$DLLFileName = "454D534D44422E444C4C00000000"
		$WrappedFlags = "00000000"
		$WrappedProviderUID = "1B55FA20AA6611CD9BC800AA002FC45A"
		$WrappedType = "0C000000"
		$StoredIdStringHex = $flags + $ProviderUID + $versionFlag + $DLLFileName + $WrappedFlags + $WrappedProviderUID + $WrappedType + $serverNameStringHex + "00" + $userdnStringHex + "00"
		return $StoredIdStringHex
	}
}
function hex2binarray($hexString){
    $i = 0
    [byte[]]$binarray = @()
    while($i -le $hexString.length - 2){
        $strHexBit = ($hexString.substring($i,2))
        $binarray += [byte]([Convert]::ToInt32($strHexBit,16))
        $i = $i + 2
    }
    return ,$binarray
}
function ConvertId($EWSid){    
    $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId      
    $aiItem.Mailbox = $MailboxName      
    $aiItem.UniqueId = $EWSid   
    $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::EWSId;      
    return $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::StoreId)     
} 
function LoadProps()
{

}
function Create-SharedCalendarShortCut
{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$SourceMailbox,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
		[Parameter(Position=3, Mandatory=$true)] [String]$TargetMailbox,
		[Parameter(Position=4, Mandatory=$true)] [CalendarColor]$CalendarColor
    )  
 	Begin
	 {
	 	$intVal = $CalendarColor -as [int]
		$service = Connect-Exchange -MailboxName $SourceMailbox -Credentials $Credentials
		
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $SourceMailbox)
		}	 
		$pidTagStoreEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(4091, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagNormalizedSubject = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E1D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String); 
		$PidTagWlinkType = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6849, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$PidTagWlinkFlags = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$PidTagWlinkOrdinal = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684B, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkFolderType = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684F, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkSection = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6852, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$PidTagWlinkGroupHeaderID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6842, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkSaveStamp = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6847, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$PidTagWlinkGroupName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6851, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
		$PidTagWlinkStoreEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684E, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkGroupClsid = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6850, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684C, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkRecordKey = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkCalendarColor = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6853, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$PidTagWlinkAddressBookEID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6854,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$PidTagWlinkROGroupType = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6892,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$PidTagWlinkAddressBookStoreEID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6891,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
	 	$CommonViewFolder = Get-CommonViews -MailboxName $SourceMailbox -service $service
		#Get the TargetUsers Calendar
		# Bind to the Calendar Folder
		$fldPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$fldPropset.Add($pidTagStoreEntryId);
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetMailbox)
		}	 	
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$TargetMailbox)   
		$TargetCalendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$fldPropset)
		#Check for existing ShortCut for TargetMailbox
		#Get AddressBook Id for TargetUser	
		Write-Host ("Getting Autodiscover Settings Target")
		Write-Host ("Getting Autodiscover Settings Mailbox")
		$adset = GetAutoDiscoverSettings -adEmailAddress $SourceMailbox -Credentials $Credentials
		$storeID = ""
		if($adset -is [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverResponse]){
			Write-Host ("Get StoreId")
			$storeID = GetStoreId -AutoDiscoverSettings $adset
		}
		$adset = $null
		$abTargetABEntryId = ""
		$adset = GetAutoDiscoverSettings -adEmailAddress $TargetMailbox -Credentials $Credentials
		if($adset -is [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverResponse]){
			Write-Host ("Get AB Id")
			$abTargetABEntryId = GetAddressBookId -AutoDiscoverSettings $adset
			$SharedUserDisplayName =  $adset.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDisplayName]
		}
		$ExistingShortCut = $false
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$psPropset.add($PidTagWlinkAddressBookEID)
		$psPropset.add($PidTagWlinkFolderType)
		$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)   
		$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
		$ivItemView.PropertySet = $psPropset
		$fiItems = $service.FindItems($CommonViewFolder.Id,$ivItemView)    
	    foreach($Item in $fiItems.Items){
			$aeidVal = $null
			if($Item.TryGetProperty($PidTagWlinkAddressBookEID,[ref]$aeidVal)){
					$fldType = $null
					if($Item.TryGetProperty($PidTagWlinkFolderType,[ref]$fldType)){
						if([System.BitConverter]::ToString($fldType).Replace("-","") -eq "0278060000000000C000000000000046"){
							if([System.BitConverter]::ToString($aeidVal).Replace("-","") -eq $abTargetABEntryId){
								$ExistingShortCut = $true
								Write-Host "Found existing Shortcut"
								###$Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
							}
						}
					}
			}						      
		}
		if($ExistingShortCut -eq $false){
			$objWunderBarLink = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service  
			$objWunderBarLink.Subject = $SharedUserDisplayName  
			$objWunderBarLink.ItemClass = "IPM.Microsoft.WunderBar.Link"  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkAddressBookEID,(hex2binarray $abTargetABEntryId))  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkAddressBookStoreEID,(hex2binarray $storeID))  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkCalendarColor,$intVal)
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkFlags,0)
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkGroupName,"Shared Calendars")
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkFolderType,(hex2binarray "0278060000000000C000000000000046"))  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkGroupClsid,(hex2binarray "B9F0060000000000C000000000000046"))  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkROGroupType,-1)
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkSection,3)  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkType,2)  
			$objWunderBarLink.IsAssociated = $true
			$objWunderBarLink.Save($CommonViewFolder.Id)
			Write-Host ("ShortCut Created for - " + $SharedUserDisplayName)
		}
		
	 }
}