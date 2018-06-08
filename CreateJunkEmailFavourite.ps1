function Connect-Exchange { 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url
    )  
    Begin {
        Load-EWSManagedAPI
		
        ## Set Exchange Version  
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
		  
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
             Import-Module  $EWSDLL
        }
        else {
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

function GetAutoDiscoverSettings{
	param (
	        $adEmailAddress = "$( throw 'emailaddress is a mandatory Parameter' )",
			$Credentials = "$( throw 'Credentials is a mandatory Parameter' )"
		  )
	process{
		$adService = New-Object Microsoft.Exchange.WebServices.AutoDiscover.AutodiscoverService($ExchangeVersion);
		$adService.Credentials = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString()) 
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


function New-JunkEmailFavourite {
    param( 
       [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url
       
    )  
    Begin {
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)        
        }
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
        #PropDefs 
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
        $pidTagEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FFF, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
        $pidTagWlinkRecKey = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
        #Get the TargetUsers Calendar
        # Bind to the Calendar Folder
        $fldPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
        $fldPropset.Add($pidTagStoreEntryId);
        $fldPropset.Add($pidTagEntryId)
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail,$MailboxName)   
        $JunkEmailFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$fldPropset)
        $pidTagEntryIdValue = $null
        [void]$JunkEmailFolder.TryGetProperty($pidTagEntryId,[ref]$pidTagEntryIdValue)
        #Check for existing ShortCut for TargetMailbox
        #Get AddressBook Id for TargetUser
        $JunkEmailFolderEntryId = [System.BitConverter]::ToString($pidTagEntryIdValue).Replace("-","")
         Write-Host ("Getting Autodiscover Settings Mailbox")
        $adset = GetAutoDiscoverSettings -adEmailAddress $MailboxName -Credentials $Credentials
        $storeID = ""
        if($adset -is [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverResponse]){
            Write-Host ("Get StoreId")
            $storeID = GetStoreId -AutoDiscoverSettings $adset
        }
        $adset = $null
        $abTargetABEntryId = ""
        $adset = GetAutoDiscoverSettings -adEmailAddress $MailboxName -Credentials $Credentials
        if($adset -is [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverResponse]){
            Write-Host ("Get AB Id")
            $abTargetABEntryId = GetAddressBookId -AutoDiscoverSettings $adset
            $SharedUserDisplayName =  $adset.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDisplayName]
        }
        Write-Host ("Getting CommonVeiwFolder")
        #Get CommonViewFolder
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
        $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
        $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"Common Views") 
        $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView) 
        if ($findFolderResults.TotalCount -gt 0){ 
            $ExistingShortCut = $false
            $cvCommonViewsFolder = $findFolderResults.Folders[0]
            #Define ItemView to retrive just 1000 Items    
            #Find Items that are unread
            $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
            $psPropset.add($PidTagWlinkAddressBookEID)
            $psPropset.add($PidTagWlinkFolderType)
            $psPropset.add($PidTagWlinkEntryId)
            $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)   
            $ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
            $ivItemView.PropertySet = $psPropset
            $fiItems = $service.FindItems($cvCommonViewsFolder.Id,$ivItemView)    
            foreach($Item in $fiItems.Items){
                $aeidVal = $null
                if($Item.TryGetProperty($PidTagWlinkEntryId,[ref]$aeidVal)){
                    if([System.BitConverter]::ToString($aeidVal).Replace("-","") -eq $JunkEmailFolderEntryId){
                                $ExistingShortCut = $true
                                Write-Host "Found existing Shortcut"
                    }
                }						      
            }
            if($ExistingShortCut -eq $false){
                If($storeID.length -gt 5){
                    $objWunderBarLink = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service  
                    $objWunderBarLink.Subject =  $JunkEmailFolder.DisplayName
                    $objWunderBarLink.ItemClass = "IPM.Microsoft.WunderBar.Link"  
                    $objWunderBarLink.SetExtendedProperty($PidTagWlinkEntryId,(hex2binarray $JunkEmailFolderEntryId))  
                    $objWunderBarLink.SetExtendedProperty($pidTagWlinkRecKey,(hex2binarray $JunkEmailFolderEntryId))  
                    $objWunderBarLink.SetExtendedProperty($PidTagWlinkFlags,0)
                    $objWunderBarLink.SetExtendedProperty($PidTagWlinkFolderType,(hex2binarray "0078060000000000C000000000000046"))  
                    $objWunderBarLink.SetExtendedProperty($PidTagWlinkOrdinal,(hex2binarray "FEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFEFE7F"))  
                    $objWunderBarLink.SetExtendedProperty($PidTagWlinkROGroupType,-1)
                    $objWunderBarLink.SetExtendedProperty($PidTagWlinkSection,1)  
                    $objWunderBarLink.SetExtendedProperty($PidTagWlinkStoreEntryId,(hex2binarray $storeID))
                    $objWunderBarLink.SetExtendedProperty($PidTagWlinkType,0)  
                    $objWunderBarLink.IsAssociated = $true
                    $objWunderBarLink.Save($findFolderResults.Folders[0].Id)
                    Write-Host ("ShortCut Created for - " + $SharedUserDisplayName)
                }
                else{
                    Write-Host ("Error with Id's")
                }
            }
        }

		
    }
}
