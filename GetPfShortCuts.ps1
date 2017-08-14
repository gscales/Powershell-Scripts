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

function ConvertPFId {      
    param (  
        [Parameter(Position = 0, Mandatory = $true)] [String]$HexId,
        [Parameter(Position = 1, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$ExchangeService
			
    )  
    process {  
        $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternatePublicFolderId       
        $aiItem.FolderId = $HexId     
        $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId        
        $convertedId = $ExchangeService.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId)   
        return , $convertedId  
    }  
}  
function ConvertToString($ipInputString) {  
    $Val1Text = ""  
    for ($clInt = 0; $clInt -lt $ipInputString.length; $clInt++) {  
        $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt, 2), 16)))  
        $clInt++  
    }  
    return $Val1Text  
}  
function Get-PublicFolderShortCuts {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url

    )  
    Begin {
        $RptCollection = @()
        if ($url) {
            $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url 
        }
        else {
            $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials
        }
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
        }
        #PropDefs 
        $pidTagStoreEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(4091, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
        $PidTagNormalizedSubject = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E1D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String); 
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
        $PidTagWlinkAddressBookEID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6854, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
        $PidTagWlinkROGroupType = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6892, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
        $PidTagWlinkAddressBookStoreEID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6891, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
        $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
        #Get CommonViewsFolder
        $fppsPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
        $fppsPropset.add($PR_Folder_Path)
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $MailboxName)   
        $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)  
        $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, "Common Views") 
        $findFolderResults = $service.FindFolders($tfTargetFolder.Id, $SfSearchFilter, $fvFolderView) 
        if ($findFolderResults.TotalCount -eq 1) {
            $cvCommonViewsFolder = $findFolderResults.Folders[0]
            #Define ItemView to retrive just 1000 Items    
            #Find Items that are unread
            $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
            $psPropset.add($PidTagWlinkEntryId)
            $psPropset.add($PidTagWlinkFolderType)
            $pfScSearch = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PidTagWlinkFlags, 32771);
            $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)   
            $ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
            $ivItemView.PropertySet = $psPropset
            $fiItems = $service.FindItems($cvCommonViewsFolder.Id, $pfScSearch, $ivItemView)    
            foreach ($Item in $fiItems.Items) {
                $rptObj = "" | Select MailboxName, FolderName, FolderPath
                $rptObj.MailboxName = $MailboxName
                $rptObj.FolderName = $Item.Subject
                $idVal = $null
                Write-Host("Processing " + $Item.Subject)
                $EntryIdValue = $null
                if ($Item.TryGetProperty($PidTagWlinkEntryId, [ref]$EntryIdValue)) {
                    $Id = ([System.BitConverter]::ToString($EntryIdValue).Replace("-", ""))
                    $TransposedId = "000000001A447390AA6611CD9BC800AA002FC45A0300" + $Id.Substring(44)
                    $PfId = ConvertPFId -HexId $TransposedId -ExchangeService $service
                    $pfFolderId =	new-object Microsoft.Exchange.WebServices.Data.FolderId($PfId.FolderId)					 	
                    $PublicFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $pfFolderId, $fppsPropset)  
                    $foldpathval = $null  
                    #Try to get the FolderPath Value and then covert it to a usable String   
                    if ($PublicFolder.TryGetProperty($PR_Folder_Path, [ref] $foldpathval)) {  
                        $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)
                        $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
                        $hexString = $hexArr -join ''  
                        $hexString = $hexString.Replace("EFBFBE", "5C") 
                        $fpath = ConvertToString($hexString)   
        			    
							
                    }  
                    "FolderPath : " + $fpath  
                    $rptObj.FolderPath = $fpath
                    $RptCollection += $rptObj				
                }	            
				      
            }
        }
        else {
            Write-Error "Error getting Common Views Folder"        
        }
        $RptCollection | Export-Csv -NoTypeInformation -Path c:\temp\$MailboxName-PfShortCuts.csv
        Write-Host ("Report written to c:\temp\$MailboxName-PfShortCuts.csv")
    }
}

