function Connect-Exchange { 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,        
        [Parameter(Position = 1, Mandatory = $false)] [string]$url,
        [Parameter(Position = 2, Mandatory = $false)] [string]$ClientId,
        [Parameter(Position = 3, Mandatory = $false)] [string]$redirectUrl,
        [Parameter(Position = 4, Mandatory = $false)] [string]$AccessToken,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$basicAuth,
        [Parameter(Position = 6, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials

    )  
    Begin {
        Invoke-LoadEWSManagedAPI
		
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1      
		
        ## Create Exchange Service Object  
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
		  
        ## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
        if ($basicAuth.IsPresent) {
            $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString())  
            $service.Credentials = $creds    
        }
        else {
            if ([String]::IsNullOrEmpty($AccessToken)) {
                $Resource = "Outlook.Office365.com"    
                if ([String]::IsNullOrEmpty($ClientId)) {
                    $ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
                }
                if ([String]::IsNullOrEmpty($redirectUrl)) {
                    $redirectUrl = "urn:ietf:wg:oauth:2.0:oob"  
                }
                $Script:Token = Get-EWSAccessToken -MailboxName $MailboxName -ClientId $ClientId -redirectUrl $redirectUrl  -ResourceURL $Resource -Prompt $Prompt -CacheCredentials   
                $OAuthCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials((ConvertFrom-SecureStringCustom -SecureToken $Script:Token.access_token))
            }
            else {
                $OAuthCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($AccessToken)
            }    
            $service.Credentials = $OAuthCredentials
        }  
        #Credentials Option 1 using UPN for the windows Account  
        #$psCred = Get-Credential  
  
        #Credentials Option 2  
        #service.UseDefaultCredentials = $true  
        #$service.TraceEnabled = $true
        ## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
		  
        Invoke-HandleSSL	
		  
        ## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  

        #CAS URL Option 1 Autodiscover  
        if ($url) {
            $uri = [system.URI] $url
            $service.Url = $uri    
        }
        else {
            $service.AutodiscoverUrl($MailboxName, { $true })  
        }
        #Write-host ("Using CAS Server : " + $Service.url)   
		   
        #CAS URL Option 2 Hardcoded  
		  
        #$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
        #$service.Url = $uri    
		  
        ## Optional section for Exchange Impersonation  
		  
        #$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
        if (!$service.Url) {
            throw "Error connecting to EWS"
        }
        else {	
            $Name = $service.ResolveName($MailboxName)	
            Write-Verbose("Exchange Version " + $service.ServerInfo.MajorVersion)
            if ($service.ServerInfo.MajorVersion -ge 15) {
                $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1	
                if ($service.ServerInfo.MinorVersion -ge 20) {                    
                    if ([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016) {
                        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
                    }
                    else {
                        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
                    }   
                }
                $UpdatedService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
                $UpdatedService.Credentials = $service.Credentials
                $UpdatedService.Url = $service.Url;
                $service = $UpdatedService
                write-verbose("Update Version")
            }
            return, $service
        }
    }
}


function Invoke-LoadEWSManagedAPI {
    param( 
    )  
    Begin {
        if (Test-Path ($script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll")) {
            Import-Module ($script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll")
            $Script:EWSDLL = $script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll"
            write-verbose ("Using EWS dll from Local Directory")
        }
        else {

        
            ## Load Managed API dll  
            ###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
            $EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
            if (Test-Path $EWSDLL) {
                Import-Module $EWSDLL
                $Script:EWSDLL = $EWSDLL 
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
}

function Invoke-HandleSSL {
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

        ## end code from http://poshcode.org/624 ##

    }
}

function ConvertTo-String($ipInputString) {    
    $Val1Text = ""    
    for ($clInt = 0; $clInt -lt $ipInputString.length; $clInt++) {    
        $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt, 2), 16)))    
        $clInt++    
    }    
    return $Val1Text    
}    

function TraceHandler() {
    $sourceCode = @"
    public class ewsTraceListener : Microsoft.Exchange.WebServices.Data.ITraceListener
    {
        public System.String LogFile {get;set;}
        public void Trace(System.String traceType, System.String traceMessage)
        {
            System.IO.File.AppendAllText(this.LogFile, traceMessage);
        }
    }
"@    

    Add-Type -TypeDefinition $sourceCode -Language CSharp -ReferencedAssemblies $Script:EWSDLL
    $TraceListener = New-Object ewsTraceListener
    return $TraceListener


}

function Invoke-MailboxChangeDiscovery {
    param( 
        [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$disableImpersonation,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$basicAuth,
        [Parameter(Position = 5, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 6, Mandatory = $true)] [Int32]$secondstolookback,
        [Parameter(Position = 7, Mandatory = $false)] [switch]$getJsonMetaData       
    )  
    Process {
        $Script:ItemCount = 0    
        $startDate = (Get-Date).AddSeconds(-$secondstolookback)
        $Script:RptCollection = @()
        if (!$basicAuth.IsPresent) {
            $service = Connect-Exchange -MailboxName $MailboxName -url $url
        }
        else {
            if (!$Credentials) {
                $Credentials = Get-Credential
            }
            $service = Connect-Exchange -MailboxName $MailboxName -url $url -basicAuth:$basicAuth.IsPresent -Credentials $Credentials
        } 
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);
        if ($disableImpersonation.IsPresent) {
            Write-Verbose ("Impersonation is disabled")    
        }
        else {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
        }
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $MailboxName)         
        $RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)  
        $FolderIndex = Get-SubFolders -ParentFolder $RootFolder -startdatetime $startDate
        foreach ($folder in $FolderIndex.Values) {
            if ($folder.FolderType -eq 1) {
                if ($folder.FolderPath -ne "\Recoverable Items\Audits") {
                    Write-Verbose ("Processing Folder " + $folder.FolderPath) 
                    Invoke-ScanFolderItems -Folder $folder -startdatetime $startDate  -getJsonMetaData:$getJsonMetaData.IsPresent 
                    Invoke-ScanFolderItems -Folder $folder -startdatetime $startDate -FAIItems -getJsonMetaData:$getJsonMetaData.IsPresent
                }
                else {
                    Write-Verbose ("****Skipping Audits folder " + $folder.FolderPath)
                }

            }
            else {
                Write-Verbose ("****Skipping SearchFolder " + $folder.FolderPath)
            }
         
        }
        $ExportFile = $script:ModuleRoot + "\ChangeDiscovery-" + (Get-Date).ToString("yyyyMMddHHmmss") + ".csv" 
        $Script:RptCollection | export-csv -Path $ExportFile -NoTypeInformation
        Write-Host ("Report saved to " + $ExportFile)
       
    }
}function Invoke-ScanFolderItems { 
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
        [Parameter(Position = 1, Mandatory = $false)] [String]$MailboxName,
        [Parameter(Position = 2, Mandatory = $false)] [DateTime]$startdatetime,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$FAIItems,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$getJsonMetaData   
    )  
    Begin {    
        $Script:ItemCount = 0
        $Sfgt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, $startdatetime)
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000) 
        $ItemPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        if ($FAIItems.IsPresent) {
            $ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
        }
        $PR_ENTRYID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FFF, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary) 
        if($getJsonMetaData.IsPresent){
            $RawJSON = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Guid]::Parse("2842957E-8ED9-439B-99B5-F681924BD972"),"RawJSON", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String) 
            $Data = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Guid]::Parse("3C896DED-22C5-450F-91F6-3D1EF0848F6E"),"Data", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String) 
            $ItemPropset.Add($RawJSON)
            $ItemPropset.Add($Data)
        }
        $ItemPropset.Add($PR_ENTRYID)
        $ivItemView.PropertySet = $ItemPropset
        $fiItems = $null    
        do {
            try {
                $error.Clear()
                $fiItems = $Folder.Service.FindItems($Folder.Id, $Sfgt, $ivItemView)   
            }
            catch {
                Write-Verbose ("Error " + $PSItem.Exception.InnerException.Message)
                Invoke-ProcessEWSError -Item $PSItem
                $error.Clear()
                $fiItems = $Folder.Service.FindItems($Folder.Id, $Sfgt, $ivItemView)
                if ($error.Count -ne 0) {
                    $Script:ErrorCount++                    
                }       
            }
            if($fiItems.Items.Count -gt 0){       
                write-verbose ("LoadPropertiesForItems")         
                $Items = Invoke-LoadPropertiesForItems -ItemList $fiItems.Items -service $Folder.Service -DetailPropSet $ItemPropset
                foreach ($Item in $Items) {
                    $Script:ItemCount++
                    $rptObj = "" | Select FolderPath, Type, Collection, Subject, ItemClass, EntryId, DateTimeCreated, LastModifiedTime,RawJSON,DataProp
                    $rptObj.FolderPath = $Folder.FolderPath                
                    $rptObj.Subject = $Item.Subject
                    if ($FAIItems.IsPresent) {
                        $rptObj.Collection = "FAIItem"
                    }
                    else {
                        $rptObj.Collection = "Messages"
                    }
                    $rptObj.ItemClass = $Item.ItemClass
                    $EntryIdValue = $null
                    if ($Item.TryGetProperty($PR_ENTRYID, [ref]$EntryIdValue)) {
                        $rptObj.EntryId = [System.BitConverter]::ToString($EntryIdValue).Replace("-", "")
                    }                
                    $rptObj.DateTimeCreated = $Item.DateTimeCreated
                    $rptObj.LastModifiedTime = $Item.LastModifiedTime
                    Write-Verbose ("Processing Item : " + $Item.Subject)
                    if ($Item.DateTimeCreated -gt $startdatetime) {
                        Write-Verbose ("New Item Found")
                        $rptObj.Type = "New"
                    }
                    else {
                        Write-Verbose ("Modified Item Found")
                        $rptObj.Type = "Modified"
                    }
                    $RawJSONValue = $null
                    if($Item.TryGetProperty($RawJSON,[ref]$RawJSONValue)){
                        $rptObj.RawJSON = $RawJSONValue
                    }
                    $RawDataValue = $null
                    if($Item.TryGetProperty($Data,[ref]$RawDataValue)){
                        $rptObj.DataProp = $RawDataValue 
                    }
                    $Script:RptCollection += $rptObj
                }
            }
            $ivItemView.Offset += $fiItems.Items.Count   
            Write-Verbose ("Processed " + $Script:ItemCount) 
        }while ($fiItems.MoreAvailable -eq $true) 

    }
}


function Get-SubFolders { 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.Folder]$ParentFolder,
        [Parameter(Position = 1, Mandatory = $false)] [Microsoft.Exchange.WebServices.Data.PropertySet]$PropertySet,
        [Parameter(Position = 2, Mandatory = $false)] [DateTime]$startdatetime

        
    )  
    Begin {
        $Folders = New-Object system.collections.hashtable
        $PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
        if ([string]::IsNullOrEmpty($PropertySet)) {
            $PR_ENTRYID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FFF, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary) 
            $PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
            #PR_POLICY_TAG 0x3019
            $PR_POLICY_TAG = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3019, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
            #PR_RETENTION_FLAGS 0x301D   
            $RetentionFlags = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
            #PR_RETENTION_PERIOD 0x301A
            $RetentionPeriod = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
            $SourceKey = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65E0, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
            $PidTagMessageSizeExtended = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0xe08, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
            $PR_ATTACH_ON_NORMAL_MSG_COUNT = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x66B1, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
            $PR_ATTR_HIDDEN = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x10F4, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean);
            $PidTagLastModificationTime = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3008, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime);
            $PR_CREATION_TIME = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3007, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime);
            $PropertySet.Add($PR_POLICY_TAG)
            $PropertySet.Add($RetentionFlags)
            $PropertySet.Add($RetentionPeriod)
            $PropertySet.Add($SourceKey)
            $PropertySet.Add($PidTagMessageSizeExtended)
            $PropertySet.Add($PR_ATTACH_ON_NORMAL_MSG_COUNT)
            $PropertySet.Add($PR_ENTRYID)
            $PropertySet.Add($PR_FOLDER_TYPE)
            $PropertySet.Add($PR_ATTR_HIDDEN)
            $PropertySet.Add($PidTagLastModificationTime)
            $PropertySet.Add($PR_CREATION_TIME)
        }	
		
        #Define Extended properties  

        #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
        $fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
        #Deep Transval will ensure all folders in the search path are returned  
        $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
        $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
        #Add Properties to the  Property Set  
        $PropertySet.Add($PR_Folder_Path);  
        $fvFolderView.PropertySet = $PropertySet;  
        #The Search filter will exclude any Search Folders  
        # $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE, "1")  
        $fiResult = $null  
        #The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
        do { 
            try {
                $error.Clear()
                $fiResult = $ParentFolder.FindFolders($fvFolderView)
            }
            catch {
                Write-Verbose ("Error " + $PSItem.Exception.InnerException.Message)
                Invoke-ProcessEWSError  -Item $PSItem
                $fiResult = $ParentFolder.FindFolders($fvFolderView)
            }              
            foreach ($ffFolder in $fiResult.Folders) {  
                $foldpathval = $null  
                #Try to get the FolderPath Value and then covert it to a usable String   
                if ($ffFolder.TryGetProperty($PR_Folder_Path, [ref] $foldpathval)) {  
                    $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
                    $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
                    $hexString = $hexArr -join ''  
                    $hexString = $hexString.Replace("FEFF", "5C00")  
                    $fpath = ConvertToString($hexString)  
                }  
                if ($Archive.IsPresent) {
                    $ffFolder | Add-Member -Name "FolderPath" -Value ("\InPlace-Archive" + $fpath) -MemberType NoteProperty
                }
                else {
                    $ffFolder | Add-Member -Name "FolderPath" -Value $fpath -MemberType NoteProperty
                }
                $ffFolder | Add-Member -Name "Mailbox" -Value $ParentFolder.Mailbox -MemberType NoteProperty
                $prop1Val = $null    
                if ($ffFolder.TryGetProperty($PR_POLICY_TAG, [ref] $prop1Val)) {
                    $rtnStringVal =	[System.BitConverter]::ToString($prop1Val).Replace("-", "");
                    $rtnStringVal = $rtnStringVal.Substring(6, 2) + $rtnStringVal.Substring(4, 2) + $rtnStringVal.Substring(2, 2) + $rtnStringVal.Substring(0, 2) + "-" + $rtnStringVal.Substring(10, 2) + $rtnStringVal.Substring(8, 2) + "-" + $rtnStringVal.Substring(14, 2) + $rtnStringVal.Substring(12, 2) + "-" + $rtnStringVal.Substring(16, 2) + $rtnStringVal.Substring(18, 2) + "-" + $rtnStringVal.Substring(20, 12)           
                    Add-Member -InputObject $ffFolder -MemberType NoteProperty -Name PR_POLICY_TAG -Value $rtnStringVal               
                    
        
                }
                $prop4Val = $null
                if ($ffFolder.TryGetProperty($PidTagMessageSizeExtended, [ref]  $prop4Val)) {
                    Add-Member -InputObject $ffFolder -MemberType NoteProperty -Name FolderSize -Value $prop4Val
                }
                $prop5Val = $null
                if ($ffFolder.TryGetProperty($PR_ATTACH_ON_NORMAL_MSG_COUNT, [ref]  $prop5Val)) {
                    Add-Member -InputObject $ffFolder -MemberType NoteProperty -Name PR_ATTACH_ON_NORMAL_MSG_COUNT -Value $prop5Val
                }
                $prop6Val = $null
                if ($ffFolder.TryGetProperty($PR_ENTRYID, [ref]  $prop6Val)) {
                    $HexVal = [System.BitConverter]::ToString($prop6Val)
                    Add-Member -InputObject $ffFolder -MemberType NoteProperty -Name PR_ENTRYID -Value $HexVal.Replace("-", "")
                }
                $prop7Val = $null
                if ($ffFolder.TryGetProperty($PidTagLastModificationTime, [ref]$prop7Val)) {                    
                    Add-Member -InputObject $ffFolder -MemberType NoteProperty -Name LastModificationTime -Value $prop7Val
                }
                $prop8Val = $null
                if ($ffFolder.TryGetProperty($PR_CREATION_TIME, [ref]$prop8Val)) {                    
                    Add-Member -InputObject $ffFolder -MemberType NoteProperty -Name CreationTime -Value $prop8Val
                }
                $prop9Val = $null
                if ($ffFolder.TryGetProperty($PR_FOLDER_TYPE, [ref]$prop9Val)) {                    
                    Add-Member -InputObject $ffFolder -MemberType NoteProperty -Name FolderType -Value $prop9Val
                }
                $rptObj = "" | Select FolderPath, Type, Collection, Subject, ItemClass, EntryId, DateTimeCreated, LastModifiedTime,RawJSON,DataProp
                $rptObj.FolderPath = $ffFolder.FolderPath   
                $rptObj.Subject = $ffFolder.DisplayName
                if ($ffFolder.FolderType -eq 2) {
                    $rptObj.Collection = "SearchFolder"
                }      
  
                if ($ffFolder.CreationTime -gt $startdatetime.ToUniversalTime()) {
                    Write-Verbose ("New Folder Found")
                    $rptObj.Type = "New"
                    $Script:RptCollection += $rptObj

                }
                else {
                    if ($ffFolder.CreationTime -gt $startdatetime.ToUniversalTime()) {
                        Write-Verbose ("Modified Folder Found")
                        $rptObj.Type = "Modified"
                        $Script:RptCollection += $rptObj
                    }

                }
                $Folders.Add($ffFolder.Id.UniqueId, $ffFolder)              
            } 
            $fvFolderView.Offset += $fiResult.Folders.Count
        }while ($fiResult.MoreAvailable -eq $true)  
        return, $Folders	
    }
}




function Invoke-LoadPropertiesForItems() {
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [PSObject]$ItemList,
        [Parameter(Position = 1, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
        [Parameter(Position = 2, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.PropertySet]$DetailPropSet
    ) Begin {
        try {
            $error.Clear()
            [Void]$service.LoadPropertiesForItems($ItemList, $DetailPropSet)     
        }
        catch {
            Write-Verbose ("Error " + $PSItem.Exception.InnerException.Message)
            Invoke-ProcessEWSError -Item $PSItem
            $error.clear()
            [Void]$service.LoadPropertiesForItems($ItemList, $DetailPropSet)
            if ($error.Count -ne 0) {
                $Script:ErrorCount ++
            }     
        }
        return ,$ItemList       
    }
}



function Invoke-ProcessEWSError {
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [PSObject]$Item
    )
    Process {
        Write-Verbose $Item.Exception.InnerException.ErrorCode.ToString()
        if ($Item.Exception.InnerException -is [Microsoft.Exchange.WebServices.Data.ServiceResponseException]) {
            switch ($Item.Exception.InnerException.ErrorCode.ToString()) {
                "ErrorServerBusy" {
                    $Seconds = [Math]::Round(($Item.Exception.InnerException.BackOffMilliseconds / 1000), 0)
                    Write-Verbose ("Resume in " + $Seconds + " Seconds")
                    Start-Sleep -Milliseconds $Item.Exception.InnerException.BackOffMilliseconds
                }
                "ErrorAccessDenied" {
                    Write-Verbose("Access Denied Error")
                }
                Default {
                    Write-Verbose ("Resume in 60 Seconds")
                    Start-Sleep -Seconds 60 
                }
            }             
                
        }
    }
}

function ConvertId {    
    param (
        [Parameter(Position = 1, Mandatory = $false)] [String]$HexId,
        [Parameter(Position = 2, Mandatory = $false)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )
    process {
        $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId      
        $aiItem.Mailbox = $MailboxName      
        $aiItem.UniqueId = $HexId   
        $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId      
        $convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId) 
        return $convertedId.UniqueId
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



function Invoke-ValidateToken {
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    
    )
    begin {
        $MailboxName = $Script:Token.mailbox
        $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
        $expiry = $minTime.AddSeconds($Script:Token.expires_on)
        if ($expiry -le [DateTime]::Now.ToUniversalTime().AddMinutes(10)) {
            if ([bool]($Script:Token.PSobject.Properties.name -match "refresh_token")) {
                write-host "Refresh Token"
                $Script:Token = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $Script:Token
                $OAuthCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials((ConvertFrom-SecureStringCustom -SecureToken $Script:Token.access_token))
                $service.Credentials = $OAuthCredentials
            }
            else {
                throw "App Token has expired"
            }        
        }
    }
}


function Get-EWSAccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $MailboxName,
    
        [Parameter(Position = 1, Mandatory = $false)]
        [string]
        $ClientId,
    
        [Parameter(Position = 2, Mandatory = $false)]
        [string]
        $redirectUrl,
    
        [Parameter(Position = 3, Mandatory = $false)]
        [string]
        $ClientSecret,
    
        [Parameter(Position = 4, Mandatory = $false)]
        [string]
        $ResourceURL,
    
        [Parameter(Position = 5, Mandatory = $false)]
        [switch]
        $Beta,
    
        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $Prompt,

        [Parameter(Position = 7, Mandatory = $false)]
        [switch]
        $CacheCredentials
    
    )
    Begin {
        Add-Type -AssemblyName System.Web
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        if ([String]::IsNullOrEmpty($ClientId)) {
            $ReturnToken = Get-ProfiledToken -MailboxName $MailboxName
            if ($ReturnToken -eq $null) {
                Write-Error ("No Access Token for " + $MailboxName)
            }
            else {				
                return $ReturnToken
            }
        }
        else {
            if ([String]::IsNullOrEmpty(($ClientSecret))) {
                $ClientSecret = $AppSetting.ClientSecret
            }
            if ([String]::IsNullOrEmpty($redirectUrl)) {
                $redirectUrl = [System.Web.HttpUtility]::UrlEncode("urn:ietf:wg:oauth:2.0:oob")
            }
            else {
                $redirectUrl = [System.Web.HttpUtility]::UrlEncode($redirectUrl)
            }
            if ([String]::IsNullOrEmpty($ResourceURL)) {
                $ResourceURL = $AppSetting.ResourceURL
            }
            if ([String]::IsNullOrEmpty($Prompt)) {
                $Prompt = "refresh_session"
            }
        
            $code = Show-OAuthWindow -Url "https://login.microsoftonline.com/common/oauth2/authorize?resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&response_type=code&redirect_uri=$redirectUrl&prompt=$Prompt&domain_hint=organizations&response_mode=form_post"
            $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
            if (![String]::IsNullOrEmpty($ClientSecret)) {
                $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_secret=$ClientSecret&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
            }
            $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
            $ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"), $content)
            $JsonObject = ConvertFrom-Json -InputObject $ClientReesult.Result.Content.ReadAsStringAsync().Result
            if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token")) {
                $JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
            }
            if ([bool]($JsonObject.PSobject.Properties.name -match "access_token")) {
                $JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
            }
            if ([bool]($JsonObject.PSobject.Properties.name -match "id_token")) {
                $JsonObject.id_token = (Get-ProtectedToken -PlainToken $JsonObject.id_token)
            }
            Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
            Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
            Add-Member -InputObject $JsonObject -NotePropertyName mailbox -NotePropertyValue $MailboxName
            if ($Beta.IsPresent) {
                Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue $True
            }
            return $JsonObject
        }
    }
}

function Show-OAuthWindow {
    [CmdletBinding()]
    param (
        [System.Uri]
        $Url
    
    )
    ## Start Code Attribution
    ## Show-AuthWindow function is the work of the following Authors and should remain with the function if copied into other scripts
    ## https://foxdeploy.com/2015/11/02/using-powershell-and-oauth/
    ## https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/
    ## End Code Attribution
    Add-Type -AssemblyName System.Web
    Add-Type -AssemblyName System.Windows.Forms

    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{ Width = 440; Height = 640 }
    $web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{ Width = 420; Height = 600; Url = ($url) }
    $Navigated = {
   
        $formElements = $web.Document.GetElementsByTagName("input");  
        foreach ($element in $formElements) {
            if ($element.Name -eq "code") {
                $Script:oAuthCode = $element.GetAttribute("value");
                Write-Verbose $Script:oAuthCode
                $form.Close();
            }
        }   
    }
    
    $web.ScriptErrorsSuppressed = $true
    $web.Add_Navigated($Navigated)
    $form.Controls.Add($web)
    $form.Add_Shown( { $form.Activate() })
    $form.ShowDialog() | Out-Null
    return $Script:oAuthCode
}

function Get-ProtectedToken {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [String]
        $PlainToken
    )
    begin {
        $SecureEncryptedToken = Protect-String -String $PlainToken
        return, $SecureEncryptedToken
    }
}
function Get-HTTPClient {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $MailboxName
    )
    process {
        Add-Type -AssemblyName System.Net.Http
        $handler = New-Object  System.Net.Http.HttpClientHandler
        $handler.CookieContainer = New-Object System.Net.CookieContainer
        $handler.AllowAutoRedirect = $true;
        $HttpClient = New-Object System.Net.Http.HttpClient($handler);
        #$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "");
        $Header = New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json")
        $HttpClient.DefaultRequestHeaders.Accept.Add($Header);
        $HttpClient.Timeout = New-Object System.TimeSpan(0, 0, 90);
        $HttpClient.DefaultRequestHeaders.TransferEncodingChunked = $false
        if (!$HttpClient.DefaultRequestHeaders.Contains("X-AnchorMailbox")) {
            $HttpClient.DefaultRequestHeaders.Add("X-AnchorMailbox", $MailboxName);
        }
        $Header = New-Object System.Net.Http.Headers.ProductInfoHeaderValue("RestClient", "1.1")
        $HttpClient.DefaultRequestHeaders.UserAgent.Add($Header);
        return $HttpClient
    }
}

function Protect-String {
    <#
.SYNOPSIS
    Uses DPAPI to encrypt strings.

.DESCRIPTION
    Uses DPAPI to encrypt strings.

.PARAMETER String
    The string to encrypt.

.EXAMPLE
    PS C:\> Protect-String -String $secret

    Encrypts the content stored in $secret and returns it.
#>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", "")]
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline = $true)]
        [string[]]
        $String
    )

    begin {
        Add-Type -AssemblyName System.Security -ErrorAction Stop
    }
    process {
        foreach ($item in $String) {
            $stringBytes = [Text.Encoding]::UTF8.GetBytes($item)
            $encodedBytes = [System.Security.Cryptography.ProtectedData]::Protect($stringBytes, $null, 'CurrentUser')
            [System.Convert]::ToBase64String($encodedBytes) | ConvertTo-SecureString -AsPlainText -Force
        }
    }
}

function Invoke-RefreshAccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $MailboxName,
    
        [Parameter(Position = 1, Mandatory = $true)]
        [psobject]
        $AccessToken
    )
    process {
        Add-Type -AssemblyName System.Web
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $ClientId = $AccessToken.clientid
        # $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AccessToken.redirectUrl)
        $redirectUrl = $AccessToken.redirectUrl
        $RefreshToken = (ConvertFrom-SecureStringCustom -SecureToken $AccessToken.refresh_token)
        $AuthorizationPostRequest = "client_id=$ClientId&refresh_token=$RefreshToken&grant_type=refresh_token&redirect_uri=$redirectUrl"
        $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
        $ClientResult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"), $content)
        if (!$ClientResult.Result.IsSuccessStatusCode) {
            Write-Output ("Error making REST POST " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
            Write-Output $ClientResult.Result
            if ($ClientResult.Content -ne $null) {
                Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);
            }
        }
        else {
            $JsonObject = ConvertFrom-Json -InputObject $ClientResult.Result.Content.ReadAsStringAsync().Result
            if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token")) {
                $JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
            }
            if ([bool]($JsonObject.PSobject.Properties.name -match "access_token")) {
                $JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
            }
            if ([bool]($JsonObject.PSobject.Properties.name -match "id_token")) {
                $JsonObject.id_token = (Get-ProtectedToken -PlainToken $JsonObject.id_token)
            }
            Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
            Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
            Add-Member -InputObject $JsonObject -NotePropertyName mailbox -NotePropertyValue $MailboxName
            if ($AccessToken.Beta) {
                Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
            }
        }
        return $JsonObject		
    }
}

function ConvertFrom-SecureStringCustom {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [System.Security.SecureString]
        $SecureToken
    )
    process {
        #$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureToken)
        $Token = Unprotect-String -String $SecureToken
        return, $Token
    }
}

function Unprotect-String {
    <#
.SYNOPSIS
    Uses DPAPI to decrypt strings.

.DESCRIPTION
    Uses DPAPI to decrypt strings.
    Designed to reverse encryption applied by Protect-String

.PARAMETER String
    The string to decrypt.

.EXAMPLE
    PS C:\> Unprotect-String -String $secret

    Decrypts the content stored in $secret and returns it.
#>
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline = $true)]
        [System.Security.SecureString[]]
        $String
    )

    begin {
        Add-Type -AssemblyName System.Security -ErrorAction Stop
    }
    process {
        foreach ($item in $String) {
            $cred = New-Object PSCredential("irrelevant", $item)
            $stringBytes = [System.Convert]::FromBase64String($cred.GetNetworkCredential().Password)
            $decodedBytes = [System.Security.Cryptography.ProtectedData]::Unprotect($stringBytes, $null, 'CurrentUser')
            [Text.Encoding]::UTF8.GetString($decodedBytes)
        }
    }
}
$Script:EWSDLL = ""
$Script:Token = $null
$Script:MaxCount = 0
$Script:UseMaxCount = $false
$Script:MaxCountExceeded = $false
$script:ModuleRoot = $PSScriptRoot