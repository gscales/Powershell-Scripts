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
        Load-EWSManagedAPI
		
        ## Set Exchange Version  
        if ([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016) {
            $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
        }
        else {
            $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
        }
        
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
		  
        #Handle-SSL	
		  
        ## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  

        #CAS URL Option 1 Autodiscover  
        if ($url) {
            $uri = [system.URI] $url
            $service.Url = $uri    
        }
        else {
            $service.AutodiscoverUrl($MailboxName, {$true})  
        }
        #Write-host ("Using CAS Server : " + $Service.url)   
		   
        #CAS URL Option 2 Hardcoded  
		  
        #$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
        #$service.Url = $uri    
		  
        ## Optional section for Exchange Impersonation  
		  
        #$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
        if (!$service.URL) {
            throw "Error connecting to EWS"
        }
        else {		
            return, $service
        }
    }
}


function Load-EWSManagedAPI {
    param( 
    )  
    Begin {
        if(Test-Path ($script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll")){
            Import-Module ($script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll")
            $Script:EWSDLL = $script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll"
            write-verbose ("Using EWS dll from Local Directory")
        }else{

        
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

        ## end code from http://poshcode.org/624 ##

    }
}

function ConvertTo-String($ipInputString){    
    $Val1Text = ""    
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){    
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))    
            $clInt++    
    }    
    return $Val1Text    
}    

function TraceHandler(){
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
function Get-InboxRulesEWS{
    param( 
        [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $false)] [string]$AccessToken,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$basicAuth,
        [Parameter(Position = 6, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials 

    )  
    Process {
        if ($basicAuth.IsPresent) {
            if (!$Credentials) {
                $Credentials = Get-Credential
            }
            $service = Connect-Exchange -MailboxName $MailboxName -url $url -basicAuth -Credentials $Credentials
        }
        else {
            $service = Connect-Exchange -MailboxName $MailboxName -url $url -AccessToken $AccessToken
        } 
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)        
        }
        $InboxRules = $service.GetInboxRules($MailboxName)
        return $InboxRules
    }
}


function Invoke-GenericFolderItemEnum {
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [psObject]$Folder,
        [Parameter(Position = 2, Mandatory = $false)] [switch]$FullDetails,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$ReturnSentiment,       
        [Parameter(Position = 4, Mandatory = $false)] [Int]$MaxCount,
        [Parameter(Position = 5, Mandatory = $false)] [psobject]$fldIndex,
        [Parameter(Position = 6, Mandatory = $false)] [String]$QueryString
    )  
    Process {
        
        $PR_ENTRYID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FFF, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)  
        $RecoveryFolder = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "FIIM-RecoveryFolder", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
        if ($MaxCount -gt 0) {
            $Script:MaxCount = $MaxCount
            $Script:UseMaxCount = $true
            $Script:MaxCountExceeded = $false
        }
        else {
            $Script:MaxCount = 0
            $Script:UseMaxCount = $false
            $Script:MaxCountExceeded = $false
        }
        if ($Script:UseMaxCount) {
            if ($Script:MaxCount -gt 1000) {
                $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
            }
            else {
                $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($Script:MaxCount)
            }
        }
        else {
            $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
        }                
        $ItemPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        $ItemPropset.Add($PR_ENTRYID)
        $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Preview)
        $ItemPropset.Add($RecoveryFolder)
        if ($ReturnSentiment.IsPresent) {
            $Sentiment = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common, "EntityExtraction/Sentiment1.0", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
            $ItemPropset.Add($Sentiment)
        }
        $ItemPropsetIdOnly = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        if ($FullDetails.IsPresent) {
            $ivItemView.PropertySet = $ItemPropsetIdOnly
            if ($Script:UseMaxCount) {
                if ($Script:MaxCount -gt 100) {
                    $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)
                }
                else {
                    $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($Script:MaxCount)
                }
            }
            else {
                $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)
            }                     
        }
        else {
            $ivItemView.PropertySet = $ItemPropset
        }
        $itemCount = 0            
        $fiItems = $null    
        do { 
            $error.clear()
            try {
                if ([String]::IsNullOrEmpty($QueryString)) {
                    $fiItems = $Folder.service.FindItems($Folder.Id, $ivItemView) 
                }
                else {
                    $fiItems = $Folder.service.FindItems($Folder.Id, $QueryString, $ivItemView) 
                }
                
            }
            catch {
                write-host ("Error " + $_.Exception.Message)
                if ($_.Exception -is [Microsoft.Exchange.WebServices.Data.ServiceResponseException]) {
                    Write-Host ("EWS Error : " + ($_.Exception.ErrorCode))
                    Start-Sleep -Seconds 60 
                }	
                if ([String]::IsNullOrEmpty($QueryString)) {
                    $fiItems = $Folder.service.FindItems($Folder.Id, $ivItemView) 
                }
                else {
                    $fiItems = $Folder.service.FindItems($Folder.Id, $QueryString, $ivItemView) 
                }
            }
            if ($FullDetails.IsPresent) {
                if ($fiItems.Items.Count -gt 0) {
                    [Void]$Folder.service.LoadPropertiesForItems($fiItems, $ItemPropset)  
                }
            }	
            if ($fiItems.Items.Count -gt 0) {
                Write-Host ("Processed " + $fiItems.Items.Count + " : " + $ItemClass)
            }	  
                    
            foreach ($Item in $fiItems.Items) { 
                $itemCount ++
                $Okay = $true;
                if ($Script:UseMaxCount) {
                    if ($itemCount -gt $Script:MaxCount) {
                        $okay = $False                               
                    }
                }
                if ($Okay) {
                    $Item | Add-Member -Name "FolderPath" -Value $Folder.FolderPath -MemberType NoteProperty
                    Write-Output $Item
                }
                         
            }   
            if ($Script:UseMaxCount) {
                if ($itemCount -ge $Script:MaxCount) {
                    $Script:MaxCountExceeded = $true
                } 
            }
            $ivItemView.Offset += $fiItems.Items.Count    
        }while ($fiItems.MoreAvailable -eq $true -band (!$Script:MaxCountExceeded)) 
    }
    
}
function Invoke-GenericFolderConnect {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 4, Mandatory = $false)] [string]$FolderPath,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$Recurse,
        [Parameter(Position = 6, Mandatory = $false)] [switch]$RootFolder,
        [Parameter(Position = 7, Mandatory = $false)] [switch]$Archive,
        [Parameter(Position = 9, Mandatory = $false)] [switch]$basicAuth,
        [Parameter(Position = 10, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials  


    )  
    Process {
        if ($basicAuth.IsPresent) {
            if (!$Credentials) {
                $Credentials = Get-Credential
            }
            $service = Connect-Exchange -MailboxName $MailboxName -url $url -basicAuth -Credentials $Credentials
        }
        else {
            $service = Connect-Exchange -MailboxName $MailboxName -url $url -AccessToken $AccessToken
        }
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
       
        $folders = Get-FolderFromPath -MailboxName $MailboxName -FolderPath $FolderPath -service $service -Recurse:$Recurse.IsPresent -RootFolder:$RootFolder.IsPresent -Archive:$Archive.IsPresent
        return $folders 
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

function Get-AllItemsSearchFolder{
    param(
        [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$Archive
    )
    Process{
        if($Archive.IsPresent){
            Write-Verbose "Use Archive"
            $folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRoot,$MailboxName)
            $fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
            $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;
            $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, "allitems") 
            $findFolderResults = $service.FindFolders($folderidcnt, $SfSearchFilter, $fvFolderView) 
            if($findFolderResults.Folders.Count -eq 1){
                return $findFolderResults.Folders[0]
            }
            else{
                throw "Error getting All Item SearchFolders"
            }
        }else{
            $folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)
            $fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
            $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;
            $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, "allitems") 
            $findFolderResults = $service.FindFolders($folderidcnt, $SfSearchFilter, $fvFolderView) 
            if($findFolderResults.Folders.Count -eq 1){
                return $findFolderResults.Folders[0]
            }
            else{
                throw "Error getting All Item SearchFolders"
            }
        }


    }

}
function Get-FolderFromPath {
    param (
        [Parameter(Position = 0, Mandatory = $false)] [string]$FolderPath,
        [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
        [Parameter(Position = 3, Mandatory = $false)] [Microsoft.Exchange.WebServices.Data.PropertySet]$PropertySet,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$Recurse,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$RootFolder,
        [Parameter(Position = 6, Mandatory = $false)] [switch]$Archive
		  )
    process {
        Write-Verbose $RootFolder.IsPresent
        ## Find and Bind to Folder based on Path  
        #Define the path to search should be seperated with \  
        #Bind to the MSGFolder Root  
        $SourceKey = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65E0, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
        $psPropset.Add($SourceKey)
        #PR_POLICY_TAG 0x3019
        $PolicyTag = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3019, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        #PR_RETENTION_FLAGS 0x301D   
       	$RetentionFlags = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        #PR_RETENTION_PERIOD 0x301A
        $PR_ENTRYID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FFF, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary) 
       	$RetentionPeriod = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        $PidTagMessageSizeExtended = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0xe08, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
        $PR_ATTACH_ON_NORMAL_MSG_COUNT = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x66B1, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        $psPropset.Add($PolicyTag)
        $psPropset.Add($RetentionFlags)
        $psPropset.Add($RetentionPeriod)
        $psPropset.Add($PidTagMessageSizeExtended)
        $psPropset.Add($PR_ATTACH_ON_NORMAL_MSG_COUNT)
        $psPropset.Add($PR_ENTRYID)
        if ($Archive.IsPresent) {
            $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $MailboxName) 
        }
        else {
            $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName) 
        }          
        $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid, $psPropset)  
        if (!$RootFolder.IsPresent) {
            #Split the Search path into an array  
            $fldArray = $FolderPath.Split("\") 
            if ($fldArray.Length -lt 2) {throw "No Root Folder"}
            #Loop through the Split Array and do a Search for each level of folder 
            for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
                #Perform search based on the displayname of each folder level 
                $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
                $fvFolderView.PropertySet = $psPropset
                $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $fldArray[$lint]) 
                $findFolderResults = $service.FindFolders($tfTargetFolder.Id, $SfSearchFilter, $fvFolderView) 
                $tfTargetFolder = $null  
                if ($findFolderResults.TotalCount -gt 0) { 
                    foreach ($folder in $findFolderResults.Folders) { 
                        $tfTargetFolder = $folder                
                    } 
                } 
                else { 
                    Write-host ("Error Folder Not Found check path and try again") -ForegroundColor Red
                    $tfTargetFolder = $null  
                    break  
                }     
            }  
        }
        else {
            $FolderPath = "\"
        }
        if ($tfTargetFolder -ne $null) {
            if ($Archive.IsPresent) {
                $tfTargetFolder | Add-Member -Name "FolderPath" -Value ("\InPlace-Archive" + $FolderPath) -MemberType NoteProperty
            }
            else {
                $tfTargetFolder | Add-Member -Name "FolderPath" -Value $FolderPath -MemberType NoteProperty
            }
            
            $tfTargetFolder | Add-Member -Name "Mailbox" -Value $MailboxName -MemberType NoteProperty
            $prop1Val = $null    
            if ($tfTargetFolder.TryGetProperty($PolicyTag, [ref] $prop1Val)) {     
                $rtnStringVal =	[System.BitConverter]::ToString($prop1Val).Replace("-", "");
                $rtnStringVal = $rtnStringVal.Substring(6, 2) + $rtnStringVal.Substring(4, 2) + $rtnStringVal.Substring(2, 2) + $rtnStringVal.Substring(0, 2) + "-" + $rtnStringVal.Substring(10, 2) + $rtnStringVal.Substring(8, 2) + "-" + $rtnStringVal.Substring(14, 2) + $rtnStringVal.Substring(12, 2) + "-" + $rtnStringVal.Substring(16, 2) + $rtnStringVal.Substring(18, 2) + "-" + $rtnStringVal.Substring(20, 12)           
                Add-Member -InputObject $tfTargetFolder -MemberType NoteProperty -Name PR_POLICY_TAG -Value $rtnStringVal      
            }
            $prop2Val = $null
            if ($tfTargetFolder.TryGetProperty($RetentionFlags, [ref] $prop2Val)) {
                Add-Member -InputObject $tfTargetFolder -MemberType NoteProperty -Name PR_RETENTION_FLAGS -Value $prop2Val
            }
            $prop3Val = $null
            if ($tfTargetFolder.TryGetProperty($RetentionPeriod, [ref] $prop3Val)) {
                Add-Member -InputObject $tfTargetFolder -MemberType NoteProperty -Name PR_RETENTION_PERIOD -Value $prop3Val
            }
            $prop4Val = $null
            if ($tfTargetFolder.TryGetProperty($PidTagMessageSizeExtended, [ref]  $prop4Val)) {
                Add-Member -InputObject $tfTargetFolder -MemberType NoteProperty -Name FolderSize -Value $prop4Val
            }
            $prop5Val = $null
            if ($tfTargetFolder.TryGetProperty($PR_ATTACH_ON_NORMAL_MSG_COUNT, [ref]  $prop5Val)) {
                Add-Member -InputObject $tfTargetFolder -MemberType NoteProperty -Name PR_ATTACH_ON_NORMAL_MSG_COUNT -Value $prop5Val
            }
            $prop6Val = $null
            if ($tfTargetFolder.TryGetProperty($PR_ENTRYID, [ref]  $prop6Val)) {
                $HexVal = [System.BitConverter]::ToString($prop6Val)
                Add-Member -InputObject $tfTargetFolder -MemberType NoteProperty -Name PR_ENTRYID -Value $HexVal.Replace("-", "")
            }
            $GetLastItem = {
                $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
                $fiItems = $this.FindItems($ivItemView)
                if ($fiItems.Items.Count -eq 1) {
                    $fiItems.Items[0].Load()                    
                    $fiItems.Items[0] | Add-Member -Name "FolderPath" -Value $this.FolderPath -MemberType NoteProperty
                    Add-Member -InputObject $this -MemberType NoteProperty -Name LastItem -Value $fiItems.Items[0] -Force
                    return $fiItems.Items[0]
                }                
            }
            Add-Member -InputObject $tfTargetFolder -MemberType ScriptMethod -Name GetLastItem -Value $GetLastItem
            if ($Recurse.IsPresent) {
                $Folders = @()
                $Folders += $tfTargetFolder
                $Folders = Get-SubFolders -ParentFolder $tfTargetFolder -Folders $Folders -Archive:$Archive.IsPresent
                return, [PSObject]$Folders
            }
            else {
                return, [Microsoft.Exchange.WebServices.Data.Folder]$tfTargetFolder
            }
           
        }
        else {
            return, $null
        }
    }
}
function Get-SubFolders { 
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.Folder]$ParentFolder,
        [Parameter(Position = 1, Mandatory = $false)] [Microsoft.Exchange.WebServices.Data.PropertySet]$PropertySet,
        [Parameter(Position = 2, Mandatory = $false)] [psObject]$Folders,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$Archive

        
    )  
    Begin {
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
            $PropertySet.Add($PR_POLICY_TAG)
            $PropertySet.Add($RetentionFlags)
            $PropertySet.Add($RetentionPeriod)
            $PropertySet.Add($SourceKey)
            $PropertySet.Add($PidTagMessageSizeExtended)
            $PropertySet.Add($PR_ATTACH_ON_NORMAL_MSG_COUNT)
            $PropertySet.Add($PR_ENTRYID)
        }	
		
        #Define Extended properties  
        $PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
        #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
        $fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
        #Deep Transval will ensure all folders in the search path are returned  
        $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
        $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
        #Add Properties to the  Property Set  
        $PropertySet.Add($PR_Folder_Path);  
        $fvFolderView.PropertySet = $PropertySet;  
        #The Search filter will exclude any Search Folders  
        $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE, "1")  
        $fiResult = $null  
        #The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
        do {  
            $fiResult = $ParentFolder.FindFolders($sfSearchFilter, $fvFolderView)  
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
                $prop2Val = $null
                if ($ffFolder.TryGetProperty($RetentionFlags, [ref] $prop2Val)) {
                    Add-Member -InputObject $ffFolder -MemberType NoteProperty -Name PR_RETENTION_FLAGS -Value $prop2Val
                }
                $prop3Val = $null
                if ($ffFolder.TryGetProperty($RetentionPeriod, [ref] $prop3Val)) {
                    Add-Member -InputObject $ffFolder -MemberType NoteProperty -Name PR_RETENTION_PERIOD -Value $prop3Val
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
                $GetLastItem = {
                    $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
                    $fiItems = $this.FindItems($ivItemView)
                    if ($fiItems.Items.Count -eq 1) {
                        $fiItems.Items[0].Load()
                        $fiItems.Items[0] | Add-Member -Name "FolderPath" -Value $this.FolderPath -MemberType NoteProperty
                        Add-Member -InputObject $this -MemberType NoteProperty -Name LastItem -Value $fiItems.Items[0] -Force
                        return $fiItems.Items[0]
                    }                
                }
                Add-Member -InputObject $ffFolder -MemberType ScriptMethod -Name GetLastItem -Value $GetLastItem
                $Folders += $ffFolder
            } 
            $fvFolderView.Offset += $fiResult.Folders.Count
        }while ($fiResult.MoreAvailable -eq $true)  
        return, $Folders	
    }
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
function ConvertToString($ipInputString) {  
    $Val1Text = ""  
    for ($clInt = 0; $clInt -lt $ipInputString.length; $clInt++) {  
        $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt, 2), 16)))  
        $clInt++  
    }  
    return $Val1Text  
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
        
            $Phase1auth = Show-OAuthWindow -Url "https://login.microsoftonline.com/common/oauth2/authorize?resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&response_type=code&redirect_uri=$redirectUrl&prompt=$Prompt&domain_hint=organizations"
            $code = $Phase1auth["code"]
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
    $DocComp = {
        $Global:uri = $web.Url.AbsoluteUri
        if ($Global:Uri -match "error=[^&]*|code=[^&]*") { $form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown( { $form.Activate() })
    $form.ShowDialog() | Out-Null
    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{ }
    foreach ($key in $queryOutput.Keys) {
        $output["$key"] = $queryOutput[$key]
    }
    return $output
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