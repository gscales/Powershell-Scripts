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




function Invoke-GenericFolderItemEnum {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 4, Mandatory = $true)] [string]$FolderPath,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$Recurse,
        [Parameter(Position = 6, Mandatory = $false)] [switch]$FullDetails,
        [Parameter(Position = 6, Mandatory = $false)] [switch]$ReturnSentiment,
        [Parameter(Position = 7, Mandatory = $false)] [Int]$MaxCount 
    )  
    Process {
        $folders = Invoke-GenericFolderConnect -MailboxName $MailboxName -Credentials $Credentials -url $url -useImpersonation:$useImpersonation.IsPresent -FolderPath $FolderPath -Recurse:$Recurse.IsPresent
        foreach ($Folder in $folders) {
                if($MaxCount -gt 0){
                    $Script:MaxCount = $MaxCount
                    $Script:UseMaxCount = $true
                    $Script:MaxCountExceeded = $false
                }
                else{
                    $Script:MaxCount = 0
                    $Script:UseMaxCount = $false
                    $Script:MaxCountExceeded = $false
                }
                if($Script:UseMaxCount){
                    if($Script:MaxCount -gt 1000){
                        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
                    }
                    else{
                        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($Script:MaxCount)
                    }
                }else{
                    $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
                }                
                $ItemPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                if($ReturnSentiment.IsPresent){
                    $Sentiment = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common,"EntityExtraction/Sentiment1.0", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
                    $ItemPropset.Add($Sentiment)
                }
                $ItemPropsetIdOnly = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                if ($FullDetails.IsPresent) {
                    $ivItemView.PropertySet = $ItemPropsetIdOnly
                    if($Script:UseMaxCount){
                        if($Script:MaxCount -gt 100){
                            $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)
                        }
                        else{
                             $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($Script:MaxCount)
                        }
                    }else{
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
                        $fiItems = $Folder.service.FindItems($Folder.Id, $ivItemView) 
                    }
                    catch {
                        write-host ("Error " + $_.Exception.Message)
                        if ($_.Exception -is [Microsoft.Exchange.WebServices.Data.ServiceResponseException]) {
                            Write-Host ("EWS Error : " + ($_.Exception.ErrorCode))
                            Start-Sleep -Seconds 60 
                        }	
                        $fiItems = $Folder.service.FindItems($Folder.Id, $ivItemView) 
                    }
                    if ($FullDetails.IsPresent) {
                        if ($fiItems.Items.Count -gt 0) {
                            [Void]$Folder.service.LoadPropertiesForItems($fiItems, $ItemPropset)  
                        }
                    }	
                    if($fiItems.Items.Count -gt 0){
                        Write-Host ("Processed " + $fiItems.Items.Count + " : " + $ItemClass)
                    }	  
                    
                    foreach ($Item in $fiItems.Items) { 
                        $itemCount ++
                        $Okay = $true;
                        if($Script:UseMaxCount){
                            if($itemCount -gt $Script:MaxCount){
                                $okay = $False                               
                            }
                        }
                        if($Okay){
                            $DownloadAttachments = {
                                param([Parameter(Position = 0, Mandatory = $true)] [string]$FolderPath)
                                $this.Load()
                                foreach($attach in $this.Attachments){
                                    $attach.Load()
                                    $fiFile = new-object System.IO.FileStream(($FolderPath + "\" + $attach.Name.ToString()), [System.IO.FileMode]::Create)
                                    $fiFile.Write($attach.Content, 0, $attach.Content.Length)
                                    $fiFile.Close()
                                    write-host ("Downloaded Attachment : " + (($FolderPath + "\" + $attach.Name.ToString())))
                                }
                            }
                             if($ReturnSentiment.IsPresent){
                                $SentimentValue = $null
                                if($Item.TryGetProperty($Sentiment,[ref]$SentimentValue)){
                                    $emotiveProfile =ConvertFrom-Json -InputObject $SentimentValue
				                    Add-Member -InputObject $Item -NotePropertyName "Sentiment" -NotePropertyValue $emotiveProfile.sentiment.polarity
				                    Add-Member -InputObject $Item -NotePropertyName "EmotiveProfile" -NotePropertyValue $emotiveProfile
                                }
                             }
                            Add-Member -InputObject $Item -MemberType ScriptMethod -Name DownloadAttachments -Value $DownloadAttachments 
                            $Item | Add-Member -Name "FolderPath" -Value $Folder.FolderPath -MemberType NoteProperty
                            Write-Output $Item
                        }
                         
                    }   
                    if($Script:UseMaxCount){
                        if($itemCount -ge $Script:MaxCount){
                            $Script:MaxCountExceeded = $true
                        } 
                    }
                    $ivItemView.Offset += $fiItems.Items.Count    
                }while ($fiItems.MoreAvailable -eq $true -band (!$Script:MaxCountExceeded)) 
        }
    }
}
function Invoke-GenericFolderConnect {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 4, Mandatory = $false)] [string]$FolderPath,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$Recurse,
        [Parameter(Position = 6, Mandatory = $false)] [switch]$RootFolder
    )  
    Process {
        $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)        
        }
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
       
        $folders = Get-FolderFromPath -MailboxName $MailboxName -FolderPath $FolderPath -service $service -Recurse:$Recurse.IsPresent -RootFolder:$RootFolder.IsPresent
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
function Get-FolderFromPath {
    param (
        [Parameter(Position = 0, Mandatory = $false)] [string]$FolderPath,
        [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
        [Parameter(Position = 3, Mandatory = $false)] [Microsoft.Exchange.WebServices.Data.PropertySet]$PropertySet,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$Recurse,
        [Parameter(Position = 6, Mandatory = $false)] [switch]$RootFolder
		  )
    process {
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
       	$RetentionPeriod = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        $PidTagMessageSizeExtended = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0xe08, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
        $PR_ATTACH_ON_NORMAL_MSG_COUNT = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x66B1, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
        $psPropset.Add($PolicyTag)
        $psPropset.Add($RetentionFlags)
        $psPropset.Add($RetentionPeriod)
        $psPropset.Add($PidTagMessageSizeExtended)
        $psPropset.Add($PR_ATTACH_ON_NORMAL_MSG_COUNT)

        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName)   
        $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)  
        if(!$RootFolder.IsPresent){
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
        else{
            $FolderPath = "\"
        }
        if ($tfTargetFolder -ne $null) {
            $tfTargetFolder | Add-Member -Name "FolderPath" -Value $FolderPath -MemberType NoteProperty
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
            $GetLastItem = {
                $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
                $fiItems =  $this.FindItems($ivItemView)
                if($fiItems.Items.Count -eq 1){
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
                $Folders = Get-SubFolders -ParentFolder $tfTargetFolder -Folders $Folders
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
        [Parameter(Position = 2, Mandatory = $false)] [psObject]$Folders
        
    )  
    Begin {
        if ([string]::IsNullOrEmpty($PropertySet)) {
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
                $ffFolder | Add-Member -Name "FolderPath" -Value $fpath -MemberType NoteProperty
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
                $GetLastItem = {
                    $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
                    $fiItems =  $this.FindItems($ivItemView)
                    if($fiItems.Items.Count -eq 1){
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

$Script:MaxCount = 0
$Script:UseMaxCount = $false
$Script:MaxCountExceeded = $false