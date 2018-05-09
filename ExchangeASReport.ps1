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


function Send-EWSMessage  {
     param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position=2, Mandatory=$false)] [switch]$useImpersonation,
        [Parameter(Position=3, Mandatory=$false)] [string]$url,
        [Parameter(Position=6, Mandatory=$true)] [String]$To,
        [Parameter(Position=7, Mandatory=$true)] [String]$Subject,
        [Parameter(Position=8, Mandatory=$true)] [String]$Body,
        [Parameter(Position=9, Mandatory=$false)] [String]$Attachment
    )  
  Begin
 {
    $service = Connect-Exchange -MailboxName $MailboxName -Credentials $Credentials -url $url
    if ($useImpersonation.IsPresent) {
        $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)        
    }
    $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);  
    $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$MailboxName)   
    $SentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
    $EmailMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service  
    $EmailMessage.Subject = $Subject
    #Add Recipients    
    $EmailMessage.ToRecipients.Add($To)  
    $EmailMessage.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody  
    $EmailMessage.Body.BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML  
    $EmailMessage.Body.Text = $Body
    $EmailMessage.From = $MailboxName
    if($Attachment)
    {   
    $EmailMessage.Attachments.AddFileAttachment($Attachment)
    }
    $EmailMessage.SendAndSaveCopy($SentItems.Id) 
  
 }
}

function Get-EWSAntiSpamReport {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url,
        [Parameter(Position = 3, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 4, Mandatory = $true)] [string]$FolderPath, 
        [Parameter(Position = 7, Mandatory = $false)] [Int]$MaxCount 
    )  
    Process {
        $folders = Invoke-GenericFolderConnect -MailboxName $MailboxName -Credentials $Credentials -url $url -useImpersonation:$useImpersonation.IsPresent -FolderPath $FolderPath -Recurse:$false
        $PR_ENTRYID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FFF,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)  
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
                $ItemPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                $ItemPropset.Add($PR_ENTRYID)
                $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Preview)
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
                    if ($fiItems.Items.Count -gt 0) {
                           $HeaderPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                           $HeaderPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageHeaders)
                           $HeaderPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject)
                           $HeaderPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender)
                           $HeaderPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From)
                           $HeaderPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived)
                           $HeaderPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::WebClientReadFormQueryString)
                           $HeaderPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments)
                           [Void]$Folder.service.LoadPropertiesForItems($fiItems, $HeaderPropset)  
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
                            $Item | Add-Member -Name "FolderPath" -Value $Folder.FolderPath -MemberType NoteProperty
                            $Item | Add-Member -Name "SenderEmailAddress" -Value $Item.From.Address -MemberType NoteProperty
                            Invoke-EXRProcessAntiSPAMHeaders -Item $Item
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

function Invoke-EXRProcessAntiSPAMHeaders {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $Item,
        [Parameter(Position = 2, Mandatory = $false)]
        [switch]
        $noForward
    )
	
    process {
           if ([bool]($Item.PSobject.Properties.name -match "IndexedInternetMessageHeaders"))
            {
                if($Item.IndexedInternetMessageHeaders.ContainsKey("Authentication-Results")){
                    $AuthResultsText = $Item.IndexedInternetMessageHeaders["Authentication-Results"]
                    $SPFResults =  [regex]::Match($AuthResultsText,("spf=(.*?)dkim="))
                    if($SPFResults.Groups.Count -gt 0){
                       $SPF =  $SPFResults.Groups[1].Value    
                    }
                    $DKIMResults =  [regex]::Match($AuthResultsText,("dkim=(.*?)dmarc="))
                    if($DKIMResults.Groups.Count -gt 0){
                       $DKIM =  $DKIMResults.Groups[1].Value    
                    }
                    $DMARCResults =  [regex]::Match($AuthResultsText,("dmarc=(.*?)compauth="))
                    if($DMARCResults.Groups.Count -gt 0){
                       $DMARC =  $DMARCResults.Groups[1].Value    
                    }
                    $CompAuthResults =  [regex]::Match($AuthResultsText,("compauth=(.*)"))
                    if($CompAuthResults.Groups.Count -gt 0){
                       $CompAuth =  $CompAuthResults.Groups[1].Value    
                    }
                    Add-Member -InputObject $Item -NotePropertyName "SPF" -NotePropertyValue $SPF -Force
                    Add-Member -InputObject $Item -NotePropertyName "DKIM" -NotePropertyValue $DKIM  -Force
                    Add-Member -InputObject $Item -NotePropertyName "DMARC" -NotePropertyValue $DMARC  -Force
                    Add-Member -InputObject $Item -NotePropertyName "CompAuth" -NotePropertyValue $CompAuth  -Force
                 }
                if($Item.IndexedInternetMessageHeaders.ContainsKey("Authentication-Results-Original")){
                    $AuthResultsText = $Item.IndexedInternetMessageHeaders["Authentication-Results-Original"]
                    $SPFResults =  [regex]::Match($AuthResultsText,("spf=(.*?)\;"))
                    if($SPFResults.Groups.Count -gt 0){
                       $SPF =  $SPFResults.Groups[1].Value    
                    }
                    $DKIMResults =  [regex]::Match($AuthResultsText,("dkim=(.*?)\;"))
                    if($DKIMResults.Groups.Count -gt 0){
                       $DKIM =  $DKIMResults.Groups[1].Value    
                    }
                    $DMARCResults =  [regex]::Match($AuthResultsText,("dmarc=(.*?)\;"))
                    if($DMARCResults.Groups.Count -gt 0){
                       $DMARC =  $DMARCResults.Groups[1].Value    
                    }
                    $CompAuthResults =  [regex]::Match($AuthResultsText,("compauth=(.*)"))
                    if($CompAuthResults.Groups.Count -gt 0){
                       $CompAuth =  $CompAuthResults.Groups[1].Value    
                    }
                    Add-Member -InputObject $Item -NotePropertyName "Original-SPF" -NotePropertyValue $SPF -Force
                    Add-Member -InputObject $Item -NotePropertyName "Original-DKIM" -NotePropertyValue $DKIM  -Force
                    Add-Member -InputObject $Item -NotePropertyName "Original-DMARC" -NotePropertyValue $DMARC  -Force
                    Add-Member -InputObject $Item -NotePropertyName "Original-CompAuth" -NotePropertyValue $CompAuth  -Force
                }
                if ($Item.IndexedInternetMessageHeaders.ContainsKey("X-Microsoft-Antispam")){
                    $ASReport = $Item.IndexedInternetMessageHeaders["X-Microsoft-Antispam"]              
                    $PCLResults = [regex]::Match($ASReport,("PCL\:(.*?)\;"))
                    if($PCLResults.Groups.Count -gt 0){
                        $PCL =  $PCLResults.Groups[1].Value    
                    }
                    $BCLResults = [regex]::Match($ASReport,("BCL\:(.*?)\;"))
                    if($BCLResults.Groups.Count -gt 0){
                        $BCL =  $BCLResults.Groups[1].Value    
                    }
                    Add-Member -InputObject $Item -NotePropertyName "PCL" -NotePropertyValue $PCL  -Force
                    Add-Member -InputObject $Item -NotePropertyName "BCL" -NotePropertyValue $BCL  -Force
                }
                if ($Item.IndexedInternetMessageHeaders.ContainsKey("X-Forefront-Antispam-Report")){
                    $ASReport = $Item.IndexedInternetMessageHeaders["X-Forefront-Antispam-Report"]              
                    $CTRYResults = [regex]::Match($ASReport,("CTRY\:(.*?)\;"))
                    if($CTRYResults.Groups.Count -gt 0){
                        $CTRY =  $CTRYResults.Groups[1].Value    
                    }
                    $SFVResults = [regex]::Match($ASReport,("SFV\:(.*?)\;"))
                    if($SFVResults.Groups.Count -gt 0){
                        $SFV =  $SFVResults.Groups[1].Value    
                    }
                    $SRVResults = [regex]::Match($ASReport,("SRV\:(.*?)\;"))
                    if($SRVResults.Groups.Count -gt 0){
                        $SRV =  $SRVResults.Groups[1].Value    
                    }
                    $PTRResults = [regex]::Match($ASReport,("PTR\:(.*?)\;"))
                    if($PTRResults.Groups.Count -gt 0){
                        $PTR =  $PTRResults.Groups[1].Value    
                    }   
                    $CIPResults = [regex]::Match($ASReport,("CIP\:(.*?)\;"))
                    if($CIPResults.Groups.Count -gt 0){
                        $CIP =  $CIPResults.Groups[1].Value    
                    }      
                    $IPVResults = [regex]::Match($ASReport,("IPV\:(.*?)\;"))
                    if($IPVResults.Groups.Count -gt 0){
                        $IPV =  $IPVResults.Groups[1].Value    
                    }                   
                    Add-Member -InputObject $Item -NotePropertyName "CTRY" -NotePropertyValue $CTRY  -Force
                    Add-Member -InputObject $Item -NotePropertyName "SFV" -NotePropertyValue $SFV  -Force
                    Add-Member -InputObject $Item -NotePropertyName "SRV" -NotePropertyValue $SRV  -Force
                    Add-Member -InputObject $Item -NotePropertyName "PTR" -NotePropertyValue $PTR  -Force
                    Add-Member -InputObject $Item -NotePropertyName "CIP" -NotePropertyValue $CIP  -Force
                    Add-Member -InputObject $Item -NotePropertyName "IPV" -NotePropertyValue $IPV  -Force
                }
                
                if ($Item.IndexedInternetMessageHeaders.ContainsKey("X-MS-Exchange-Organization-SCL")){
                    Add-Member -InputObject $Item -NotePropertyName "SCL" -NotePropertyValue $Item.IndexedInternetMessageHeaders["X-MS-Exchange-Organization-SCL"]  -Force 
                }
                if ($Item.IndexedInternetMessageHeaders.ContainsKey("X-CustomSpam")){
                    Add-Member -InputObject $Item -NotePropertyName "ASF" -NotePropertyValue $Item.IndexedInternetMessageHeaders["X-CustomSpam"]  -Force 
                }
                
               
            }else{
                if(!$noForward.IsPresent){
                    if ([bool]($Item.PSobject.Properties.name -match "InternetMessageHeaders")){
                        $IndexedHeaders = New-Object 'system.collections.generic.dictionary[string,string]'
                        foreach($header in $Item.InternetMessageHeaders){
                            if(!$IndexedHeaders.ContainsKey($header.name)){
                                $IndexedHeaders.Add($header.name,$header.value)
                            }
                        }
                        Add-Member -InputObject $Item -NotePropertyName "IndexedInternetMessageHeaders" -NotePropertyValue $IndexedHeaders
                    }
                    Invoke-EXRProcessAntiSPAMHeaders -Item $Item -noForward
                }

            }  

            
               
		
    }
}
function Get-EWSDigestEmailBody
{
	[CmdletBinding()] 
    param (
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$MessageList,
		[Parameter(Position = 2, Mandatory = $false)]
		[switch]
		$weblink,
		[Parameter(Position = 3, Mandatory = $false)]
		[switch]
		$Detail,
		[Parameter(Position = 4, Mandatory = $false)]
		[String]
		$InfoField1Name,
		[Parameter(Position = 5, Mandatory = $false)]
		[String]
		$InfoField2Name,
		[Parameter(Position = 6, Mandatory = $false)]
		[String]
		$InfoField3Name,
		[Parameter(Position = 7, Mandatory = $false)]
		[String]
		$InfoField4Name,
		[Parameter(Position = 8, Mandatory = $false)]
		[String]
        $InfoField5Name
	)
	
 	process
	{
        $PR_ENTRYID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FFF,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)  
		if($Detail.IsPresent){
		$rpReport = ""
		foreach ($message in $MessageList){
			$PR_ENTRYIDValue = $null
            if($message.TryGetProperty($PR_ENTRYID,[ref]$PR_ENTRYIDValue)){
                $Oulookid = [System.BitConverter]::ToString($PR_ENTRYIDValue).Replace("-","")
            }	
			$fromstring = $message.From.Address
			if ($fromstring.length -gt 30){$fromstring = $fromstring.Substring(0,30)}
			$HeaderLine  = $message.DateTimeReceived.ToString("G") + " : " + $fromstring + " : " + $message.Subject
			$BodyLine = $message.Preview
			if($weblink.IsPresent){
				$BodyLine += "`r`n</br></br><a href=`"" + $message.WebClientReadFormQueryString + "`">MoreInfo</a href>"
			}else{
				$BodyLine += "`r`n</br></br><a href=`"outlook:" + $Oulookid + "`">MoreInfo</a href>"
			}
			
			$InfoField1Value = $message.$InfoField1Name
			$InfoField2Value = $message.$InfoField2Name
			$InfoField3Value = $message.$InfoField3Name
			$InfoField4Value = $message.$InfoField4Name
			$InfoField5Value = $message.$InfoField5Name
$nextTable = @"
<div style=" text-align: left; text-indent: 0px; padding: 0px 0px 0px 0px; margin: 0px 0px 0px 0px;">
<table width="100%" border="1" cellpadding="0" cellspacing="0" style="border-width: 0px; background-color: #ffffff;">
<tr valign="top">
<td colspan=5 style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; border-style: solid;">
<p style=" text-align: center; text-indent: 0px; padding: 0px 0px 0px 0px; margin: 0px 0px 0px 0px; background-color: #3366ff;">
<span style=" font-size: 10pt; alignment-adjust: central; font-family: 'Arial', 'Helvetica', sans-serif; font-style: normal; font-weight: bold; color: #ffffff;  text-decoration: none;">
$HeaderLine</span></p>
</td>
</tr>
<tr valign="top">
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; font-weight: bold; border-style: solid;">$InfoField1Name<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; font-weight: bold; border-style: solid;">$InfoField2Name<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; font-weight: bold; border-style: solid;">$InfoField3Name<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; font-weight: bold; border-style: solid;">$InfoField4Name<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; font-weight: bold; border-style: solid;">$InfoField5Name<br />
</td>
</tr>
<tr valign="top">
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; border-style: solid;">$InfoField1Value<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; border-style: solid;">$InfoField2Value<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; border-style: solid;">$InfoField3Value<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; border-style: solid;">$InfoField4Value<br />
</td>
<td width="20%" style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; text-align: center; border-style: solid;">$InfoField5Value<br />
</td>
</tr>
<td colspan=5 style="border-width : 1px; border-color : #000000 #000000 #000000 #000000; border-style: solid;">
<p style=" text-align: left; text-indent: 0px; padding: 0px 0px 0px 0px; margin: 0px 0px 0px 0px; >
<span style=" font-size: 10pt; alignment-adjust: central; font-family: 'Arial', 'Helvetica', sans-serif; font-style: normal; font-weight: bold; color: #ffffff;  text-decoration: none;">$BodyLine</span></p>
</td>
</table>
</div>
</br>
"@
			 $rpReport += $nextTable
		}
	}
	else{
		$rpReport = $rpReport + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:15%;`" ><b>Recieved</b></td>" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:20%;`" ><b>From</b></td>" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:60%;`" ><b>Subject</b></td>" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:5%;`" ><b>Size</b></td>" +"`r`n"
		$rpReport = $rpReport + "</tr>" + "`r`n"
		foreach ($message in $MessageList){
            $fromstring = $message.From.Address
            $PR_ENTRYIDValue = $null
            if($message.TryGetProperty($PR_ENTRYID,[ref]$PR_ENTRYIDValue)){
                $Oulookid = [System.BitConverter]::ToString($PR_ENTRYIDValue).Replace("-","")
            }			
			if ($fromstring.length -gt 30){$fromstring = $fromstring.Substring(0,30)}
			$rpReport = $rpReport + "  <tr>"  + "`r`n"
			$rpReport = $rpReport + "<td>" + [DateTime]::Parse($message.DateTimeReceived).ToString("G") + "</td>"  + "`r`n"
			$rpReport = $rpReport + "<td>" +  $fromstring + "</td>"  + "`r`n"
			if($weblink.IsPresent){
				$rpReport = $rpReport + "<td><a href=`"" + $message.WebClientReadFormQueryString + "`">" + $message.Subject + "</td>"  + "`r`n"
			}
			else{
				$rpReport = $rpReport + "<td><a href=`"outlook:" + $Oulookid + "`">" + $message.Subject + "</td>"  + "`r`n"
			}			
			$rpReport = $rpReport + "<td>" +  ($message.Size/1024).ToString(0.00) + "</td>"  + "`r`n"
			$rpReport = $rpReport + "</tr>"  + "`r`n"
		}
		$rpReport = $rpReport + "</table>"  + "  " 
		
	}
	return $rpReport
    }
}


$Script:MaxCount = 0
$Script:UseMaxCount = $false
$Script:MaxCountExceeded = $false