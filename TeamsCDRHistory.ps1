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

function Get-TeamsMeetingsFolder{
    param( 
        [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $false)] [string]$AccessToken,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$basicAuth,
        [Parameter(Position = 6, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position =7, Mandatory = $false) ]  [Microsoft.Exchange.WebServices.Data.ExchangeService]$service

    )  
    Process {
        if($service -eq $null){
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
        }
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
        $TeamMeetingsFolderEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([System.Guid]::Parse("{07857F3C-AC6C-426B-8A8D-D1F7EA59F3C8}"), "TeamsMeetingsFolderEntryId", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
        $psPropset.Add($TeamMeetingsFolderEntryId)
        $RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)
        $FolderIdVal = $null

        $TeamsMeetingFolder = $null
        if ($RootFolder.TryGetProperty($TeamMeetingsFolderEntryId,[ref]$FolderIdVal))
        {
            $TeamMeetingFolderId= new-object Microsoft.Exchange.WebServices.Data.FolderId((ConvertId -HexId ([System.BitConverter]::ToString($FolderIdVal).Replace("-","")) -service $service))
            $TeamsMeetingFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$TeamMeetingFolderId);
        }
        return $TeamsMeetingFolder       
        

    }
}

function Get-TeamsCDRItems{
    param( 
        [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $false)] [string]$AccessToken,
        [Parameter(Position = 3, Mandatory = $false)] [string]$url,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$basicAuth,
        [Parameter(Position = 6, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 10, Mandatory = $false)] [DateTime]$startdatetime = (Get-Date).AddDays(-30),
        [Parameter(Position = 11, Mandatory = $false)] [datetime]$enddatetime = (Get-Date)

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
        $TeamsMeetingFolder = Get-TeamsMeetingsFolder -service $service -MailboxName $MailboxName
        $CalendarItemsHash = Get-CalendarItems -MailboxName $MailboxName -service $service -startdatetime $startdatetime -enddatetime $enddatetime
        $fiItems = $null 
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)   
        $ItemPropsetIdOnly = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $ivItemView.PropertySet = $ItemPropsetIdOnly
        $ItemPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        $CleanObjectId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting,0x23, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)  
        $PR_START_DATE  = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0060, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)  
        $PR_END_DATE  = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0061, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)  
        $ItemPropset.Add($CleanObjectId)
        $ItemPropset.Add($PR_START_DATE)
        $ItemPropset.Add($PR_END_DATE)
        #$ItemPropset.Add($SkypeTeamsProperties)
        $ItemPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
        $Sfgt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $startdatetime)
        $Sflt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $enddatetime)
        $sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
        $sfCollection.add($Sfgt)
        $sfCollection.add($Sflt)
        do { 
            $error.clear()
            try {
                $fiItems = $TeamsMeetingFolder.service.FindItems($TeamsMeetingFolder.Id,$sfCollection, $ivItemView)               
            }
            catch {
                write-host ("Error " + $_.Exception.Message)
                if ($_.Exception -is [Microsoft.Exchange.WebServices.Data.ServiceResponseException]) {
                    Write-Host ("EWS Error : " + ($_.Exception.ErrorCode))
                    Start-Sleep -Seconds 60 
                }
                $fiItems = $TeamsMeetingFolder.service.FindItems($TeamsMeetingFolder.Id, $ivItemView)           	
            }
            if ($fiItems.Items.Count -gt 0) {
                [Void]$TeamsMeetingFolder.service.LoadPropertiesForItems($fiItems, $ItemPropset)  
            }	
            if ($fiItems.Items.Count -gt 0) {
                Write-Host ("Processed " + $fiItems.Items.Count + " : " + $ItemClass)
                foreach ($Item in $fiItems.Items) { 
                    $rptObj = "" | Select-Object Subject,Start,End,Organizer,MeetingCid,CDRLog 
                    $Regex = [Regex]::new("(?<=/Thread Id\:)(.*)(?=/Communication Id\:)")           
                    $Match = $Regex.Match($Item.Subject)  
                    $rptObj.MeetingCid = $Match.Value.Trim()      
                    $rptObj.Subject = $Item.Subject
                    $rptObj.Start = $Item.Start
                    $rptObj.End = $Item.End
                    $PR_START_DATE_Value = $null
                    if($Item.TryGetProperty($PR_START_DATE,[ref]$PR_START_DATE_Value)){
                        $rptObj.Start = $PR_START_DATE_Value
                    }
                    $PR_END_DATE_Value = $null
                    if($Item.TryGetProperty($PR_END_DATE,[ref]$PR_END_DATE_Value)){
                        $rptObj.End = $PR_END_DATE_Value
                    }
                    if($CalendarItemsHash.ContainsKey($rptObj.MeetingCid)){
                        $rptObj.Subject = $CalendarItemsHash[$rptObj.MeetingCid].Subject
                        $rptObj.Organizer = $CalendarItemsHash[$rptObj.MeetingCid].Organizer
                    }
                    $rptObj.CDRLog = $Item.Body.Text                
                    Write-Output $rptObj
                }   
            }	  
            $ivItemView.Offset += $fiItems.Items.Count    
        }while ($fiItems.MoreAvailable) 
        

    }
}

function Get-CalendarItems{
    param (
    [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
    [Parameter(Position = 2, Mandatory = $false)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
    [Parameter(Position = 3, Mandatory = $false)] [DateTime]$startdatetime,
    [Parameter(Position = 4, Mandatory = $false)] [datetime]$enddatetime
    )
    process{
        $rptHash = @{}
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)   
        $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
        $Recurring = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment, 0x8223,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean); 
        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
        $SkypeTeamsProperties = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "SkypeTeamsProperties", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
        $psPropset.Add($Recurring)
        $psPropset.Add($SkypeTeamsProperties)
        $psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;      
  
          
        #Define the calendar view  
        $CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startdatetime,$enddatetime,1000)    
        $fiItems = $service.FindAppointments($Calendar.Id,$CalendarView)
        if($fiItems.Items.Count -gt 0){
            $type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
            $type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.Item" -as "Type")
            $ItemColl = [Activator]::CreateInstance($type)
            foreach($Item in $fiItems.Items){
                $ItemColl.Add($Item)
            }	
            [Void]$service.LoadPropertiesForItems($ItemColl,$psPropset)  
        }
        foreach($Item in $ItemColl){
            $SkypeTeamsProperties_Value = $null
            if($Item.TryGetProperty($SkypeTeamsProperties,[ref]$SkypeTeamsProperties_Value)){                 
                $SkypeTeamsPropertiesObj = ConvertFrom-Json  $SkypeTeamsProperties_Value
                if(!$rptHash.ContainsKey($SkypeTeamsPropertiesObj.cid)){
                    $rptHash.Add($SkypeTeamsPropertiesObj.cid,$Item)
                }
            }
        }        
        return $rptHash
    
    }
}


function ConvertId{    
    param (
            [Parameter(Position=1, Mandatory=$false)] [String]$HexId,
            [Parameter(Position=2, Mandatory=$false)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
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