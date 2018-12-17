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

        ## end code from http://poshcode.org/624 ##

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
   
function Get-ZapStatistics {
    param( 
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $false)] [string]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [string]$url,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$useImpersonation,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$basicAuth,
        [Parameter(Position = 6, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position = 10, Mandatory = $false)] [DateTime]$startdatetime = (Get-Date).AddDays(-365),
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
        if ($useImpersonation.IsPresent) {
            $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)        
        }
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName); 
        if ($AccessToken) {
            $Script:Token = $AccessToken
        }
        $rptCollection = @()
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail, $MailboxName)         
        $JunkMailFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)  
        $ZapInfo = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common, "X-Microsoft-Antispam-ZAP-Message-Info", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
        $PidTagInternetMessageId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(4149, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
        $Sfexists = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Exists($ZapInfo) 
        $Sfgt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $startdatetime)
        $Sflt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $enddatetime)
        $sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
        $sfCollection.add($Sfgt)
        $sfCollection.add($Sflt)
        $sfCollection.add($Sfexists)
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
        $PidTagSenderSmtpAddress = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x5D01,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $psPropset.Add($PidTagSenderSmtpAddress)
        $ivItemView.PropertySet = $psPropset
        $fiItems = $null    
        do { 
            $error.clear()
            try {
                $fiItems = $JunkMailFolder.FindItems($sfCollection, $ivItemView)                 
            }
            catch {
                write-host ("Error " + $_.Exception.Message)
                if ($_.Exception -is [Microsoft.Exchange.WebServices.Data.ServiceResponseException]) {
                    Write-Host ("EWS Error : " + ($_.Exception.ErrorCode))
                    Start-Sleep -Seconds 60 
                }	
                $fiItems = $JunkMailFolder.FindItems($sfCollection, $ivItemView)     
            }
            if ($fiItems.Items.Count -gt 0) {
                $HeaderPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $HeaderPropset.Add($PidTagInternetMessageId)
                $HeaderPropset.Add($PidTagSenderSmtpAddress)
                [Void]$JunkMailFolder.service.LoadPropertiesForItems($fiItems, $HeaderPropset)  
            }                              
            foreach ($Item in $fiItems.Items) {
                $rptObject = "" | Select DateTimeReceived,LastModified,MinutesInInbox,HoursInInbox,Subject,Sender,Read,InternetMessageId
                $rptObject.DateTimeReceived = $Item.DateTimeReceived 
                $rptObject.LastModified = $Item.LastModifiedTime
                $TimeSpan = New-TimeSpan -Start $Item.DateTimeReceived -End $Item.LastModifiedTime
                $rptObject.MinutesInInbox = [Math]::Round($TimeSpan.TotalMinutes,0)
                $rptObject.HoursInInbox = [Math]::Round($TimeSpan.TotalHours,0)
                $rptObject.Subject = $Item.Subject
                $rptObject.Read = $Item.isRead
                $EmailValue = $null
                if($Item.TryGetProperty($PidTagSenderSmtpAddress,[ref]$EmailValue)){
                    $rptObject.Sender = $EmailValue
                }
                $InternetMessageIdValue = $null
                if($Item.TryGetProperty($PidTagInternetMessageId,[ref]$InternetMessageIdValue)){
                    $rptObject.InternetMessageId = $InternetMessageIdValue
                }
                Invoke-EXRProcessAntiSPAMHeaders -Item $Item -rptObject $rptObject
                $rptCollection += $rptObject
            }
            $ivItemView.Offset += $fiItems.Items.Count
                    
            
        }while ($fiItems.MoreAvailable -eq $true)  
        return $rptCollection
    }
}

function Invoke-EXRProcessAntiSPAMHeaders {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $Item,
        [Parameter(Position = 3, Mandatory = $false)]
        [psobject]
        $rptObject,

        [Parameter(Position = 4, Mandatory = $false)]
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
                    Add-Member -InputObject $rptObject -NotePropertyName "SPF" -NotePropertyValue $SPF -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "DKIM" -NotePropertyValue $DKIM  -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "DMARC" -NotePropertyValue $DMARC  -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "CompAuth" -NotePropertyValue $CompAuth  -Force
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
                    Add-Member -InputObject $rptObject -NotePropertyName "Original-SPF" -NotePropertyValue $SPF -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "Original-DKIM" -NotePropertyValue $DKIM  -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "Original-DMARC" -NotePropertyValue $DMARC  -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "Original-CompAuth" -NotePropertyValue $CompAuth  -Force
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
                    Add-Member -InputObject $rptObject -NotePropertyName "PCL" -NotePropertyValue $PCL  -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "BCL" -NotePropertyValue $BCL  -Force
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
                    Add-Member -InputObject $rptObject -NotePropertyName "CTRY" -NotePropertyValue $CTRY  -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "SFV" -NotePropertyValue $SFV  -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "SRV" -NotePropertyValue $SRV  -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "PTR" -NotePropertyValue $PTR  -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "CIP" -NotePropertyValue $CIP  -Force
                    Add-Member -InputObject $rptObject -NotePropertyName "IPV" -NotePropertyValue $IPV  -Force
                }
                
                if ($Item.IndexedInternetMessageHeaders.ContainsKey("X-MS-Exchange-Organization-SCL")){
                    Add-Member -InputObject $rptObject -NotePropertyName "SCL" -NotePropertyValue $Item.IndexedInternetMessageHeaders["X-MS-Exchange-Organization-SCL"]  -Force 
                }
                if ($Item.IndexedInternetMessageHeaders.ContainsKey("X-CustomSpam")){
                    Add-Member -InputObject $rptObject -NotePropertyName "ASF" -NotePropertyValue $Item.IndexedInternetMessageHeaders["X-CustomSpam"]  -Force 
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
                    Invoke-EXRProcessAntiSPAMHeaders -Item $Item -noForward -rptObject $rptObject
                }

            }  

            
               
		
    }
}



function Get-UserFromGuid {
    param(
        [Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $false)] [string]$GUID

    )
    Process {
 
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $Script:GraphToken.access_token));
        $URL = "https://graph.microsoft.com/v1.0/users('" + $GUID + "')"
        $ClientResult = $HttpClient.GetAsync([Uri]$URL)
        return ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result
       
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
        $AccessToken,

        [Parameter(Position = 2, Mandatory = $true)]
        [string]
        $ResourceURL
    )
    process {
        Add-Type -AssemblyName System.Web
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $ClientId = $AccessToken.clientid
        # $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AccessToken.redirectUrl)
        $redirectUrl = $AccessToken.redirectUrl
        $RefreshToken = (ConvertFrom-SecureStringCustom -SecureToken $AccessToken.refresh_token)
        $AuthorizationPostRequest = "client_id=$ClientId&refresh_token=$RefreshToken&grant_type=refresh_token&redirect_uri=$redirectUrl"
        if ($ResourceURL) {
            $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&refresh_token=$RefreshToken&grant_type=refresh_token&redirect_uri=$redirectUrl"            
        }
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
$Script:Token = $null
$Script:GraphToken = $null
$Script:MaxCount = 0
$Script:UseMaxCount = $false
$Script:MaxCountExceeded = $false