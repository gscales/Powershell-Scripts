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

function CovertBitValue($String) {  
	$numItempattern = '(?=\().*(?=bytes)'  
	$matchedItemsNumber = [regex]::matches($String, $numItempattern)   
	$Mb = [INT64]$matchedItemsNumber[0].Value.Replace("(", "").Replace(",", "")  
	return [math]::round($Mb / 1048576, 0)  
}  

function Get-UnReadMessageCount {
	param( 
		[Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
		[Parameter(Position = 2, Mandatory = $false)] [string]$url,
		[Parameter(Position = 3, Mandatory = $false)] [switch]$disableImpersonation,
		[Parameter(Position = 4, Mandatory = $false)] [switch]$basicAuth,
		[Parameter(Position = 5, Mandatory = $false)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position = 6, Mandatory = $false)] [switch]$useImpersonation,
		[Parameter(Position = 7, Mandatory = $false)] [switch]$AllItems,
		[Parameter(Position = 8, Mandatory = $false)] [switch]$Archive,
		[Parameter(Position = 9, Mandatory = $true)] [Int32]$Months
	)  
	Begin {
		$eval1 = "Last" + $Months + "MonthsTotal"
		$eval2 = "Last" + $Months + "MonthsUnread"
		$eval3 = "Last" + $Months + "MonthsSent"
		$eval4 = "Last" + $Months + "MonthsReplyToSender"
		$eval5 = "Last" + $Months + "MonthsReplyToAll"
		$eval6 = "Last" + $Months + "MonthForward"
		$reply = 0;
		$replyall = 0
		$forward = 0
		$rptObj = "" | select  MailboxName, Mailboxsize, LastLogon, LastLogonAccount, $eval1, $eval2, $eval4, $eval5, $eval6, LastMailRecieved, $eval3, LastMailSent  
		$rptObj.MailboxName = $MailboxName  
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
		if ($useImpersonation.IsPresent) {
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		}
		$AQSString1 = "System.Message.DateReceived:>" + [system.DateTime]::Now.AddMonths(-$Months).ToString("yyyy-MM-dd")   
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MailboxName)
		if($AllItems.IsPresent){
			$Inbox =  Get-AllItemsSearchFolder -MailboxName $MailboxName -service $service -Archive:$Archive.IsPresent
		}else{
			$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)  
		}	
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems, $MailboxName)     
		$SentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)    
		$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
		$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)  
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead)
		$PidTagLastVerbExecuted = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1081, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
		#Find Items Greater then 35 MB
		$sfItemSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, [system.DateTime]::Now.AddMonths(-$Months)) 
		$psPropset.Add($PidTagLastVerbExecuted)
		$ivItemView.PropertySet = $psPropset
		$MailboxStats = Get-MailboxStatistics $MailboxName  
		$ts = CovertBitValue($MailboxStats.TotalItemSize.ToString())  
		write-host ("Total Size : " + $MailboxStats.TotalItemSize) 
		$rptObj.MailboxSize = $ts  
		write-host ("Last Logon Time : " + $MailboxStats.LastLogonTime) 
		$rptObj.LastLogon = $MailboxStats.LastLogonTime  
		write-host ("Last Logon Account : " + $MailboxStats.LastLoggedOnUserAccount ) 
		$rptObj.LastLogonAccount = $MailboxStats.LastLoggedOnUserAccount  
		$fiItems = $null
		$unreadCount = 0
		$settc = $true
		do { 
			$fiItems = $Inbox.findItems($sfItemSearchFilter, $ivItemView)  
			if ($settc) {
				$rptObj.$eval1 = $fiItems.TotalCount  
				write-host ("Last " + $Months + " Months : " + $fiItems.TotalCount)
				if ($fiItems.TotalCount -gt 0) {  
					write-host ("Last Mail Recieved : " + $fiItems.Items[0].DateTimeReceived ) 
					$rptObj.LastMailRecieved = $fiItems.Items[0].DateTimeReceived  
				}		    
				$settc = $false
			}
			foreach ($Item in $fiItems.Items) {
				$unReadVal = $null
				if ($Item.TryGetProperty([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, [ref]$unReadVal)) {
					if (!$unReadVal) {
						$unreadCount++
					}
				} 
				$lastVerb = $null
				if ($Item.TryGetProperty($PidTagLastVerbExecuted, [ref]$lastVerb)) {
					switch ($lastVerb) {
						102 { $reply++ }
						103 { $replyall++ }
						104 { $forward++ }
					}
				} 
			}    
			$ivItemView.Offset += $fiItems.Items.Count    
		}while ($fiItems.MoreAvailable -eq $true) 

		write-host ("Last " + $Months + " Months Unread : " + $unreadCount ) 
		$rptObj.$eval2 = $unreadCount  
		$rptObj.$eval4 = $reply
		$rptObj.$eval5 = $replyall
		$rptObj.$eval6 = $forward
		$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)  
		$fiResults = $SentItems.findItems($sfItemSearchFilter, $ivItemView)  
		write-host ("Last " + $Months + " Months Sent : " + $fiResults.TotalCount  )
		$rptObj.$eval3 = $fiResults.TotalCount  
		if ($fiResults.TotalCount -gt 0) {  
			write-host ("Last Mail Sent Date : " + $fiResults.Items[0].DateTimeSent  )
			$rptObj.LastMailSent = $fiResults.Items[0].DateTimeSent  
		}  
		Write-Output $rptObj  
	}
}

function Get-AllItemsSearchFolder{
    param(
        [Parameter(Position = 1, Mandatory = $true)] [string]$MailboxName,
        [Parameter(Position = 2, Mandatory = $true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
		[Parameter(Position = 3, Mandatory = $false)] [switch]$Archive
    )
    Process{
        $folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)
		if($Archive.IsPresent){
			$folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRoot,$MailboxName)
		}
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