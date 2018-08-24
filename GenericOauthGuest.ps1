function Test-GuestAuth  {
     param( 
		[Parameter(Position=0, Mandatory=$true)] [string]$UserName,
		[Parameter(Position=1, Mandatory=$true)] [string]$TargetDomain,
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$ClientId,
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$redirectUrl      

    )  
  Begin
 {
        $Resource = "graph.microsoft.com"    
        if([String]::IsNullOrEmpty($ClientId)){
            $ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
		}
		if([String]::IsNullOrEmpty($redirectUrl)){
			$redirectUrl = "urn:ietf:wg:oauth:2.0:oob"  
		}  
		$Script:Token = Get-EXRGuestAccessToken -MailboxName $UserName -ClientId $ClientId -redirectUrl $redirectUrl  -ResourceURL $Resource -Prompt $Prompt -CacheCredentials -TargetDomain $TargetDomain
		$headers = @{}
    	$headers.Add('Authorization',("Bearer " + (ConvertFrom-SecureStringCustom -SecureToken $Script:Token.access_token)))
		$Response = (Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/me/memberOf" -Headers $headers) 
		$JsonResponse = ConvertFrom-Json $Response.Content
		foreach($Group in $JsonResponse.value){
				Write-Output $Group
				if($Group.groupTypes -eq "Unified"){
					$cnvResponse = (Invoke-WebRequest -Uri ("https://graph.microsoft.com/v1.0/groups/" + $Group.id + "/conversations?Top=2") -Headers $headers) 
					$cnvJsonResponse = ConvertFrom-Json $cnvResponse.Content
					foreach($cnv in $cnvJsonResponse.value){
						Write-Output $cnv
					}	
				}		
			
		}

 }
}


function Invoke-ValidateToken{
    	param (
		
	)
	begin{
        $MailboxName = $Script:Token.mailbox
        $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
        $expiry = $minTime.AddSeconds($Script:Token.expires_on)
        if ($expiry -le [DateTime]::Now.ToUniversalTime().AddMinutes(10))
        {
            if ([bool]($Script:Token.PSobject.Properties.name -match "refresh_token"))
            {
                    write-host "Refresh Token"
                    $Script:Token = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $Script:Token
            }
            else
            {
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


function Get-EXRGuestAccessToken
{
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
		$CacheCredentials,
				
		[Parameter(Position = 8, Mandatory = $false)]
		[string]
		$TargetDomain
		
	)
	Begin
	{
		Add-Type -AssemblyName System.Web
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		if ([String]::IsNullOrEmpty($ClientId))
		{
			$ReturnToken = Get-ProfiledToken -MailboxName $MailboxName
			if($ReturnToken -eq $null){
				Write-Error ("No Access Token for " + $MailboxName)
			}
			else{				
				return $ReturnToken
			}
		}
		else{
			if ([String]::IsNullOrEmpty(($ClientSecret)))
			{
				$ClientSecret = $AppSetting.ClientSecret
			}
			if ([String]::IsNullOrEmpty($redirectUrl))
			{
				$redirectUrl = [System.Web.HttpUtility]::UrlEncode("urn:ietf:wg:oauth:2.0:oob")
			}
			else
			{
				$redirectUrl = [System.Web.HttpUtility]::UrlEncode($redirectUrl)
			}
			if ([String]::IsNullOrEmpty($ResourceURL))
			{
				$ResourceURL = $AppSetting.ResourceURL
			}
			if ([String]::IsNullOrEmpty($Prompt))
			{
				$Prompt = "refresh_session"
			}
			$RequestURL = "https://login.windows.net/{0}/.well-known/openid-configuration" -f $TargetDomain
            $Response = Invoke-WebRequest -Uri  $RequestURL
            $JsonResponse = ConvertFrom-Json  $Response.Content
            
			$phase1Url = $JsonResponse.authorization_endpoint + "?resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&response_type=code&redirect_uri=$redirectUrl&prompt=$Prompt&domain_hint=organizations"
			$Phase1auth = Show-OAuthWindow -Url $phase1Url
            $code = $Phase1auth["code"]
			$AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
			if (![String]::IsNullOrEmpty($ClientSecret))
			{
				$AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_secret=$ClientSecret&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
			}
			$content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
			$ClientReesult = $HttpClient.PostAsync([Uri]($JsonResponse.token_endpoint), $content)
			$JsonObject = ConvertFrom-Json -InputObject $ClientReesult.Result.Content.ReadAsStringAsync().Result
			if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token"))
			{
				$JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
			}
			if ([bool]($JsonObject.PSobject.Properties.name -match "access_token"))
			{
				$JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
			}
			if ([bool]($JsonObject.PSobject.Properties.name -match "id_token"))
			{
				$JsonObject.id_token = (Get-ProtectedToken -PlainToken $JsonObject.id_token)
			}
			Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
			Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
			Add-Member -InputObject $JsonObject -NotePropertyName mailbox -NotePropertyValue $MailboxName
			if ($Beta.IsPresent)
			{
				Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue $True
            }
			return $JsonObject
		}
	}
}
function Show-OAuthWindow
{
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
	$form.Add_Shown({ $form.Activate() })
	$form.ShowDialog() | Out-Null
	$queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
	$output = @{ }
	foreach ($key in $queryOutput.Keys)
	{
		$output["$key"] = $queryOutput[$key]
	}
	return $output
}

function Get-ProtectedToken
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$PlainToken
	)
	begin
	{
		$SecureEncryptedToken = Protect-String -String $PlainToken
		return, $SecureEncryptedToken
	}
}
function Get-HTTPClient
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName
	)
	process
	{
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
		if (!$HttpClient.DefaultRequestHeaders.Contains("X-AnchorMailbox"))
		{
			$HttpClient.DefaultRequestHeaders.Add("X-AnchorMailbox", $MailboxName);
		}
		$Header = New-Object System.Net.Http.Headers.ProductInfoHeaderValue("RestClient", "1.1")
		$HttpClient.DefaultRequestHeaders.UserAgent.Add($Header);
		return $HttpClient
	}
}

function Protect-String
{
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
	
	begin
	{
		Add-Type -AssemblyName System.Security -ErrorAction Stop
	}
	process
	{
		foreach ($item in $String)
		{
			$stringBytes = [Text.Encoding]::UTF8.GetBytes($item)
			$encodedBytes = [System.Security.Cryptography.ProtectedData]::Protect($stringBytes, $null, 'CurrentUser')
			[System.Convert]::ToBase64String($encodedBytes) | ConvertTo-SecureString -AsPlainText -Force
		}
	}
}

function Invoke-RefreshAccessToken
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[psobject]
		$AccessToken
	)
	process
	{
		Add-Type -AssemblyName System.Web
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$ClientId = $AccessToken.clientid
		# $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AccessToken.redirectUrl)
		$redirectUrl = $AccessToken.redirectUrl
		$RefreshToken = (ConvertFrom-SecureStringCustom -SecureToken $AccessToken.refresh_token)
		$AuthorizationPostRequest = "client_id=$ClientId&refresh_token=$RefreshToken&grant_type=refresh_token&redirect_uri=$redirectUrl"
		$content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
		$ClientResult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"), $content)
		if (!$ClientResult.Result.IsSuccessStatusCode)
		{
			Write-Output ("Error making REST POST " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
			Write-Output $ClientResult.Result
			if ($ClientResult.Content -ne $null)
			{
				Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);
			}
		}
		else
		{
			$JsonObject = ConvertFrom-Json -InputObject $ClientResult.Result.Content.ReadAsStringAsync().Result
			if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token"))
			{
				$JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
			}
			if ([bool]($JsonObject.PSobject.Properties.name -match "access_token"))
			{
				$JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
			}
			if ([bool]($JsonObject.PSobject.Properties.name -match "id_token"))
			{
				$JsonObject.id_token = (Get-ProtectedToken -PlainToken $JsonObject.id_token)
			}
			Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
			Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
			Add-Member -InputObject $JsonObject -NotePropertyName mailbox -NotePropertyValue $MailboxName
			if ($AccessToken.Beta)
			{
				Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
			}
		}
		return $JsonObject		
	}
}

function ConvertFrom-SecureStringCustom
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[System.Security.SecureString]
		$SecureToken
	)
	process
	{
		#$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureToken)
		$Token = Unprotect-String -String $SecureToken
		return, $Token
	}
}

function Unprotect-String
{
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
	
	begin
	{
		Add-Type -AssemblyName System.Security -ErrorAction Stop
	}
	process
	{
		foreach ($item in $String)
		{
			$cred = New-Object PSCredential("irrelevant", $item)
			$stringBytes = [System.Convert]::FromBase64String($cred.GetNetworkCredential().Password)
			$decodedBytes = [System.Security.Cryptography.ProtectedData]::Unprotect($stringBytes, $null, 'CurrentUser')
			[Text.Encoding]::UTF8.GetString($decodedBytes)
		}
	}
}
$Script:Token = $null